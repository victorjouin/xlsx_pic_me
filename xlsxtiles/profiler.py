"""
xlsxtiles.profiler — pré-scan d'un classeur avant tout rendu.

Assemble les briques de `scan` et `fonts` en un profil sérialisable qui pilote :
  - le périmètre d'ingestion (onglets visibles / masqués / techniques),
  - le routage d'exécution (Lambda / Fargate / Batch),
  - la stratégie de découpage en tuiles,
  - la fiabilité attendue de la grille pixel.

Dépendances : stdlib uniquement.

Usage :
    from xlsxtiles import profile_workbook
    profile = profile_workbook("budget2026.xlsx")
    print(profile.to_json())
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .config import (
    BATCH_MIN_CELLS, CACHED_VALUE_MIN_RATIO, DEEP_SCAN_MAX_MB,
    LAMBDA_MAX_CELLS, LAMBDA_MAX_MB, LAYOUT_MERGE_COUNT,
    MAX_TRACKED_ROW_HEIGHTS, SPARSE_FILL_RATIO,
)
from .fonts import check_fonts, fonts_used
from .models import GridCalibration, SheetProfile, Verdict, WorkbookProfile
from .ooxml import NS_MAIN, NS_REL_DOC, read_xml, rels_for
from .scan import attach_sheet_parts, scan_sheet_tail, scan_sheet_xml, style_map

__all__ = ["profile_workbook", "self_check"]


# --------------------------------------------------------------------------
# Verdicts
# --------------------------------------------------------------------------

def classify_sheet(sp: SheetProfile) -> None:
    """Pose `image_indexing` et `tiling_strategy` sur une feuille scannée.

    ATTENTION : `image_indexing` est une RECOMMANDATION transmise à l'étage
    d'indexation, pas un filtre de rendu. Le découpage produit l'image de
    toutes les feuilles ingérables — une image non produite ici ne peut plus
    l'être en aval.
    """
    sp.is_sparse = sp.deep_scanned and 0 < sp.fill_ratio < SPARSE_FILL_RATIO
    sp.is_layout_sheet = sp.n_merged_cells >= LAYOUT_MERGE_COUNT
    sp.is_pure_data = (sp.n_charts == 0 and sp.n_images == 0
                       and sp.n_conditional_formats == 0)

    # Une feuille purement numérique en grille régulière n'apporte pas grand
    # chose en embedding visuel : le markdown la décrit mieux et pour moins
    # cher. L'image reste produite, c'est l'embedding qui est discutable.
    if sp.n_charts or sp.n_images:
        sp.image_indexing = "full"
    elif sp.is_layout_sheet or sp.n_conditional_formats:
        sp.image_indexing = "full"
    elif sp.is_pure_data and sp.deep_scanned and sp.fill_ratio > 0.8:
        sp.image_indexing = "skip"
    else:
        sp.image_indexing = "lazy"

    if sp.tables:
        sp.tiling_strategy = "declared_tables"
    elif sp.row_breaks or sp.col_breaks:
        sp.tiling_strategy = "page_breaks"
    elif sp.is_layout_sheet or (sp.deep_scanned and sp.true_n_rows <= 60
                                and sp.true_n_cols <= 20):
        sp.tiling_strategy = "whole_sheet"
    elif sp.is_sparse:
        sp.tiling_strategy = "geometric"
    else:
        sp.tiling_strategy = "xy_cut"


def build_verdict(profile_bits: dict[str, Any], sheets: list[SheetProfile],
                  fonts: list[str], file_mb: float,
                  cal: GridCalibration) -> Verdict:
    warnings: list[str] = []
    missing, font_warnings = check_fonts(fonts)
    warnings.extend(font_warnings)

    total_cells = sum(s.effective_cells for s in sheets)
    any_shallow = any(not s.deep_scanned for s in sheets)
    if file_mb > LAMBDA_MAX_MB * 3 or total_cells > BATCH_MIN_CELLS:
        target = "batch"
    elif file_mb > LAMBDA_MAX_MB or total_cells > LAMBDA_MAX_CELLS or any_shallow:
        # une feuille non scannée est une inconnue : on ne la confie pas à Lambda
        target = "fargate"
    else:
        target = "lambda"

    n_f = sum(s.n_formulas for s in sheets)
    n_cached = sum(s.n_formulas_with_cached_value for s in sheets)
    needs_recalc = n_f > 0 and (n_cached / n_f) < CACHED_VALUE_MIN_RATIO
    if needs_recalc:
        warnings.append(
            f"{n_f - n_cached}/{n_f} formules sans valeur en cache : passe de "
            "recalcul LibreOffice obligatoire avant rendu, sinon cellules vides")

    to_ingest, skipped = [], {}
    for s in sheets:
        if s.state == "veryHidden":
            skipped[s.name] = "veryHidden (feuille technique)"
        elif s.state == "hidden":
            skipped[s.name] = "masquée"
        elif s.n_cells_with_content == 0 and s.deep_scanned:
            skipped[s.name] = "vide"
        else:
            to_ingest.append(s.name)

    if not cal.calibrated:
        warnings.append(
            "grille pixel non calibrée (source=%s) : surlignage de cellule désactivé, "
            "lancer la feuille sonde au build" % cal.source)
    if missing:
        warnings.append(
            f"polices absentes du conteneur ({', '.join(missing)}) : "
            "substitution LibreOffice -> dérive de grille, surlignage désactivé")
    if profile_bits.get("has_external_links"):
        warnings.append("liens externes : valeurs potentiellement obsolètes")
    if profile_bits.get("has_pivot_cache"):
        warnings.append("tableaux croisés dynamiques : rendu figé sur le dernier "
                        "rafraîchissement connu")
    n_skip = sum(1 for s in sheets if s.image_indexing == "skip")
    if n_skip:
        warnings.append(
            f"{n_skip} feuille(s) recommandée(s) en image_indexing=skip : "
            "données pures en grille régulière, l'embedding image y apporte "
            "peu — les images sont produites malgré tout")
    for s in sheets:
        if not s.deep_scanned:
            warnings.append(f"feuille '{s.name}' scannée superficiellement "
                            f"({s.xml_size_mb:.1f} Mo) : bornes et densité inconnues")
        if s.is_sparse:
            warnings.append(f"feuille '{s.name}' éparse (remplissage "
                            f"{s.fill_ratio:.1%}) : durcir les seuils de coupe XY")
        if s.row_heights_truncated:
            warnings.append(f"feuille '{s.name}' : plus de {MAX_TRACKED_ROW_HEIGHTS} "
                            "hauteurs de ligne personnalisées, les suivantes sont "
                            "ignorées — cadrage vertical approximatif")

    return Verdict(
        compute_target=target,
        needs_recalc=needs_recalc,
        grid_reliable=bool(cal.calibrated) and not missing and not font_warnings,
        missing_fonts=missing,
        sheets_to_ingest=to_ingest,
        sheets_skipped=skipped,
        warnings=warnings,
    )


# --------------------------------------------------------------------------
# Point d'entrée
# --------------------------------------------------------------------------

def profile_workbook(path: str | Path, *,
                     deep_scan_max_mb: float = DEEP_SCAN_MAX_MB,
                     calibration: GridCalibration | None = None,
                     include_hidden: bool = False) -> WorkbookProfile:
    """Profile un classeur .xlsx/.xlsm sans le charger en mémoire.

    Args:
        path: chemin du classeur.
        deep_scan_max_mb: au-delà de cette taille de XML décompressé, la feuille
            n'est scannée qu'en surface (métadonnées de tête, pas de comptage).
        calibration: calibration pixel connue ; à défaut, valeurs par défaut
            marquées non calibrées (surlignage cellule interdit). Voir
            `xlsxtiles.calibration.load_calibration()`.
        include_hidden: si True, les feuilles masquées passent en ingestion.

    Returns:
        WorkbookProfile sérialisable en JSON.
    """
    path = Path(path)
    file_mb = path.stat().st_size / 1e6
    cal = calibration or GridCalibration()

    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        sizes = {i.filename: i.file_size / 1e6 for i in z.infolist()}

        wb_root = read_xml(z, "xl/workbook.xml")
        if wb_root is None:
            raise ValueError(f"{path} : xl/workbook.xml introuvable "
                             "(fichier .xls ancien format ou archive corrompue ?)")
        wb_rels = rels_for(z, "xl/workbook.xml")

        # --- onglets ---
        sheets: list[SheetProfile] = []
        sheets_el = wb_root.find(f"{{{NS_MAIN}}}sheets")
        for i, sh in enumerate(list(sheets_el) if sheets_el is not None else []):
            rid = sh.get(f"{{{NS_REL_DOC}}}id")
            part = wb_rels.get(rid or "", "")
            if part not in names:
                continue
            sheets.append(SheetProfile(
                name=sh.get("name", f"Sheet{i+1}"),
                index=i,
                state=sh.get("state", "visible"),
                part=part,
                xml_size_mb=round(sizes.get(part, 0.0), 3),
                deep_scanned=False,
            ))

        known_names = {s.name for s in sheets}

        # --- noms définis (dont zones et titres d'impression) ---
        defined_names: dict[str, str] = {}
        dn_el = wb_root.find(f"{{{NS_MAIN}}}definedNames")
        for dn in (list(dn_el) if dn_el is not None else []):
            nm = dn.get("name", "")
            val = (dn.text or "").strip()
            lsid = dn.get("localSheetId")
            target = sheets[int(lsid)] if lsid is not None and int(lsid) < len(sheets) else None
            if nm == "_xlnm.Print_Area" and target:
                target.print_area = val
            elif nm == "_xlnm.Print_Titles" and target:
                for part_ref in val.split(","):
                    if "$" in part_ref and re.search(r"\$\d+", part_ref):
                        target.print_title_rows = part_ref.strip()
                    elif "$" in part_ref:
                        target.print_title_cols = part_ref.strip()
            else:
                defined_names[nm] = val

        # --- scan par feuille ---
        styles = style_map(z)
        for sp in sheets:
            deep = sp.xml_size_mb <= deep_scan_max_mb
            try:
                scan_sheet_xml(z, sp, known_names, deep, styles)
                if not deep:
                    scan_sheet_tail(z, sp)
            except ET.ParseError as exc:
                sp.deep_scanned = False
                defined_names.setdefault("__parse_error__", str(exc))
            attach_sheet_parts(z, sp)
            classify_sheet(sp)

        fonts = fonts_used(z)
        bits = {
            "has_external_links": any(n.startswith("xl/externalLinks/") for n in names),
            "has_pivot_cache": any(n.startswith("xl/pivotCache/") for n in names),
            "has_vba": "xl/vbaProject.bin" in names,
        }

    edges = sorted({(s.name, dep) for s in sheets for dep in s.depends_on_sheets})
    verdict = build_verdict(bits, sheets, fonts, file_mb, cal)
    if include_hidden:
        for name, reason in list(verdict.sheets_skipped.items()):
            if reason == "masquée":
                verdict.sheets_to_ingest.append(name)
                del verdict.sheets_skipped[name]

    return WorkbookProfile(
        path=str(path),
        file_size_mb=round(file_mb, 3),
        n_sheets=len(sheets),
        sheets=sheets,
        fonts_used=fonts,
        defined_names=defined_names,
        has_external_links=bits["has_external_links"],
        has_pivot_cache=bits["has_pivot_cache"],
        has_vba=bits["has_vba"],
        total_cells_with_content=sum(s.n_cells_with_content for s in sheets),
        estimated_total_cells=sum(s.effective_cells for s in sheets),
        total_charts=sum(s.n_charts for s in sheets),
        total_images=sum(s.n_images for s in sheets),
        sheet_dependency_edges=[list(e) for e in edges],  # JSON-friendly
        calibration=cal,
        verdict=verdict,
    )


def self_check() -> None:
    """Invariants de conversion. À exécuter au démarrage du conteneur."""
    from .numfmt import classify_number_format
    from .refs import col_to_idx, excel_height_to_px, excel_width_to_px, parse_cell_ref

    cal = GridCalibration()
    assert int(excel_width_to_px(8.43, cal)) == 64, "largeur défaut != 64 px"
    assert int(excel_width_to_px(None, cal)) == 64, "fallback largeur défaut"
    assert int(excel_height_to_px(15.0, cal)) == 20, "hauteur défaut != 20 px"
    assert int(excel_height_to_px(None, cal)) == 20, "fallback hauteur défaut"
    assert col_to_idx("A") == 1 and col_to_idx("Z") == 26 and col_to_idx("AA") == 27
    assert parse_cell_ref("AB12") == (28, 12)

    cases = {
        "dd/mm/yyyy": "date",
        '"€"#,##0.00': "currency",          # symbole entre guillemets
        "[$€-40C]#,##0.00": "currency",     # bloc locale
        "#,##0.00;[Red](#,##0.00)": "number",  # [Red] contient un 'd'
        "0.00%": "percent",
        "#,##0.00": "number",
        "[h]:mm:ss": "time",                # durée écoulée, crochets conservés
        "m/d/yy h:mm": "datetime",
        "@": "text",
        "General": "general",
    }
    for code, expected in cases.items():
        got = classify_number_format(164, code)
        assert got == expected, f"{code!r} -> {got}, attendu {expected}"
    assert classify_number_format(14, None) == "date"      # intégré mm-dd-yy
    assert classify_number_format(9, None) == "percent"    # intégré 0%
