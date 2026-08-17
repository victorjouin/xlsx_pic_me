"""
xlsxtiles.scan — parcours du XML d'une feuille et des parties qui lui sont liées.

Tout ce qui remplit un `SheetProfile` à partir de l'archive : streaming du
corps de la feuille, relecture de sa queue en scan superficiel, tableaux
déclarés, dessins et leurs ancres, formats de colonne.
"""

from __future__ import annotations

import re
import zipfile
from xml.etree import ElementTree as ET

from .config import MAX_TRACKED_ROW_HEIGHTS, TEXT_DOMINANCE
from .models import AnchorBox, SheetProfile, TableInfo
from .numfmt import BUILTIN_NUMFMT, classify_number_format
from .ooxml import NS_MAIN, NS_XDR, local, read_xml, rels_for
from .refs import parse_cell_ref

__all__ = ["scan_sheet_xml", "scan_sheet_tail", "attach_sheet_parts", "style_map"]

# Attribut `t` de <c> -> nature de la donnée. Absent = numérique.
TYPE_KIND = {"s": "text", "inlineStr": "text", "str": "text",
             "b": "bool", "e": "error", "d": "date", "n": "number"}

# Références inter-onglets dans une formule : 'Mon Onglet'!A1 ou Onglet!A1
_SHEET_REF_RE = re.compile(r"(?:'((?:[^']|'')+)'|([A-Za-z_\\][A-Za-z0-9_.\\]*))!")

EMU_PER_PX = 9525.0          # 914400 EMU/pouce / 96 px/pouce
DEFAULT_COL_PX = 64.0        # 8.43 caractères Calibri 11
DEFAULT_ROW_PX = 20.0        # 15 pt
_TAIL_BYTES = 512 * 1024


# --------------------------------------------------------------------------
# Corps de la feuille
# --------------------------------------------------------------------------

def scan_sheet_xml(z: zipfile.ZipFile, sp: SheetProfile,
                   known_sheet_names: set[str], deep: bool,
                   styles: tuple[list[int], dict[int, str]] = ([], {})) -> None:
    """Parcours en streaming du XML de la feuille.

    deep=False : on s'arrête dès le début de <sheetData> sans compter les
    cellules (protection mémoire/temps sur les feuilles massives). Les
    métadonnées de tête (cols, pane, sheetFormatPr) sont toujours lues ; la
    queue se récupère ensuite via scan_sheet_tail().
    """
    formula_refs: set[str] = set()
    min_r = min_c = None
    max_r = max_c = 0
    break_mode: str | None = None   # "row" | "col" — contexte parent des <brk>
    col_styles: dict[int, dict[int, int]] = {}   # colonne -> {xf index: compte}
    col_types: dict[int, dict[str, int]] = {}    # colonne -> {type: compte}
    cur_style = 0
    cur_type = "n"
    cur_row = 0

    with z.open(sp.part) as fh:
        ctx = ET.iterparse(fh, events=("start", "end"))
        sheet_data_el: ET.Element | None = None
        cur_cell_ref: str | None = None
        cur_has_formula = False
        cur_has_value = False

        for event, el in ctx:
            tag = local(el.tag)

            if event == "start":
                if tag == "c":
                    cur_cell_ref = el.get("r")
                    cur_has_formula = False
                    cur_has_value = False
                    cur_style = int(el.get("s", 0))
                    cur_type = el.get("t") or "n"
                elif tag == "row":
                    cur_row = int(el.get("r", 0) or 0)
                elif tag == "f":
                    cur_has_formula = True
                elif tag == "rowBreaks":
                    break_mode = "row"
                elif tag == "colBreaks":
                    break_mode = "col"
                elif tag == "sheetData":
                    if not deep:
                        # scan superficiel : cols/pane/sheetFormatPr sont déjà
                        # lus (ils précèdent sheetData). On abandonne le corps
                        # ici, pas à sa fermeture, sinon on parse tout quand même.
                        break
                    sheet_data_el = el
                continue

            # --- events "end" ---
            if tag in ("v", "is"):
                # PIÈGE : un <v></v> vide est écrit par certaines libs (openpyxl)
                # pour une formule non évaluée. Ce n'est PAS une valeur en cache.
                cur_has_value = bool("".join(el.itertext()).strip())

            elif tag == "dimension":
                sp.declared_dimension = el.get("ref")

            elif tag == "sheetFormatPr":
                dcw = el.get("defaultColWidth")
                drh = el.get("defaultRowHeight")
                sp.default_col_width = float(dcw) if dcw else None
                sp.default_row_height = float(drh) if drh else None

            elif tag == "col":
                cmin, cmax = el.get("min"), el.get("max")
                w = el.get("width")
                if cmin and cmax:
                    if w and el.get("customWidth") == "1":
                        sp.custom_col_widths[f"{cmin}-{cmax}"] = float(w)
                    if el.get("hidden") == "1":
                        sp.hidden_cols.extend(range(int(cmin), min(int(cmax), 16384) + 1))
                    lvl = el.get("outlineLevel")
                    if lvl:
                        sp.max_outline_level = max(sp.max_outline_level, int(lvl))

            elif tag == "pane":
                # volets figés : xSplit/ySplit = nb de colonnes/lignes figées
                if el.get("state") in ("frozen", "frozenSplit"):
                    sp.freeze_cols = int(float(el.get("xSplit", 0)))
                    sp.freeze_rows = int(float(el.get("ySplit", 0)))

            elif tag == "autoFilter":
                sp.auto_filter_ref = el.get("ref")

            elif tag == "mergeCells":
                sp.n_merged_cells = int(el.get("count", 0))

            elif tag == "conditionalFormatting":
                sp.n_conditional_formats += 1

            elif tag == "dataValidation":
                sp.n_data_validations += 1

            elif tag == "hyperlink":
                sp.n_hyperlinks += 1

            elif tag == "brk":
                # L'attribut `max` ne permet PAS de distinguer : il vaut 16383
                # (dernière colonne) pour un saut de ligne et 1048575 (dernière
                # ligne) pour un saut de colonne. Seul le parent fait foi.
                bid = el.get("id")
                if bid and el.get("man") == "1":
                    if break_mode == "row":
                        sp.row_breaks.append(int(bid))
                    elif break_mode == "col":
                        sp.col_breaks.append(int(bid))

            elif tag in ("rowBreaks", "colBreaks"):
                break_mode = None

            elif tag == "f" and deep:
                if el.text:
                    for m in _SHEET_REF_RE.finditer(el.text):
                        nm = (m.group(1) or m.group(2) or "").replace("''", "'")
                        if nm in known_sheet_names:
                            formula_refs.add(nm)

            elif tag == "c" and deep:
                if cur_has_value or cur_has_formula:
                    sp.n_cells_with_content += 1
                    if cur_has_formula:
                        sp.n_formulas += 1
                        if cur_has_value:
                            sp.n_formulas_with_cached_value += 1
                    pos = parse_cell_ref(cur_cell_ref or "")
                    if pos:
                        c, r = pos
                        if cur_row > 1:   # l'en-tête a souvent un autre style
                            col_styles.setdefault(c, {})
                            col_styles[c][cur_style] = col_styles[c].get(cur_style, 0) + 1
                            kind = TYPE_KIND.get(cur_type, "number")
                            col_types.setdefault(c, {})
                            col_types[c][kind] = col_types[c].get(kind, 0) + 1
                        min_c = c if min_c is None else min(min_c, c)
                        min_r = r if min_r is None else min(min_r, r)
                        max_c, max_r = max(max_c, c), max(max_r, r)
                cur_cell_ref = None

            elif tag == "row":
                if el.get("hidden") == "1":
                    sp.hidden_rows.append(int(el.get("r", 0)))
                # Hauteur personnalisée : sans elle la grille pixel prédit une
                # hauteur de tuile fausse, la taille de page l'est aussi, et
                # LibreOffice repagine selon sa propre logique — le mapping
                # page PDF -> tuile saute. C'est LE mode d'échec du cadrage.
                ht = el.get("ht")
                if ht and len(sp.custom_row_heights) < MAX_TRACKED_ROW_HEIGHTS:
                    try:
                        sp.custom_row_heights[el.get("r", "0")] = float(ht)
                    except ValueError:
                        pass
                elif ht:
                    sp.row_heights_truncated = True
                lvl = el.get("outlineLevel")
                if lvl:
                    sp.max_outline_level = max(sp.max_outline_level, int(lvl))
                el.clear()
                # détache la ligne du parent, sinon <sheetData> accumule des
                # noeuds vides et la mémoire croît malgré le clear()
                if sheet_data_el is not None:
                    try:
                        sheet_data_el.remove(el)
                    except ValueError:
                        pass
                continue

    if deep and min_r is not None:
        sp.true_min_row, sp.true_max_row = min_r, max_r
        sp.true_min_col, sp.true_max_col = min_c, max_c
        area = sp.true_n_rows * sp.true_n_cols
        sp.fill_ratio = (sp.n_cells_with_content / area) if area else 0.0

    sp.depends_on_sheets = sorted(formula_refs - {sp.name})
    sp.deep_scanned = deep
    if deep and col_styles:
        _resolve_column_formats(sp, col_styles, col_types, *styles)


def scan_sheet_tail(z: zipfile.ZipFile, sp: SheetProfile) -> None:
    """Récupère les blocs postérieurs à <sheetData> sans parser le corps.

    mergeCells, conditionalFormatting, autoFilter, rowBreaks/colBreaks se
    trouvent après les données. En scan superficiel on les perdrait ; on relit
    donc uniquement la queue du flux décompressé (peu coûteux : c'est le parsing
    XML qui coûte, pas la décompression).
    """
    tail = b""
    with z.open(sp.part) as fh:
        while True:
            chunk = fh.read(1 << 20)
            if not chunk:
                break
            tail = (tail + chunk)[-_TAIL_BYTES:]
    txt = tail.decode("utf-8", errors="ignore")
    txt = txt[txt.rfind("</sheetData>") + 1:] if "</sheetData>" in txt else txt

    m = re.search(r"<mergeCells[^>]*count=\"(\d+)\"", txt)
    if m:
        sp.n_merged_cells = int(m.group(1))
    m = re.search(r"<autoFilter[^>]*ref=\"([^\"]+)\"", txt)
    if m:
        sp.auto_filter_ref = m.group(1)
    sp.n_conditional_formats = len(re.findall(r"<conditionalFormatting", txt))
    sp.n_data_validations = len(re.findall(r"<dataValidation[ >]", txt))
    sp.n_hyperlinks = len(re.findall(r"<hyperlink[ >]", txt))
    for block, target in (("rowBreaks", sp.row_breaks), ("colBreaks", sp.col_breaks)):
        bm = re.search(rf"<{block}[^>]*>(.*?)</{block}>", txt, re.S)
        if bm:
            target.extend(int(i) for i in
                          re.findall(r'<brk id="(\d+)"[^>]*man="1"', bm.group(1)))


# --------------------------------------------------------------------------
# Parties liées : tableaux, dessins, pivots
# --------------------------------------------------------------------------

def attach_sheet_parts(z: zipfile.ZipFile, sp: SheetProfile) -> None:
    """Tableaux déclarés, dessins (graphiques/images) et leurs ancres."""
    rels = rels_for(z, sp.part)
    for target in rels.values():
        if "/tables/" in target:
            root = read_xml(z, target)
            if root is not None:
                af = root.find(f"{{{NS_MAIN}}}autoFilter")
                if af is not None and af.get("ref") and not sp.auto_filter_ref:
                    # les tables Excel portent leur filtre dans tableN.xml,
                    # pas dans le XML de la feuille
                    sp.auto_filter_ref = af.get("ref")
                sp.tables.append(TableInfo(
                    name=root.get("displayName") or root.get("name") or "",
                    ref=root.get("ref", ""),
                    header_row_count=int(root.get("headerRowCount", 1)),
                    totals_row_count=int(root.get("totalsRowCount", 0)),
                ))
        elif "/drawings/drawing" in target:
            _scan_drawing(z, sp, target)
        elif "/pivotTables/" in target:
            sp.has_pivot_table = True

    sp.n_charts = sum(1 for a in sp.anchors if a.kind == "chart")
    sp.n_images = sum(1 for a in sp.anchors if a.kind == "image")


def _extent_to_cells(anchor: ET.Element, frm: tuple[int, int]) -> tuple[int, int]:
    """oneCellAnchor : convertit <xdr:ext cx cy> (EMU) en cellule de fin.

    Approximation à largeur/hauteur par défaut : suffisant pour marquer une
    zone incoupable, pas pour un cadrage au pixel près. Le cadrage exact se
    fait au rendu, à partir de la grille calibrée.
    """
    ext = anchor.find(f"{{{NS_XDR}}}ext")
    if ext is None:
        return frm
    try:
        w_px = float(ext.get("cx", 0)) / EMU_PER_PX
        h_px = float(ext.get("cy", 0)) / EMU_PER_PX
    except ValueError:
        return frm
    return (frm[0] + max(1, int(w_px / DEFAULT_COL_PX + 0.999)) - 1,
            frm[1] + max(1, int(h_px / DEFAULT_ROW_PX + 0.999)) - 1)


def _scan_drawing(z: zipfile.ZipFile, sp: SheetProfile, drawing_part: str) -> None:
    root = read_xml(z, drawing_part)
    if root is None:
        return
    drels = rels_for(z, drawing_part)

    for anchor in root:
        atag = local(anchor.tag)
        if atag not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
            continue

        def _pos(node_name: str) -> tuple[int, int] | None:
            node = anchor.find(f"{{{NS_XDR}}}{node_name}")
            if node is None:
                return None
            col = node.find(f"{{{NS_XDR}}}col")
            row = node.find(f"{{{NS_XDR}}}row")
            if col is None or row is None:
                return None
            return int(col.text or 0) + 1, int(row.text or 0) + 1

        frm = _pos("from")
        if frm is None:
            continue
        to = _pos("to")
        if to is None:
            # oneCellAnchor : pas de <to>, l'étendue est en EMU dans <ext>
            to = _extent_to_cells(anchor, frm)

        kind = "shape"
        blob = ET.tostring(anchor, encoding="unicode")
        if "graphicFrame" in blob:
            # un graphicFrame peut être un chart, un diagramme ou un tableau
            kind = "chart"
            for rid in re.findall(r'r:id="([^"]+)"', blob):
                tgt = drels.get(rid, "")
                if "/charts/" in tgt:
                    kind = "chart"
                    break
        elif "<xdr:pic" in blob or ":pic " in blob or "<pic:" in blob:
            kind = "image"

        sp.anchors.append(AnchorBox(kind, frm[0], frm[1], to[0], to[1]))


# --------------------------------------------------------------------------
# Styles et formats de colonne
# --------------------------------------------------------------------------

def style_map(z: zipfile.ZipFile) -> tuple[list[int], dict[int, str]]:
    """-> (index xf -> numFmtId, numFmtId -> code)"""
    root = read_xml(z, "xl/styles.xml")
    if root is None:
        return [], {}
    codes = {int(n.get("numFmtId")): n.get("formatCode", "")
             for n in root.iter(f"{{{NS_MAIN}}}numFmt") if n.get("numFmtId")}
    xfs: list[int] = []
    cell_xfs = root.find(f"{{{NS_MAIN}}}cellXfs")
    if cell_xfs is not None:
        xfs = [int(xf.get("numFmtId", 0)) for xf in cell_xfs]
    return xfs, codes


def _resolve_column_formats(sp: SheetProfile, col_styles: dict[int, dict[int, int]],
                            col_types: dict[int, dict[str, int]],
                            xfs: list[int], codes: dict[int, str]) -> None:
    """Format dominant par colonne, croisé avec le type réel des cellules.

    Un format de nombre posé sur une colonne de texte est INERTE : Excel
    l'ignore à l'affichage, mais il reste dans styles.xml. Sans ce croisement,
    une colonne 'Product' héritant du format comptable appliqué en bloc sur la
    plage ressort en `currency`, et l'extracteur markdown tente de formater
    'January' en dollars.
    """
    for col, hist in col_styles.items():
        if not hist:
            continue
        types = col_types.get(col, {})
        total = sum(types.values())
        text_ratio = types.get("text", 0) / total if total else 0.0
        dominant_type = (max(types.items(), key=lambda kv: kv[1])[0]
                         if types else "number")
        sp.column_types[str(col)] = dominant_type

        if text_ratio >= TEXT_DOMINANCE:
            sp.column_formats[str(col)] = "text"
            continue

        xf_idx = max(hist.items(), key=lambda kv: kv[1])[0]
        fmt_id = xfs[xf_idx] if 0 <= xf_idx < len(xfs) else 0
        code = codes.get(fmt_id) or BUILTIN_NUMFMT.get(fmt_id)
        kind = classify_number_format(fmt_id, codes.get(fmt_id))
        if kind != "general":
            sp.column_formats[str(col)] = kind
            if code:
                sp.column_format_codes[str(col)] = code
