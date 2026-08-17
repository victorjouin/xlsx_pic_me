"""
xlsxtiles.manifest — pilote classeur et contrat de sortie.

`manifest.json` est ce que l'étage aval consomme : pour chaque image produite,
la plage A1 exacte qu'elle représente. C'est ce lien image <-> plage qui permet
de rendre au user le morceau d'Excel correspondant à l'information citée.
"""

from __future__ import annotations

import hashlib
import json
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .calibration import load_calibration
from .config import DEFAULT_PATCH_BUDGET, RENDER_DPI
from .encode import PNG_DATA_URI_PREFIX, png_to_base64
from .models import RenderedTile, WorkbookProfile
from .profiler import profile_workbook
from .refs import col_letters
from .render import render_sheet_tiles

__all__ = ["render_workbook", "sha256"]


def sha256(path: str | Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _tile_entry(t: RenderedTile, include_base64: bool) -> dict[str, Any]:
    """Une tuile telle que l'aval la consomme : l'image ET sa plage exacte.

    `sha256` porte sur les octets du PNG, donc sur ce que `png_base64` encode :
    l'intégrité reste vérifiable après le transport en base64.
    """
    entry: dict[str, Any] = {
        "index": t.plan.index,
        "range_a1": t.plan.range_a1,
        "row_start": t.plan.row_start, "row_end": t.plan.row_end,
        "col_start": t.plan.col_start, "col_end": t.plan.col_end,
        "header_rows": t.plan.header_rows,
        "png": Path(t.png_path).name,
        "width_px": t.width_px, "height_px": t.height_px,
        "sha256": sha256(t.png_path),
        "grid_verified": t.grid_verified,
        "grid_drift_px": t.grid_drift_px,
    }
    if include_base64:
        entry["png_base64"] = png_to_base64(t.png_path)
    return entry


def render_workbook(src: str | Path, outdir: str | Path, *,
                    patch_budget: int = DEFAULT_PATCH_BUDGET,
                    dpi: int = RENDER_DPI,
                    include_base64: bool = False,
                    profile: WorkbookProfile | None = None) -> dict[str, Any]:
    """Découpe tout le classeur en images et écrit `manifest.json` à côté.

    Toutes les feuilles ingérables sont rendues, y compris celles classées
    `image_indexing="skip"` : ce champ est transmis tel quel dans le manifeste
    et c'est à l'étage d'indexation, pas au découpage, de décider quoi
    vectoriser. Une image manquante ici ne peut plus être rendue à l'aval.

    `include_base64` ajoute `png_base64` à chaque tuile : l'image voyage alors
    dans le même document que sa plage A1, sans dépendre d'un stockage partagé.
    Le manifeste grossit d'environ 4/3 du poids total des PNG — voir
    `xlsxtiles.encode` pour les ordres de grandeur.
    """
    src, outdir = Path(src), Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    prof = profile or profile_workbook(src, calibration=load_calibration())

    sheets_out: list[dict[str, Any]] = []
    for name in prof.verdict.sheets_to_ingest:
        sp = prof.sheet(name)
        assert sp is not None
        tiles = render_sheet_tiles(src, prof, name, outdir,
                                   patch_budget=patch_budget, dpi=dpi)
        sheets_out.append({
            "sheet": name,
            "image_indexing": sp.image_indexing,
            "tiling_strategy": sp.tiling_strategy,
            "range_a1": (f"{col_letters(sp.true_min_col or 1)}{sp.true_min_row}"
                         f":{col_letters(sp.true_max_col or 1)}{sp.true_max_row}"),
            "n_tiles": len(tiles),
            "tiles": [_tile_entry(t, include_base64) for t in tiles],
        })

    manifest = {
        "source": src.name,
        "source_sha256": sha256(src),
        "generated_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "dpi": dpi,
        "patch_budget": patch_budget,
        "base64_included": include_base64,
        # préfixe à ajouter devant png_base64 pour obtenir une data URI
        # affichable telle quelle ; absent du champ pour laisser le choix
        "base64_data_uri_prefix": PNG_DATA_URI_PREFIX if include_base64 else None,
        "grid_reliable": prof.verdict.grid_reliable,
        "calibration": asdict(prof.calibration),
        "warnings": prof.verdict.warnings,
        "sheets_skipped": prof.verdict.sheets_skipped,
        "sheets": sheets_out,
    }
    (outdir / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest
