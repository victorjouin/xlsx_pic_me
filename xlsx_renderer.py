"""
Compatibilité — le moteur de rendu vit désormais dans le package `xlsxtiles`.

Ce module ne fait que ré-exporter l'API publique pour ne pas casser un import
existant en amont de la pipeline. Le code réel est réparti dans :

    xlsxtiles/render.py       mise en scène et rasterisation
    xlsxtiles/manifest.py     pilote classeur + manifest.json
    xlsxtiles/plan.py         largeurs d'affichage + plan de tuiles
    xlsxtiles/grid.py         grille prédite + mesure du PDF
    xlsxtiles/calibration.py  feuille sonde + cache
    xlsxtiles/soffice.py      pilotage de LibreOffice

À supprimer une fois les appelants migrés vers `from xlsxtiles import ...`.
"""

from __future__ import annotations

from xlsxtiles.calibration import calibrate_grid, load_calibration
from xlsxtiles.grid import SheetGrid, build_grid, verify_grid
from xlsxtiles.manifest import render_workbook
from xlsxtiles.models import RenderedTile, TilePlan
from xlsxtiles.plan import apply_display_widths, compute_display_widths, plan_tiles
from xlsxtiles.render import render_citation, render_sheet_tiles
from xlsxtiles.soffice import find_soffice, recalc_with_libreoffice, seed_profile

__all__ = [
    "render_workbook", "render_sheet_tiles", "render_citation",
    "plan_tiles", "build_grid", "verify_grid", "SheetGrid",
    "calibrate_grid", "load_calibration",
    "compute_display_widths", "apply_display_widths",
    "find_soffice", "seed_profile", "recalc_with_libreoffice",
    "TilePlan", "RenderedTile",
]

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        raise SystemExit("usage: python -m xlsxtiles render <fichier.xlsx> <outdir>")
    from xlsxtiles.__main__ import main
    raise SystemExit(main(["render", sys.argv[1], sys.argv[2]]))
