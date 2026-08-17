"""
xlsxtiles — découpe un classeur Excel en images fidèles et localisées.

Produit, pour chaque feuille, des tuiles PNG cadrées sur leur contenu, chacune
accompagnée de la plage A1 exacte qu'elle représente. Le périmètre s'arrête là :
la vectorisation, l'indexation et la restitution à l'utilisateur sont des étages
aval, qui consomment `manifest.json`.

    from xlsxtiles import render_workbook
    manifest = render_workbook("budget2026.xlsx", "out/")

Ou en ligne de commande :

    python -m xlsxtiles render budget2026.xlsx out/
    python -m xlsxtiles profile budget2026.xlsx
    python -m xlsxtiles calibrate --refresh

Organisation des modules, du bas vers le haut (aucun cycle) :

    config      seuils et budgets réglables
    models      dataclasses sérialisables
    refs        références A1, unités Excel -> pixels
    numfmt      classification des formats de nombre
    ooxml       accès bas niveau au zip
    scan        parcours du XML d'une feuille
    fonts       polices demandées vs disponibles
    profiler    profil du classeur + verdicts
    soffice     pilotage de LibreOffice
    grid        grille prédite + mesure du PDF rendu
    calibration feuille sonde + cache
    plan        largeurs d'affichage + plan de tuiles
    render      mise en scène et rasterisation
    encode      transport des tuiles en base64
    manifest    pilote classeur + contrat de sortie

Dépendances : openpyxl, pymupdf, et LibreOffice accessible (voir `soffice`).
"""

from __future__ import annotations

from .calibration import calibrate_grid, load_calibration
from .encode import (
    PNG_DATA_URI_PREFIX, base64_to_png, png_to_base64, png_to_data_uri,
)
from .grid import SheetGrid, build_grid, verify_grid
from .manifest import render_workbook
from .models import (
    AnchorBox, GridCalibration, RenderedTile, SheetProfile, TableInfo,
    TilePlan, Verdict, WorkbookProfile,
)
from .plan import plan_tiles
from .profiler import profile_workbook, self_check
from .render import render_citation, render_sheet_tiles
from .soffice import find_soffice

__version__ = "0.2.0"

__all__ = [
    # profilage
    "profile_workbook", "self_check",
    # rendu
    "render_workbook", "render_sheet_tiles", "render_citation",
    # transport base64
    "png_to_base64", "png_to_data_uri", "base64_to_png", "PNG_DATA_URI_PREFIX",
    # calibration et grille
    "calibrate_grid", "load_calibration", "build_grid", "verify_grid",
    "SheetGrid", "find_soffice",
    # plan
    "plan_tiles",
    # modèles
    "WorkbookProfile", "SheetProfile", "GridCalibration", "Verdict",
    "TableInfo", "AnchorBox", "TilePlan", "RenderedTile",
]
