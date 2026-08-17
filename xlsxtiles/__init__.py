"""
    python -m xlsxtiles render budget2026.xlsx -o manifest.json
    python -m xlsxtiles plan   budget2026.xlsx

    profile.py  _ géométrie des feuilles, via openpyxl
    soffice.py  _ pilotage de LibreOffice
    tiles.py    _ plan de découpe, mise en scène, rasterisation, manifeste

Dépendances : openpyxl, pymupdf, LibreOffice
"""
from __future__ import annotations

from .profile import Sheet, profile_workbook
from .soffice import find_soffice
from .tiles import PNG_DATA_URI_PREFIX, Tile, plan_tiles, render_workbook

__version__ = "1.0.0"

__all__ = [
    "render_workbook", "plan_tiles", "profile_workbook",
    "Sheet", "Tile", "find_soffice", "PNG_DATA_URI_PREFIX",
]
