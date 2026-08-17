"""
xlsxtiles.refs — références de cellules et conversion des unités Excel en pixels.

Deux familles de fonctions pures, sans état ni E/S :
  - traduction entre notation A1 et indices 1-based,
  - traduction des largeurs (en caractères) et hauteurs (en points) d'Excel
    vers des pixels, sous une calibration donnée.
"""

from __future__ import annotations

import re

from .models import GridCalibration

__all__ = ["col_letters", "col_to_idx", "parse_cell_ref",
           "excel_width_to_px", "excel_height_to_px"]

_CELL_REF_RE = re.compile(r"^([A-Z]+)([0-9]+)$")


def col_letters(idx: int) -> str:
    """1 -> 'A', 27 -> 'AA'."""
    out = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        out = chr(65 + rem) + out
    return out


def col_to_idx(letters: str) -> int:
    """'A' -> 1, 'AA' -> 27."""
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - 64)
    return idx


def parse_cell_ref(ref: str) -> tuple[int, int] | None:
    """'AB12' -> (28, 12). None si la référence n'est pas une cellule simple."""
    m = _CELL_REF_RE.match(ref)
    if not m:
        return None
    return col_to_idx(m.group(1)), int(m.group(2))


def excel_width_to_px(width_chars: float | None, cal: GridCalibration,
                      default_width: float = 8.43) -> float:
    """Largeur de colonne Excel (en caractères) -> pixels.

    La spec OOXML donne la conversion PIXELS -> LARGEUR STOCKÉE :
        width = trunc((px * MDW + 5) / MDW * 256) / 256
    Le sens qui nous intéresse est l'inverse :
        px = trunc(chars * MDW + padding)

    Contrôle de non-régression : 8.43 caractères (défaut Calibri 11 @96dpi)
    doit rendre 64 px SOUS LA CALIBRATION PAR DÉFAUT. Toute modification qui
    casse cette égalité est un bug. Appliquer ici la formule de la spec telle
    quelle donne 59 px — erreur de 5 px par colonne, soit ~100 px de dérive sur
    20 colonnes.

    Sous une calibration mesurée, mdw et le padding changent : LibreOffice 26
    en Calibri 11 rend px = 7.4005 * chars, sans padding.
    """
    w = default_width if width_chars is None else width_chars
    px = int(w * cal.mdw + cal.col_padding_px)
    return px * cal.col_scale * (cal.dpi / 96.0)


def excel_height_to_px(height_pt: float | None, cal: GridCalibration,
                       default_height: float = 15.0) -> float:
    """Hauteur de ligne Excel (en points) -> pixels."""
    h = default_height if height_pt is None else height_pt
    return h * (cal.dpi / 72.0) * cal.row_scale
