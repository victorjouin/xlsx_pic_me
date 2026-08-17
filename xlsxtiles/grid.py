"""
xlsxtiles.grid — la grille pixel, prédite puis mesurée.

Deux moitiés qui ne doivent pas être confondues :

  - PRÉDICTION : `build_grid()` déduit la position de chaque colonne et ligne
    des métadonnées du classeur et d'une calibration. C'est une hypothèse.
  - MESURE : `segments()` / `verify_grid()` lisent la géométrie RÉELLEMENT
    dessinée dans le PDF rendu. C'est la seule vérité disponible.

`verify_grid()` confronte les deux et chiffre l'écart. Tant que cet écart n'est
pas mesuré, aucune affirmation sur le cadrage n'est fondée.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import pymupdf

from .config import EDGE_CLUSTER_PT, RENDER_DPI
from .models import GridCalibration, SheetProfile
from .refs import excel_height_to_px, excel_width_to_px

__all__ = ["SheetGrid", "build_grid", "segments", "cluster", "verify_grid",
           "content_clip"]


# --------------------------------------------------------------------------
# Prédiction
# --------------------------------------------------------------------------

@dataclass
class SheetGrid:
    """Position pixel cumulée de chaque colonne / ligne, à RENDER_DPI."""
    col_edges: dict[int, tuple[float, float]] = field(default_factory=dict)
    row_edges: dict[int, tuple[float, float]] = field(default_factory=dict)

    def bbox(self, col: int, row: int) -> tuple[float, float, float, float] | None:
        c, r = self.col_edges.get(col), self.row_edges.get(row)
        if c is None or r is None:
            return None
        return (c[0], r[0], c[1], r[1])

    def width_px(self, c0: int, c1: int) -> float:
        if c0 not in self.col_edges or c1 not in self.col_edges:
            return 0.0
        return self.col_edges[c1][1] - self.col_edges[c0][0]

    def height_px(self, r0: int, r1: int) -> float:
        if r0 not in self.row_edges or r1 not in self.row_edges:
            return 0.0
        return self.row_edges[r1][1] - self.row_edges[r0][0]


def build_grid(sp: SheetProfile, cal: GridCalibration,
               row_heights: dict[int, float] | None = None) -> SheetGrid:
    """Grille pixel de la feuille depuis les largeurs/hauteurs du profil.

    Les colonnes hors `custom_col_widths` prennent la largeur par défaut ; idem
    pour les lignes. C'est cette grille que verify_grid() confronte au PDF réel.
    """
    grid = SheetGrid()
    default_w = sp.default_col_width or 8.43
    default_h = sp.default_row_height or 15.0
    scale = RENDER_DPI / 96.0

    # Hauteurs personnalisées relevées par le profileur. Les ignorer fait
    # dériver la hauteur de tuile, donc la taille de page, donc la pagination.
    if row_heights is None:
        row_heights = {int(r): h for r, h in sp.custom_row_heights.items()}

    # largeurs personnalisées : "min-max" -> largeur
    custom: dict[int, float] = {}
    for span, w in sp.custom_col_widths.items():
        lo, hi = (int(x) for x in span.split("-"))
        for c in range(lo, min(hi, 16384) + 1):
            custom[c] = w

    hidden_cols = set(sp.hidden_cols)
    x = 0.0
    for c in range(sp.true_min_col or 1, (sp.true_max_col or 1) + 1):
        w = (0.0 if c in hidden_cols
             else excel_width_to_px(custom.get(c), cal, default_w) * scale)
        grid.col_edges[c] = (x, x + w)
        x += w

    hidden_rows = set(sp.hidden_rows)
    y = 0.0
    for r in range(sp.true_min_row or 1, (sp.true_max_row or 1) + 1):
        h = 0.0 if r in hidden_rows else excel_height_to_px(
            row_heights.get(r), cal, default_h) * scale
        grid.row_edges[r] = (y, y + h)
        y += h

    return grid


# --------------------------------------------------------------------------
# Mesure du PDF rendu
# --------------------------------------------------------------------------

def segments(page: pymupdf.Page) -> tuple[list[float], list[float]]:
    """Positions (en points) des traits verticaux et horizontaux dessinés.

    Les traits de grille sont émis par LibreOffice tantôt comme lignes, tantôt
    comme rectangles très fins ; on accepte les deux. C'est la seule mesure
    objective de la géométrie réellement rendue — tout le reste est prédiction.
    """
    vx: list[float] = []
    hy: list[float] = []
    for d in page.get_drawings():
        for item in d.get("items", ()):
            op = item[0]
            if op == "l":
                p1, p2 = item[1], item[2]
                if abs(p1.x - p2.x) < 0.3 and abs(p1.y - p2.y) > 1.0:
                    vx.append((p1.x + p2.x) / 2)
                elif abs(p1.y - p2.y) < 0.3 and abs(p1.x - p2.x) > 1.0:
                    hy.append((p1.y + p2.y) / 2)
            elif op == "re":
                r = item[1]
                if r.width < 1.5 and r.height > 1.0:
                    vx.append((r.x0 + r.x1) / 2)
                elif r.height < 1.5 and r.width > 1.0:
                    hy.append((r.y0 + r.y1) / 2)
    return vx, hy


def cluster(values: list[float], tol: float = EDGE_CLUSTER_PT) -> list[float]:
    """Regroupe des positions voisines en arêtes uniques."""
    if not values:
        return []
    vals = sorted(values)
    out, cur = [], [vals[0]]
    for v in vals[1:]:
        if v - cur[-1] <= tol:
            cur.append(v)
        else:
            out.append(sum(cur) / len(cur))
            cur = [v]
    out.append(sum(cur) / len(cur))
    return out


def verify_grid(page: pymupdf.Page, grid: SheetGrid, col_start: int, col_end: int,
                dpi: int = RENDER_DPI) -> tuple[float, int]:
    """Confronte les largeurs de colonne PRÉDITES aux traits réellement dessinés.

    Retourne (dérive max en px de rendu, nombre d'arêtes confrontées). On
    compare des LARGEURS successives et non des positions absolues : la tuile
    est rognée et translatée, seules les distances entre arêtes sont
    comparables d'un référentiel à l'autre.
    """
    vx = cluster(segments(page)[0])
    expected = [grid.col_edges[c][1] - grid.col_edges[c][0]
                for c in range(col_start, col_end + 1)
                if c in grid.col_edges]
    expected = [w for w in expected if w > 0.5]
    if len(vx) < 2 or not expected:
        return float("inf"), 0

    px_per_pt = dpi / 72.0
    measured = [(b - a) * px_per_pt * (RENDER_DPI / dpi)
                for a, b in zip(vx, vx[1:])]
    # les colonnes clés répétées ajoutent des arêtes en tête : on aligne sur la
    # fin, qui correspond toujours à la dernière colonne de la tuile
    n = min(len(measured), len(expected))
    if n == 0:
        return float("inf"), 0
    drift = max(abs(m - e) for m, e in zip(measured[-n:], expected[-n:]))
    return drift, n


def content_clip(page: pymupdf.Page, pad_pt: float) -> pymupdf.Rect | None:
    """Rectangle englobant l'encre réelle de la page (traits de grille + texte).

    La page PDF est dimensionnée sur la plus GRANDE tuile de la bande ; les
    tuiles plus courtes laissent donc du blanc en bas, et la dernière bande de
    colonnes du blanc à droite. Ce blanc n'est pas neutre : le vectorizer
    redimensionne l'image entière, donc chaque pixel vide consomme du budget de
    patches et rétrécit d'autant le texte utile.
    """
    rect = pymupdf.Rect()
    rect.x0, rect.y0, rect.x1, rect.y1 = 0, 0, -1, -1   # rect vide
    for d in page.get_drawings():
        rect = d["rect"] if rect.is_empty else rect | d["rect"]
    for w in page.get_text("words"):
        wr = pymupdf.Rect(w[:4])
        rect = wr if rect.is_empty else rect | wr
    if rect.is_empty:
        return None
    rect = pymupdf.Rect(rect.x0 - pad_pt, rect.y0 - pad_pt,
                        rect.x1 + pad_pt, rect.y1 + pad_pt)
    return rect & page.rect
