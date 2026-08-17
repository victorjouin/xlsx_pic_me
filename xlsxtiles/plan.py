"""
xlsxtiles.plan — où couper la feuille, et à quelle largeur d'affichage.

Deux décisions prises avant toute conversion :

  1. les LARGEURS d'affichage : une colonne trop étroite ne tronque pas, elle
     affiche `###` — perte totale d'information dans le screenshot ;
  2. le PLAN DE TUILES : des plages qui tiennent dans le budget du vectorizer
     sans couper un graphique en deux.

Aucune E/S vers LibreOffice ici : ce module raisonne sur la grille prédite.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl

from .config import (
    AUTOFIT_SAMPLE_ROWS, CALIBRI_11_CAP_PX, DEFAULT_PATCH_BUDGET,
    MAX_AUTOFIT_CHARS, MIN_ROWS_PER_TILE, MIN_TEXT_PX_AFTER_RESIZE,
    PX_PER_PATCH, RENDER_DPI,
)
from .grid import SheetGrid
from .models import SheetProfile, TilePlan
from .refs import col_letters

__all__ = ["compute_display_widths", "apply_display_widths", "plan_tiles"]


# --------------------------------------------------------------------------
# Largeurs d'affichage
# --------------------------------------------------------------------------

_FMT_LEN = {"date": 10.0, "datetime": 19.0, "time": 8.0}


def _display_len(value: Any, kind: str) -> float:
    """Longueur affichée approximative d'une valeur selon sa sémantique."""
    if value is None:
        return 0.0
    if kind in _FMT_LEN:
        return _FMT_LEN[kind]
    if kind in ("currency", "number", "percent") and isinstance(value, (int, float)):
        base = f"{value:,.2f}"
        return len(base) + (2.0 if kind == "currency" else 0.0) + (
            1.0 if kind == "percent" else 0.0)
    return float(len(str(value)))


def compute_display_widths(src: Path, sheet_name: str,
                           sp: SheetProfile) -> dict[int, float]:
    """Largeur minimale, en caractères, pour qu'aucune colonne ne tronque.

    Une colonne numérique ou date trop étroite ne tronque pas : Excel et
    LibreOffice affichent `###`. C'est une perte TOTALE d'information dans le
    screenshot, pire qu'un libellé coupé. On élargit donc avant rendu — la
    largeur d'origine de l'auteur est un plancher, jamais un plafond.
    """
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    ws = wb[sheet_name]
    c0, c1 = sp.true_min_col or 1, sp.true_max_col or 1
    r0 = sp.true_min_row or 1
    needed: dict[int, float] = {}
    try:
        for i, row in enumerate(ws.iter_rows(min_row=r0, max_row=min(
                r0 + AUTOFIT_SAMPLE_ROWS, sp.true_max_row or r0),
                min_col=c0, max_col=c1, values_only=True)):
            for j, val in enumerate(row):
                col = c0 + j
                kind = sp.column_formats.get(str(col), "general")
                # la ligne d'en-tête est du texte quel que soit le format
                length = _display_len(val, "general" if i == 0 else kind)
                if length > needed.get(col, 0.0):
                    needed[col] = length
    finally:
        wb.close()
    return {c: min(v + 1.5, MAX_AUTOFIT_CHARS) for c, v in needed.items() if v}


def apply_display_widths(sp: SheetProfile,
                         widths: dict[int, float]) -> dict[int, float]:
    """Fusionne les largeurs calculées dans le profil (plancher = existant).

    Retourne la table fusionnée, qui est la SEULE à écrire dans le classeur mis
    en scène. Écrire les largeurs calculées seules ferait diverger la grille
    (qui, elle, part du profil fusionné) du rendu réel dès qu'une colonne
    d'origine est plus large que son besoin d'affichage.
    """
    current: dict[int, float] = {}
    for span, w in sp.custom_col_widths.items():
        lo, hi = (int(x) for x in span.split("-"))
        for c in range(lo, min(hi, 16384) + 1):
            current[c] = w
    default_w = sp.default_col_width or 8.43
    merged = {c: max(current.get(c, default_w), widths.get(c, 0.0))
              for c in set(current) | set(widths)}
    sp.custom_col_widths = {f"{c}-{c}": round(w, 4) for c, w in sorted(merged.items())}
    return merged


# --------------------------------------------------------------------------
# Découpe
# --------------------------------------------------------------------------

def _max_tile_height_px(grid: SheetGrid, col_start: int, col_end: int,
                        patch_budget: int) -> float:
    """Hauteur de tuile maximale sans rendre le texte illisible après resize.

    On part de la contrainte du vectorizer et on remonte : la surface totale de
    la tuile ne doit pas dépasser le budget de pixels, ET le facteur de
    réduction ne doit pas faire tomber la hauteur de capitale sous le seuil.
    """
    width = grid.width_px(col_start, col_end)
    if width <= 0:
        return 0.0
    budget_px = patch_budget * PX_PER_PATCH
    cap_px = CALIBRI_11_CAP_PX * (RENDER_DPI / 96.0)
    max_shrink = cap_px / MIN_TEXT_PX_AFTER_RESIZE
    # une réduction s'applique aux deux axes, donc surface_max = budget * shrink^2
    return budget_px * (max_shrink ** 2) / width


def _cut_row_spans(grid: SheetGrid, body_start: int, row_end: int,
                   max_height: float, header_px: float) -> list[tuple[int, int]]:
    """Découpe verticale en accumulant les hauteurs RÉELLES ligne à ligne.

    Un pas fixe (hauteur de la première ligne x N) est faux dès qu'une feuille
    a des lignes de hauteurs différentes — cas courant des en-têtes hauts et des
    lignes de commentaire. L'accumulation garantit que chaque tuile respecte le
    budget quelle que soit l'irrégularité de la grille.
    """
    budget = max_height - header_px      # l'en-tête est répété sur chaque tuile
    if budget <= 0:
        budget = max_height
    spans: list[tuple[int, int]] = []
    start, acc, n = body_start, 0.0, 0
    for r in range(body_start, row_end + 1):
        h = grid.height_px(r, r)
        if acc + h > budget and n >= MIN_ROWS_PER_TILE:
            spans.append((start, r - 1))
            start, acc, n = r, 0.0, 0
        acc += h
        n += 1
    if start <= row_end:
        spans.append((start, row_end))
    return spans or [(body_start, row_end)]


def _snap_to_anchors(spans: list[tuple[int, int]],
                     sp: SheetProfile) -> list[tuple[int, int]]:
    """Repousse toute coupe qui tomberait au milieu d'un graphique ou d'une image.

    Couper un graphique en deux produit deux tuiles inexploitables : ni l'une ni
    l'autre ne porte l'information, et l'utilisateur reçoit une moitié de
    visuel. On préfère une tuile plus haute que le budget à une tuile fausse.
    """
    if not sp.anchors or len(spans) < 2:
        return spans
    boxes = [(a.from_row, a.to_row) for a in sp.anchors if a.to_row >= a.from_row]
    if not boxes:
        return spans

    out: list[tuple[int, int]] = []
    pending_start = spans[0][0]
    for i, (s, e) in enumerate(spans):
        last = i == len(spans) - 1
        cut = e
        if not last:
            # tant que la frontière traverse une ancre, on la descend sous elle
            moved = True
            while moved:
                moved = False
                for f, t in boxes:
                    if f <= cut < t:
                        cut = t
                        moved = True
        if not last and cut >= spans[-1][1]:
            break            # l'ancre absorbe tout le reste de la feuille
        out.append((pending_start, cut))
        pending_start = cut + 1
    if pending_start <= spans[-1][1]:
        out.append((pending_start, spans[-1][1]))
    return [(s, e) for s, e in out if s <= e]


def plan_tiles(sp: SheetProfile, grid: SheetGrid, *,
               patch_budget: int = DEFAULT_PATCH_BUDGET) -> list[TilePlan]:
    """Plan de tuiles selon la stratégie décidée par le profileur."""
    r0, r1 = sp.true_min_row or 1, sp.true_max_row or 1
    c0, c1 = sp.true_min_col or 1, sp.true_max_col or 1

    # En-tête à répéter : table déclarée > volets figés > rien
    header_n = 0
    if sp.tables:
        header_n = max(t.header_row_count for t in sp.tables)
    elif sp.freeze_rows:
        header_n = sp.freeze_rows
    header_rows = f"{r0}:{r0 + header_n - 1}" if header_n else None
    key_cols = (f"{col_letters(c0)}:{col_letters(c0 + sp.freeze_cols - 1)}"
                if sp.freeze_cols else None)

    body_start = r0 + header_n
    header_px = grid.height_px(r0, r0 + header_n - 1) if header_n else 0.0
    max_h = _max_tile_height_px(grid, c0, c1, patch_budget)

    # --- découpage vertical ---
    if sp.tiling_strategy == "whole_sheet":
        row_spans = [(r0, r1)]
    elif sp.tiling_strategy == "page_breaks" and sp.row_breaks:
        bounds = sorted(set(sp.row_breaks))
        row_spans, prev = [], body_start
        for b in bounds:
            if b >= prev:
                row_spans.append((prev, b))
                prev = b + 1
        if prev <= r1:
            row_spans.append((prev, r1))
    else:
        # declared_tables et xy_cut aboutissent ici : l'en-tête de la table est
        # déjà sorti du corps via header_n, le reste se coupe au budget.
        row_spans = _cut_row_spans(grid, body_start, r1, max_h, header_px)

    row_spans = _snap_to_anchors(row_spans, sp)

    # --- découpage horizontal ---
    # On n'ajoute des tuiles colonnes que si la largeur dépasse le budget seule.
    col_spans = [(c0, c1)]
    total_w = grid.width_px(c0, c1)
    budget_w = (patch_budget * PX_PER_PATCH) ** 0.5 * (CALIBRI_11_CAP_PX
                                                       * (RENDER_DPI / 96.0)
                                                       / MIN_TEXT_PX_AFTER_RESIZE)
    if total_w > budget_w * 1.6:
        col_spans, cur, acc = [], c0, 0.0
        key_w = grid.width_px(c0, c0 + sp.freeze_cols - 1) if sp.freeze_cols else 0.0
        for c in range(c0, c1 + 1):
            w = grid.col_edges[c][1] - grid.col_edges[c][0]
            if acc + w + key_w > budget_w and c > cur:
                col_spans.append((cur, c - 1))
                cur, acc = c, 0.0
            acc += w
        col_spans.append((cur, c1))

    plans, i = [], 0
    for cs, ce in col_spans:
        for rs, re_ in row_spans:
            rng = f"{col_letters(cs)}{rs}:{col_letters(ce)}{re_}"
            # key_cols est un réglage d'impression du classeur, pas de la tuile :
            # LibreOffice l'applique à toutes les pages d'une même conversion.
            plans.append(TilePlan(i, rs, re_, cs, ce, header_rows, key_cols, rng))
            i += 1
    return plans
