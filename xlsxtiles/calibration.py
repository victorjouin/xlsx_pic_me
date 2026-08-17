from __future__ import annotations
import json
import os
import tempfile
from dataclasses import asdict
from pathlib import Path
from typing import Sequence

import openpyxl
import pymupdf

from .config import PROBE_HEIGHTS, PROBE_WIDTHS
from .grid import cluster, segments
from .models import GridCalibration
from .refs import col_letters
from .soffice import seed_profile, to_pdf

__all__ = ["calibrate_grid", "load_calibration"]


def _fit_line(xs: Sequence[float], ys: Sequence[float]) -> tuple[float, float]:
    """Moindres carrés y = a*x + b."""
    n = len(xs)
    mx, my = sum(xs) / n, sum(ys) / n
    den = sum((x - mx) ** 2 for x in xs)
    a = sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / den if den else 0.0
    return a, my - a * mx


def calibrate_grid(*, font: str = "Calibri", size: int = 11) -> GridCalibration:
    """Mesure les constantes de conversion sur une feuille aux cotes connues.

    Pas de paramètre `dpi` : le résultat est exprimé à 96 dpi par construction
    (`GridCalibration.dpi = 96`) et la mise à l'échelle vers le DPI de rendu se
    fait au moment de bâtir la grille. Un `dpi` ici n'aurait rien changé.

    Les valeurs par défaut de GridCalibration dépendent de la police Normal, du
    DPI et de la version de LibreOffice — ce ne sont pas des constantes. Cette
    passe rend une sonde dont on connaît exactement les largeurs et hauteurs,
    mesure les traits réellement dessinés, et en déduit `mdw`,
    `col_padding_px` et `row_scale`.

    À exécuter UNE FOIS au build du conteneur et à sérialiser : c'est ce qui
    fait passer `calibrated` à True, donc ce qui autorise la vérification de
    grille et, plus tard, le surlignage de cellule.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    for j, w in enumerate(PROBE_WIDTHS, start=1):
        ws.column_dimensions[col_letters(j)].width = w
    for i, h in enumerate(PROBE_HEIGHTS, start=1):
        ws.row_dimensions[i].height = h
        for j in range(1, len(PROBE_WIDTHS) + 1):
            cell = ws.cell(row=i, column=j, value="X")
            cell.font = openpyxl.styles.Font(name=font, size=size)

    ws.print_area = f"A1:{col_letters(len(PROBE_WIDTHS))}{len(PROBE_HEIGHTS)}"
    ws.page_setup.scale = 100
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    ws.print_options.gridLines = True
    ws.print_options.headings = False
    for attr in ("left", "right", "top", "bottom"):
        setattr(ws.page_margins, attr, 0.1)
    # page volontairement large : on mesure, on ne cadre pas
    ws.page_setup.paperWidth = f"{sum(PROBE_WIDTHS) * 3.0:.1f}mm"
    ws.page_setup.paperHeight = f"{sum(PROBE_HEIGHTS) * 2.0:.1f}mm"
    ws.page_setup.paperSize = None

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        prof_dir = seed_profile(tmp / "loprofile")
        probe = tmp / "probe.xlsx"
        wb.save(probe)
        pdf = to_pdf(probe, tmp, prof_dir)
        with pymupdf.open(pdf) as doc:
            vx, hy = segments(doc[0])
            vx, hy = cluster(vx), cluster(hy)

    px_per_pt = 96.0 / 72.0
    w_px = [(b - a) * px_per_pt for a, b in zip(vx, vx[1:])]
    h_px = [(b - a) * px_per_pt for a, b in zip(hy, hy[1:])]
    # LibreOffice ne trace ni la bordure gauche ni la bordure haute de la zone
    # imprimée : on récupère N-1 écarts pour N colonnes. Les écarts obtenus
    # correspondent donc aux DERNIÈRES cotes de la sonde, d'où l'alignement par
    # la fin — s'aligner par le début décalerait tout l'ajustement d'un cran.
    if len(w_px) < 3 or len(h_px) < 3:
        raise RuntimeError(
            f"calibration : sonde illisible ({len(w_px)} largeurs / "
            f"{len(h_px)} hauteurs mesurées). Traits de grille absents du PDF ?")

    nw, nh = min(len(w_px), len(PROBE_WIDTHS)), min(len(h_px), len(PROBE_HEIGHTS))
    mdw, pad = _fit_line(PROBE_WIDTHS[-nw:], w_px[-nw:])
    ratios = [m / (h * px_per_pt)
              for m, h in zip(h_px[-nh:], PROBE_HEIGHTS[-nh:])]
    row_scale = sum(ratios) / len(ratios)

    return GridCalibration(mdw=round(mdw, 4), dpi=96,
                           col_padding_px=round(pad, 4),
                           row_scale=round(row_scale, 5), col_scale=1.0,
                           calibrated=True, source="probe-sheet")


def _default_cache() -> Path:
    env = os.environ.get("XLSX_CALIBRATION")
    return Path(env) if env else Path.home() / ".cache" / "xlsx_calibration.json"


def load_calibration(*, cache: Path | None = None,
                     refresh: bool = False) -> GridCalibration:
    """Calibration mise en cache, sonde relancée si absente ou illisible.

    La sonde coûte une conversion soffice : elle ne doit tourner qu'une fois
    par image de conteneur, pas une fois par document. Le cache est invalidé
    à la main (refresh=True) — après une mise à jour de LibreOffice ou un
    changement de police par défaut, les constantes bougent.
    """
    cache = cache or _default_cache()
    if not refresh and cache.exists():
        try:
            return GridCalibration(**json.loads(cache.read_text(encoding="utf-8")))
        except (json.JSONDecodeError, TypeError, OSError):
            pass                      # cache illisible : on remesure
    cal = calibrate_grid()
    try:
        cache.parent.mkdir(parents=True, exist_ok=True)
        cache.write_text(json.dumps(asdict(cal), indent=2), encoding="utf-8")
    except OSError:
        pass                          # système de fichiers en lecture seule
    return cal
