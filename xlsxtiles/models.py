"""
xlsxtiles.models — structures de données partagées, sérialisables en JSON.

Aucune logique métier ici : uniquement la forme des objets qui circulent entre
le profileur, le planificateur et le moteur de rendu. Un module qui a besoin
d'un calcul sur ces objets le met chez lui, pas ici.
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from typing import Any

from .config import BYTES_PER_CELL_ESTIMATE

__all__ = ["GridCalibration", "TableInfo", "AnchorBox", "SheetProfile",
           "Verdict", "WorkbookProfile", "TilePlan", "RenderedTile"]


@dataclass
class GridCalibration:
    """Paramètres de conversion unités Excel -> pixels.

    IMPORTANT : ces valeurs par défaut sont un point de départ, PAS une vérité.
    Elles dépendent de la police Normal du classeur, du DPI de rendu et de la
    version de LibreOffice. Elles doivent être :
      1. calibrées au build sur une feuille sonde (`calibration.calibrate_grid`),
      2. re-vérifiées à chaque document contre les traits réels du PDF rendu
         (`grid.verify_grid`).

    mdw : largeur en px du chiffre le plus large de la police Normal
          (Calibri 11 @96dpi -> 7 ; Arial 10 -> 6 ; Calibri 12 -> 8)
    """
    mdw: float = 7.0
    dpi: int = 96
    col_padding_px: float = 5.0     # 2+2 marge + 1 bordure
    row_scale: float = 1.0          # correctif mesuré sur la feuille sonde
    col_scale: float = 1.0
    calibrated: bool = False        # False => surlignage cellule interdit
    source: str = "default"         # "default" | "probe-sheet" | "pdf-refit"


@dataclass
class TableInfo:
    name: str
    ref: str
    header_row_count: int
    totals_row_count: int


@dataclass
class AnchorBox:
    """Zone occupée par un graphique ou une image — incoupable au tuilage."""
    kind: str          # "chart" | "image" | "shape"
    from_col: int
    from_row: int
    to_col: int
    to_row: int


@dataclass
class SheetProfile:
    name: str
    index: int
    state: str                      # visible | hidden | veryHidden
    part: str                       # chemin dans le zip
    xml_size_mb: float
    deep_scanned: bool

    declared_dimension: str | None = None
    # Bornes réelles = cellules porteuses de valeur/formule, hors formatage seul
    true_min_row: int | None = None
    true_max_row: int | None = None
    true_min_col: int | None = None
    true_max_col: int | None = None

    n_cells_with_content: int = 0
    n_formulas: int = 0
    n_formulas_with_cached_value: int = 0
    fill_ratio: float = 0.0

    freeze_rows: int = 0
    freeze_cols: int = 0
    auto_filter_ref: str | None = None
    print_area: str | None = None
    print_title_rows: str | None = None
    print_title_cols: str | None = None
    row_breaks: list[int] = field(default_factory=list)
    col_breaks: list[int] = field(default_factory=list)

    n_merged_cells: int = 0
    n_conditional_formats: int = 0
    n_data_validations: int = 0
    n_hyperlinks: int = 0
    hidden_rows: list[int] = field(default_factory=list)
    hidden_cols: list[int] = field(default_factory=list)
    max_outline_level: int = 0

    default_col_width: float | None = None
    default_row_height: float | None = None
    custom_col_widths: dict[str, float] = field(default_factory=dict)   # "min-max" -> width
    custom_row_heights: dict[str, float] = field(default_factory=dict)  # "row" -> hauteur pt
    row_heights_truncated: bool = False

    # index de colonne (1-based, en str pour JSON) -> sémantique dominante
    column_formats: dict[str, str] = field(default_factory=dict)
    column_format_codes: dict[str, str] = field(default_factory=dict)
    column_types: dict[str, str] = field(default_factory=dict)   # text | number | bool | error

    tables: list[TableInfo] = field(default_factory=list)
    anchors: list[AnchorBox] = field(default_factory=list)
    n_charts: int = 0
    n_images: int = 0
    has_pivot_table: bool = False

    depends_on_sheets: list[str] = field(default_factory=list)

    # Verdicts dérivés
    is_sparse: bool = False
    is_layout_sheet: bool = False
    is_pure_data: bool = False        # ni graphique, ni image, ni MFC
    tiling_strategy: str = "xy_cut"   # declared_tables | page_breaks | xy_cut | whole_sheet
    image_indexing: str = "full"      # skip | lazy | full

    @property
    def effective_cells(self) -> int:
        """Cellules réelles si scan profond, estimation sinon.

        Sans ce repli, une feuille scannée en surface compte pour 0 et le
        routage renvoie vers Lambda un classeur qui l'y fera exploser.
        """
        if self.deep_scanned:
            return self.n_cells_with_content
        return int(self.xml_size_mb * 1e6 / BYTES_PER_CELL_ESTIMATE)

    @property
    def true_n_rows(self) -> int:
        if self.true_min_row is None or self.true_max_row is None:
            return 0
        return self.true_max_row - self.true_min_row + 1

    @property
    def true_n_cols(self) -> int:
        if self.true_min_col is None or self.true_max_col is None:
            return 0
        return self.true_max_col - self.true_min_col + 1


@dataclass
class Verdict:
    compute_target: str              # "lambda" | "fargate" | "batch"
    needs_recalc: bool               # valeurs en cache absentes -> passe LibreOffice
    grid_reliable: bool              # polices présentes -> surlignage cellule possible
    missing_fonts: list[str] = field(default_factory=list)
    sheets_to_ingest: list[str] = field(default_factory=list)
    sheets_skipped: dict[str, str] = field(default_factory=dict)  # nom -> raison
    warnings: list[str] = field(default_factory=list)


@dataclass
class WorkbookProfile:
    path: str
    file_size_mb: float
    n_sheets: int
    sheets: list[SheetProfile]
    fonts_used: list[str]
    defined_names: dict[str, str]
    has_external_links: bool
    has_pivot_cache: bool
    has_vba: bool
    total_cells_with_content: int
    estimated_total_cells: int
    total_charts: int
    total_images: int
    sheet_dependency_edges: list[tuple[str, str]]
    calibration: GridCalibration
    verdict: Verdict

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    def to_json(self, **kw: Any) -> str:
        kw.setdefault("ensure_ascii", False)
        kw.setdefault("indent", 2)
        return json.dumps(self.to_dict(), **kw)

    def sheet(self, name: str) -> SheetProfile | None:
        return next((s for s in self.sheets if s.name == name), None)


@dataclass
class TilePlan:
    """Une tuile planifiée : sa plage exacte dans la feuille, avant rendu."""
    index: int
    row_start: int
    row_end: int
    col_start: int
    col_end: int
    header_rows: str | None       # "1:1" — répété par LibreOffice sur la tuile
    key_cols: str | None          # "A:A"
    range_a1: str

    @property
    def n_rows(self) -> int:
        return self.row_end - self.row_start + 1


@dataclass
class RenderedTile:
    """Une tuile effectivement rendue, avec le résultat de sa vérification."""
    plan: TilePlan
    png_path: str
    width_px: int
    height_px: int
    grid_verified: bool
    grid_drift_px: float
    n_anchor_checks: int
