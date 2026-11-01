"""Renderer for Data Table slide (CORE-10).

Renders dynamic, structured tables with automated column distribution,
row highlighting, and cell typography scaling.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Length
from pydantic import Field, model_validator

from slide_engine.core_base import BoundingBox, ColorType, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import BasePayload, SafeText, ShortText, TemplateSpec

if TYPE_CHECKING:
    from pptx.slide import Slide
    from pptx.table import Table


class Core10Payload(BasePayload):
    """Payload for Data Table slide (CORE-10)."""

    headers: list[ShortText] = Field(min_length=1, max_length=8)
    rows: list[list[SafeText]] = Field(min_length=1, max_length=15)
    highlight_row_index: int | None = Field(default=None, ge=0)

    @model_validator(mode="after")
    def validate_table_integrity(self) -> Core10Payload:
        """Validate column count consistency across all rows and highlight row index."""
        col_count = len(self.headers)
        for r_idx, row in enumerate(self.rows):
            if len(row) != col_count:
                raise ValueError(
                    f"Row {r_idx} column count ({len(row)}) does not match headers count ({col_count})."
                )
        if self.highlight_row_index is not None and self.highlight_row_index >= len(self.rows):
            raise ValueError(
                f"highlight_row_index ({self.highlight_row_index}) exceeds total row count ({len(self.rows)})."
            )
        return self

    def validate_split_density(self, pane_name: str, template_id: str) -> None:
        """Enforce row and column limits for split layout."""
        if len(self.rows) > 8:
            raise ValueError(
                f"{pane_name} table '{template_id}' in split layout exceeds legibility limit: "
                f"{len(self.rows)} rows > 8 rows maximum."
            )
        if len(self.headers) > 4:
            raise ValueError(
                f"{pane_name} table '{template_id}' in split layout exceeds legibility limit: "
                f"{len(self.headers)} columns > 4 columns maximum."
            )


_SAMPLE_CORE10: dict[str, Any] = {
    "headers": [f"指標 {c}" for c in range(1, 9)],
    "rows": [[f"値{r}-{c}" for c in range(1, 9)] for r in range(1, 16)],
    "highlight_row_index": 0,
}

_SAMPLE_CORE10_SPLIT: dict[str, Any] = {
    "headers": [f"評価指標 {c}" for c in range(1, 5)],
    "rows": [[f"データ {r}-{c}" for c in range(1, 5)] for r in range(1, 9)],
    "highlight_row_index": 0,
}

_INVALID_SAMPLES_CORE10: list[dict[str, Any]] = [
    {"headers": [], "rows": [["1"]]},  # Violates min_length=1 for headers
    {"headers": ["A", "B"], "rows": [["1"]]},  # Row column count mismatch
    {"headers": ["A"], "rows": [["1"]], "highlight_row_index": 5},  # Highlight index out of bounds
]

SPEC = TemplateSpec(
    payload_type=Core10Payload,
    supports_split=True,
    description="Data Table",
    sample_payload=_SAMPLE_CORE10,
    split_sample_payload=_SAMPLE_CORE10_SPLIT,
    invalid_sample_payloads=_INVALID_SAMPLES_CORE10,
)


def _resolve_row_colors(is_highlighted: bool, is_even: bool) -> tuple[ColorType, ColorType]:
    """Determine background and text colors for a data row."""
    if is_highlighted:
        return SlideMaster.COLORS["fallback_bg"], SlideMaster.COLORS["primary"]
    if is_even:
        return SlideMaster.COLORS["bg_white"], SlideMaster.COLORS["text_dark"]
    return SlideMaster.COLORS["bg_light"], SlideMaster.COLORS["text_dark"]


def _render_data_row(
    table: Table,
    row_idx: int,
    row_data: list[str],
    num_cols: int,
    avail_w: Length,
    avail_h: Length,
    is_highlighted: bool,
) -> None:
    """Render a single data row within the table."""
    bg_col, text_col = _resolve_row_colors(is_highlighted, row_idx % 2 == 0)

    for col_idx in range(num_cols):
        cell = table.cell(row_idx + 1, col_idx)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        ShapeBuilder.set_cell_solid_fill(cell, bg_col)

        cell_text = row_data[col_idx] if col_idx < len(row_data) else ""
        tf = cell.text_frame
        tf.word_wrap = True
        ShapeBuilder.add_styled_paragraph(
            tf,
            cell_text,
            avail_w,
            avail_h,
            max_pt=10.0,
            min_pt=8.0,
            color_key="primary" if is_highlighted else "text_dark",
            bold=is_highlighted,
            is_first=True,
        )


def render_core_10(slide: Slide, payload: Core10Payload, box: BoundingBox) -> None:
    """Render structured data table matching bounding box dimensions.

    Args:
        slide: Target slide.
        payload: Core10Payload instance.
        box: Content BoundingBox.
    """
    num_cols = len(payload.headers)
    num_rows = len(payload.rows) + 1  # Headers + data rows

    table_shape = slide.shapes.add_table(num_rows, num_cols, box.left, box.top, box.width, box.height)
    table = table_shape.table

    # 1. Uniform column width & row height distribution
    col_width = Length(int(box.width / num_cols))
    row_height = Length(int(box.height / num_rows))
    for col_idx in range(num_cols):
        table.columns[col_idx].width = col_width
    for r_idx in range(num_rows):
        table.rows[r_idx].height = row_height

    # Deduct cell horizontal (6pt) and vertical (3pt) margins for font fitting
    avail_w = Length(max(int(FontEngine.pt(12)), int(col_width - FontEngine.pt(6.0))))
    avail_h = Length(max(int(FontEngine.pt(10)), int(row_height - FontEngine.pt(3.0))))

    # 2. Render Header Row
    for col_idx, header_text in enumerate(payload.headers):
        cell = table.cell(0, col_idx)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        ShapeBuilder.set_cell_solid_fill(cell, SlideMaster.COLORS["primary"])

        tf = cell.text_frame
        tf.word_wrap = True
        ShapeBuilder.add_styled_paragraph(
            tf,
            header_text,
            avail_w,
            avail_h,
            max_pt=11.0,
            min_pt=8.0,
            color_key="text_light",
            bold=True,
            alignment=PP_ALIGN.CENTER,
            is_first=True,
        )

    # 3. Render Data Rows
    for row_idx, row_data in enumerate(payload.rows):
        is_highlighted = (
            payload.highlight_row_index is not None and row_idx == payload.highlight_row_index
        )
        _render_data_row(table, row_idx, row_data, num_cols, avail_w, avail_h, is_highlighted)
