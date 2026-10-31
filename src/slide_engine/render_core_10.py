"""Renderer for Data Table slide (CORE-10).

Renders dynamic, structured tables with automated column distribution,
row highlighting, and cell typography scaling.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

from slide_engine.core_base import BoundingBox, FontEngine, SlideMaster
from slide_engine.models import Core10Payload

if TYPE_CHECKING:
    from pptx.slide import Slide


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

    # 1. Uniform column width distribution
    col_width = Inches(box.width / num_cols)
    for col in table.columns:
        col.width = col_width

    # 2. Render Header Row
    for col_idx, header_text in enumerate(payload.headers):
        cell = table.cell(0, col_idx)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        cell.fill.fore_color.rgb = SlideMaster.COLORS["primary"]

        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = header_text
        p.alignment = PP_ALIGN.CENTER
        p.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p.font.size = FontEngine.pt(11.0)
        p.font.bold = True
        p.font.color.rgb = SlideMaster.COLORS["text_light"]

    # 3. Render Data Rows
    for row_idx, row_data in enumerate(payload.rows):
        is_highlighted = (payload.highlight_row_index is not None and row_idx == payload.highlight_row_index)
        is_even = (row_idx % 2 == 0)

        # Alternating background or highlight accent
        if is_highlighted:
            bg_col = SlideMaster.COLORS["fallback_bg"]
            text_col = SlideMaster.COLORS["primary"]
        elif is_even:
            bg_col = SlideMaster.COLORS["bg_white"]
            text_col = SlideMaster.COLORS["text_dark"]
        else:
            bg_col = SlideMaster.COLORS["bg_light"]
            text_col = SlideMaster.COLORS["text_dark"]

        for col_idx in range(num_cols):
            cell = table.cell(row_idx + 1, col_idx)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_col

            cell_text = row_data[col_idx] if col_idx < len(row_data) else ""
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = cell_text
            p.font.name = SlideMaster.FONT_FAMILY_PRIMARY
            p.font.size = FontEngine.pt(10.0)
            p.font.bold = is_highlighted
            p.font.color.rgb = text_col