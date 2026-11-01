"""Renderer for 2x2 Matrix slide (CORE-40).

Renders a four-quadrant positioning map with custom X/Y axis labels,
distinct card backgrounds, and category badges.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Length

from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import Core40Payload, MatrixQuadrant

if TYPE_CHECKING:
    from pptx.slide import Slide


# Matrix Geometry Constants
AXIS_LABEL_SIZE = Inches(0.350)
GRID_GAP = Inches(0.180)
QUADRANT_PADDING = Inches(0.180)


def _render_quadrant_card(
    slide: Slide,
    quadrant: MatrixQuadrant,
    box: BoundingBox,
) -> None:
    """Render a single quadrant card with title, badge, and description."""
    ShapeBuilder.add_card(
        slide,
        box,
        bg_color=SlideMaster.COLORS["bg_light"],
        border_color=SlideMaster.COLORS["border"],
        border_width_pt=1.0,
    )

    pad_x = Length(min(int(QUADRANT_PADDING), int(box.width * 0.08)))
    pad_y = Length(min(int(QUADRANT_PADDING), int(box.height * 0.08)))

    inner_box = BoundingBox(
        left=Length(box.left + pad_x),
        top=Length(box.top + pad_y),
        width=Length(max(int(Inches(0.400)), int(box.width - (pad_x * 2)))),
        height=Length(max(int(Inches(0.300)), int(box.height - (pad_y * 2)))),
    )
    tb = ShapeBuilder.add_textbox(slide, inner_box)
    tf = tb.text_frame

    header_avail_h = Length(int(inner_box.height * 0.35))
    desc_avail_h = Length(int(inner_box.height * 0.60))

    # Quadrant Title and Optional Badge
    p_header = tf.paragraphs[0]
    header_text = quadrant.title
    if quadrant.badge:
        header_text = f"[{quadrant.badge}] {header_text}"
    p_header.text = header_text
    p_header.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_header.font.size = FontEngine.calculate_font_size(
        header_text,
        inner_box.width,
        header_avail_h,
        max_pt=12.5,
        min_pt=8.5,
    )
    p_header.font.bold = True
    p_header.font.color.rgb = SlideMaster.COLORS["primary"]
    p_header.space_after = FontEngine.pt(4.0)

    # Description
    p_desc = tf.add_paragraph()
    p_desc.text = quadrant.description
    p_desc.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_desc.font.size = FontEngine.calculate_font_size(
        quadrant.description,
        inner_box.width,
        desc_avail_h,
        max_pt=10.0,
        min_pt=8.0,
    )
    p_desc.font.color.rgb = SlideMaster.COLORS["text_dark"]


def render_core_40(slide: Slide, payload: Core40Payload, box: BoundingBox) -> None:
    """Render 2x2 matrix with axis headers.

    Args:
        slide: Target slide.
        payload: Core40Payload instance.
        box: Content BoundingBox.
    """
    # 1. Axis Label Zones
    axis_h = Length(min(int(AXIS_LABEL_SIZE), int(box.height * 0.10)))
    tb_x = ShapeBuilder.add_textbox(
        slide,
        BoundingBox(box.left, box.top, box.width, axis_h),
    )
    p_x = tb_x.text_frame.paragraphs[0]
    axis_text = f"◀ {payload.x_axis_label}  |  {payload.y_axis_label} ▶"
    p_x.text = axis_text
    p_x.alignment = PP_ALIGN.CENTER
    p_x.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_x.font.size = FontEngine.calculate_font_size(
        axis_text,
        box.width,
        axis_h,
        max_pt=11.0,
        min_pt=8.0,
    )
    p_x.font.bold = True
    p_x.font.color.rgb = SlideMaster.COLORS["secondary"]

    # 2. Quadrants Layout Calculation
    gap_x = Length(min(int(GRID_GAP), int(box.width * 0.04)))
    gap_y = Length(min(int(GRID_GAP), int(box.height * 0.04)))
    top_offset = Length(min(int(Inches(0.080)), int(box.height * 0.02)))

    quad_top = Length(box.top + axis_h + top_offset)
    quad_height = Length(int((box.height - axis_h - top_offset - gap_y) / 2))
    quad_width = Length(int((box.width - gap_x) / 2))

    col1_left = box.left
    col2_left = Length(box.left + quad_width + gap_x)
    row1_top = quad_top
    row2_top = Length(quad_top + quad_height + gap_y)

    # Top-Left
    _render_quadrant_card(
        slide,
        payload.top_left,
        BoundingBox(col1_left, row1_top, quad_width, quad_height),
    )
    # Top-Right
    _render_quadrant_card(
        slide,
        payload.top_right,
        BoundingBox(col2_left, row1_top, quad_width, quad_height),
    )
    # Bottom-Left
    _render_quadrant_card(
        slide,
        payload.bottom_left,
        BoundingBox(col1_left, row2_top, quad_width, quad_height),
    )
    # Bottom-Right
    _render_quadrant_card(
        slide,
        payload.bottom_right,
        BoundingBox(col2_left, row2_top, quad_width, quad_height),
    )
