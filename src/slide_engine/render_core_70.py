"""Renderer for Comparison slide (CORE-70).

Renders two contrasting columns (Before/After, Pros/Cons, Current/Target)
with distinct headers, bullets, and optional highlight accents.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.util import Inches

from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import ComparisonColumn, Core70Payload

if TYPE_CHECKING:
    from pptx.slide import Slide


COL_GAP = Inches(0.400)
COL_INNER_PAD = Inches(0.300)


def _render_comparison_column(
    slide: Slide,
    column: ComparisonColumn,
    box: BoundingBox,
) -> None:
    """Render a single comparison column card."""
    # Determine highlight background and border
    if column.is_highlighted:
        bg_color = SlideMaster.COLORS["bg_white"]
        border_color = SlideMaster.COLORS["accent"]
        border_width = 1.5
    else:
        bg_color = SlideMaster.COLORS["bg_light"]
        border_color = SlideMaster.COLORS["border"]
        border_width = 1.0

    ShapeBuilder.add_card(
        slide,
        box,
        bg_color=bg_color,
        border_color=border_color,
        border_width_pt=border_width,
    )

    inner_box = BoundingBox(
        left=Inches(box.left + COL_INNER_PAD),
        top=Inches(box.top + COL_INNER_PAD),
        width=Inches(box.width - (COL_INNER_PAD * 2)),
        height=Inches(box.height - (COL_INNER_PAD * 2)),
    )
    tb = ShapeBuilder.add_textbox(slide, inner_box)
    tf = tb.text_frame

    # Header Title
    p_header = tf.paragraphs[0]
    p_header.text = column.header
    p_header.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_header.font.size = FontEngine.pt(16.0)
    p_header.font.bold = True
    p_header.font.color.rgb = SlideMaster.COLORS["accent"] if column.is_highlighted else SlideMaster.COLORS["primary"]

    # Sub-header
    if column.sub_header:
        p_sub = tf.add_paragraph()
        p_sub.text = column.sub_header
        p_sub.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p_sub.font.size = FontEngine.pt(10.5)
        p_sub.font.color.rgb = SlideMaster.COLORS["text_muted"]
        p_sub.space_after = FontEngine.pt(12.0)
    else:
        p_header.space_after = FontEngine.pt(12.0)

    # Point Bullet List
    for point in column.points:
        p_pt = tf.add_paragraph()
        p_pt.text = f"•  {point}"
        p_pt.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p_pt.font.size = FontEngine.pt(11.0)
        p_pt.font.color.rgb = SlideMaster.COLORS["text_dark"]
        p_pt.space_after = FontEngine.pt(6.0)


def render_core_70(slide: Slide, payload: Core70Payload, box: BoundingBox) -> None:
    """Render dual-column comparison boxes.

    Args:
        slide: Target slide.
        payload: Core70Payload instance.
        box: Content BoundingBox.
    """
    col_width = Inches((box.width - COL_GAP) / 2)

    left_box = BoundingBox(
        left=box.left,
        top=box.top,
        width=col_width,
        height=box.height,
    )
    right_box = BoundingBox(
        left=Inches(box.left + col_width + COL_GAP),
        top=box.top,
        width=col_width,
        height=box.height,
    )

    _render_comparison_column(slide, payload.left_column, left_box)
    _render_comparison_column(slide, payload.right_column, right_box)