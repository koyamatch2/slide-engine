"""Renderer for Comparison slide (CORE-70).

Renders two contrasting columns (Before/After, Pros/Cons, Current/Target)
with distinct headers, bullets, and optional highlight accents. Automatically
switches to vertical top/bottom stacking inside narrow split panes.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.util import Inches, Length

from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import ComparisonColumn, Core70Payload

if TYPE_CHECKING:
    from pptx.slide import Slide


COL_GAP = Inches(0.250)
COL_INNER_PAD = Inches(0.240)
MIN_SIDE_BY_SIDE_WIDTH = Inches(4.200)


def _render_comparison_column(
    slide: Slide,
    column: ComparisonColumn,
    box: BoundingBox,
) -> None:
    """Render a single comparison column card with adaptive typography."""
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

    pad_x = Length(min(int(COL_INNER_PAD), int(box.width * 0.08)))
    pad_y = Length(min(int(COL_INNER_PAD), int(box.height * 0.08)))

    inner_box = BoundingBox(
        left=Length(box.left + pad_x),
        top=Length(box.top + pad_y),
        width=Length(max(int(Inches(0.400)), int(box.width - (pad_x * 2)))),
        height=Length(max(int(Inches(0.300)), int(box.height - (pad_y * 2)))),
    )
    tb = ShapeBuilder.add_textbox(slide, inner_box)
    tf = tb.text_frame

    header_budget_h = Length(int(inner_box.height * 0.18))
    sub_budget_h = Length(int(inner_box.height * 0.12))
    points_Count = max(1, len(column.points))
    point_budget_h = Length(int((inner_box.height * 0.65) / points_Count))
    space_gap_pt = min(6.0, max(1.5, inner_box.height.pt * 0.02))

    # Header Title
    p_header = tf.paragraphs[0]
    p_header.text = column.header
    p_header.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_header.font.size = FontEngine.calculate_font_size(
        column.header,
        inner_box.width,
        header_budget_h,
        max_pt=15.0,
        min_pt=9.0,
    )
    p_header.font.bold = True
    p_header.font.color.rgb = (
        SlideMaster.COLORS["accent"] if column.is_highlighted else SlideMaster.COLORS["primary"]
    )

    # Sub-header
    if column.sub_header:
        p_sub = tf.add_paragraph()
        p_sub.text = column.sub_header
        p_sub.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p_sub.font.size = FontEngine.calculate_font_size(
            column.sub_header,
            inner_box.width,
            sub_budget_h,
            max_pt=10.5,
            min_pt=8.0,
        )
        p_sub.font.color.rgb = SlideMaster.COLORS["text_muted"]
        p_sub.space_after = FontEngine.pt(space_gap_pt * 1.2)
    else:
        p_header.space_after = FontEngine.pt(space_gap_pt * 1.2)

    # Point Bullet List
    for point in column.points:
        bullet_text = f"•  {point}"
        p_pt = tf.add_paragraph()
        p_pt.text = bullet_text
        p_pt.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p_pt.font.size = FontEngine.calculate_font_size(
            bullet_text,
            inner_box.width,
            point_budget_h,
            max_pt=11.0,
            min_pt=8.0,
        )
        p_pt.font.color.rgb = SlideMaster.COLORS["text_dark"]
        p_pt.space_after = FontEngine.pt(space_gap_pt * 0.6)


def render_core_70(slide: Slide, payload: Core70Payload, box: BoundingBox) -> None:
    """Render dual-column or vertically stacked comparison boxes.

    Args:
        slide: Target slide.
        payload: Core70Payload instance.
        box: Content BoundingBox.
    """
    if box.width >= MIN_SIDE_BY_SIDE_WIDTH:
        gap = Length(min(int(COL_GAP), int(box.width * 0.05)))
        col_width = Length(int((box.width - gap) / 2))
        first_box = BoundingBox(box.left, box.top, col_width, box.height)
        second_box = BoundingBox(Length(box.left + col_width + gap), box.top, col_width, box.height)
    else:
        gap = Length(min(int(COL_GAP), int(box.height * 0.05)))
        row_height = Length(int((box.height - gap) / 2))
        first_box = BoundingBox(box.left, box.top, box.width, row_height)
        second_box = BoundingBox(box.left, Length(box.top + row_height + gap), box.width, row_height)

    _render_comparison_column(slide, payload.left_column, first_box)
    _render_comparison_column(slide, payload.right_column, second_box)
