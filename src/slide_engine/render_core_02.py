"""Renderer for Agenda and Section Divider slide (CORE-02).

Renders numbered agenda items with dynamic vertical distribution
and active section highlighting.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import Core02Payload

if TYPE_CHECKING:
    from pptx.slide import Slide


# Agenda Geometry Constants
BADGE_WIDTH = Inches(0.850)
GAP_BETWEEN_ROWS = Inches(0.200)
PADDING_CONTENT = Inches(0.250)


def render_core_02(slide: Slide, payload: Core02Payload, box: BoundingBox) -> None:
    """Render numbered agenda items with active highlighting.

    Args:
        slide: Target slide.
        payload: Core02Payload instance.
        box: Content BoundingBox.
    """
    num_items = len(payload.items)
    total_gaps = GAP_BETWEEN_ROWS * (num_items - 1)
    available_height = box.height - total_gaps
    row_height = Inches(available_height / num_items)

    current_top = box.top

    for item in payload.items:
        row_box = BoundingBox(
            left=box.left,
            top=current_top,
            width=box.width,
            height=row_height,
        )

        # Container card background based on active state
        bg_col = SlideMaster.COLORS["primary"] if item.is_active else SlideMaster.COLORS["bg_light"]
        border_col = SlideMaster.COLORS["primary"] if item.is_active else SlideMaster.COLORS["border"]
        ShapeBuilder.add_card(slide, row_box, bg_color=bg_col, border_color=border_col, border_width_pt=1.0)

        # 1. Number Badge
        badge_box = BoundingBox(
            left=Inches(row_box.left + PADDING_CONTENT),
            top=Inches(row_box.top + (row_height - Inches(0.500)) / 2),
            width=BADGE_WIDTH,
            height=Inches(0.500),
        )
        tb_badge = ShapeBuilder.add_textbox(slide, badge_box)
        p_badge = tb_badge.text_frame.paragraphs[0]
        p_badge.text = str(item.number)
        p_badge.alignment = PP_ALIGN.CENTER
        p_badge.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p_badge.font.size = FontEngine.pt(18.0)
        p_badge.font.bold = True
        p_badge.font.color.rgb = SlideMaster.COLORS["accent"] if item.is_active else SlideMaster.COLORS["secondary"]

        # 2. Title & Description Box
        content_left = Inches(row_box.left + PADDING_CONTENT + BADGE_WIDTH + Inches(0.200))
        content_width = Inches(row_box.width - (content_left - row_box.left) - PADDING_CONTENT)
        content_box = BoundingBox(
            left=content_left,
            top=Inches(row_box.top + Inches(0.120)),
            width=content_width,
            height=Inches(row_height - Inches(0.240)),
        )
        tb_content = ShapeBuilder.add_textbox(slide, content_box)
        tf_content = tb_content.text_frame

        p_title = tf_content.paragraphs[0]
        p_title.text = item.title
        p_title.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p_title.font.size = FontEngine.pt(14.0)
        p_title.font.bold = True
        p_title.font.color.rgb = SlideMaster.COLORS["text_light"] if item.is_active else SlideMaster.COLORS["text_dark"]

        if item.description:
            p_desc = tf_content.add_paragraph()
            p_desc.text = item.description
            p_desc.font.name = SlideMaster.FONT_FAMILY_PRIMARY
            p_desc.font.size = FontEngine.pt(10.0)
            p_desc.font.color.rgb = SlideMaster.COLORS["bg_light"] if item.is_active else SlideMaster.COLORS["text_muted"]

        current_top = Inches(current_top + row_height + GAP_BETWEEN_ROWS)