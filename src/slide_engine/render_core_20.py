"""Renderer for Process Flow slide (CORE-20).

Renders sequential timeline or procedural steps with badges,
connector indicators, and cards.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.util import Inches

from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import Core20Payload

if TYPE_CHECKING:
    from pptx.slide import Slide


# Process Flow Layout Constants
GAP_BETWEEN_STEPS = Inches(0.200)
CARD_PADDING_X = Inches(0.200)
CARD_PADDING_Y = Inches(0.200)
BADGE_HEIGHT = Inches(0.400)


def render_core_20(slide: Slide, payload: Core20Payload, box: BoundingBox) -> None:
    """Render horizontal sequential steps.

    Args:
        slide: Target slide.
        payload: Core20Payload instance.
        box: Content BoundingBox.
    """
    num_steps = len(payload.steps)
    total_gaps = GAP_BETWEEN_STEPS * (num_steps - 1)
    available_width = box.width - total_gaps
    step_width = Inches(available_width / num_steps)

    current_left = box.left

    for step in payload.steps:
        step_box = BoundingBox(
            left=current_left,
            top=box.top,
            width=step_width,
            height=box.height,
        )

        # 1. Step Card Container
        ShapeBuilder.add_card(
            slide,
            step_box,
            bg_color=SlideMaster.COLORS["bg_light"],
            border_color=SlideMaster.COLORS["border"],
            border_width_pt=1.0,
        )

        # 2. Number Badge Header Box
        badge_box = BoundingBox(
            left=Inches(step_box.left + CARD_PADDING_X),
            top=Inches(step_box.top + CARD_PADDING_Y),
            width=Inches(step_box.width - (CARD_PADDING_X * 2)),
            height=BADGE_HEIGHT,
        )
        tb_badge = ShapeBuilder.add_textbox(slide, badge_box)
        p_badge = tb_badge.text_frame.paragraphs[0]
        p_badge.text = f"STEP {step.step_number:02d}"
        p_badge.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p_badge.font.size = FontEngine.pt(12.0)
        p_badge.font.bold = True
        p_badge.font.color.rgb = SlideMaster.COLORS["secondary"]

        # 3. Step Title & Description Box
        content_box = BoundingBox(
            left=Inches(step_box.left + CARD_PADDING_X),
            top=Inches(step_box.top + CARD_PADDING_Y + BADGE_HEIGHT + Inches(0.100)),
            width=Inches(step_box.width - (CARD_PADDING_X * 2)),
            height=Inches(step_box.height - BADGE_HEIGHT - (CARD_PADDING_Y * 2) - Inches(0.100)),
        )
        tb_content = ShapeBuilder.add_textbox(slide, content_box)
        tf_content = tb_content.text_frame

        p_title = tf_content.paragraphs[0]
        p_title.text = step.title
        p_title.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p_title.font.size = FontEngine.pt(13.0)
        p_title.font.bold = True
        p_title.font.color.rgb = SlideMaster.COLORS["primary"]
        p_title.space_after = FontEngine.pt(8.0)

        p_desc = tf_content.add_paragraph()
        p_desc.text = step.description
        p_desc.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p_desc.font.size = FontEngine.pt(10.0)
        p_desc.font.color.rgb = SlideMaster.COLORS["text_dark"]

        current_left = Inches(current_left + step_width + GAP_BETWEEN_STEPS)