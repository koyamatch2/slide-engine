"""Renderer for Process Flow slide (CORE-20).

Renders sequential timeline or procedural steps with badges,
connector indicators, and cards. Automatically adapts from single-row
horizontal flow to multi-row grid when rendered inside narrow split panes.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.util import Inches, Length

from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import Core20Payload, ProcessStep

if TYPE_CHECKING:
    from pptx.slide import Slide


# Process Flow Layout Constants
GAP_BETWEEN_STEPS = Inches(0.180)
CARD_PADDING_X = Inches(0.180)
CARD_PADDING_Y = Inches(0.150)
BADGE_HEIGHT = Inches(0.320)
MIN_STEP_COL_WIDTH = Inches(1.650)
COMPACT_CARD_HEIGHT_THRESHOLD = Inches(1.300)


def _render_compact_step_card(
    slide: Slide,
    step: ProcessStep,
    step_box: BoundingBox,
    pad_x: Length,
    pad_y: Length,
) -> None:
    """Render a compact step card with inline badge header to prevent shape height collapse."""
    inner_box = BoundingBox(
        left=Length(step_box.left + pad_x),
        top=Length(step_box.top + pad_y),
        width=Length(max(int(Inches(0.400)), int(step_box.width - (pad_x * 2)))),
        height=Length(max(int(Inches(0.300)), int(step_box.height - (pad_y * 2)))),
    )
    tb = ShapeBuilder.add_textbox(slide, inner_box)
    tf = tb.text_frame

    title_budget_h = Length(int(inner_box.height * 0.45))
    desc_budget_h = Length(int(inner_box.height * 0.50))

    p_title = tf.paragraphs[0]
    badge_label = f"STEP {step.step_number:02d}"
    full_title = f"[{badge_label}] {step.title}"
    p_title.text = full_title
    p_title.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_title.font.size = FontEngine.calculate_font_size(
        full_title,
        inner_box.width,
        title_budget_h,
        max_pt=11.5,
        min_pt=8.0,
    )
    p_title.font.bold = True
    p_title.font.color.rgb = SlideMaster.COLORS["primary"]
    p_title.space_after = FontEngine.pt(2.5)

    p_desc = tf.add_paragraph()
    p_desc.text = step.description
    p_desc.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_desc.font.size = FontEngine.calculate_font_size(
        step.description,
        inner_box.width,
        desc_budget_h,
        max_pt=9.5,
        min_pt=8.0,
    )
    p_desc.font.color.rgb = SlideMaster.COLORS["text_dark"]


def _render_standard_step_card(
    slide: Slide,
    step: ProcessStep,
    step_box: BoundingBox,
    pad_x: Length,
    pad_y: Length,
) -> None:
    """Render a standard step card with dedicated badge box."""
    inner_width = Length(max(int(Inches(0.400)), int(step_box.width - (pad_x * 2))))
    badge_h = Length(min(int(BADGE_HEIGHT), int(step_box.height * 0.22)))
    inner_gap = Length(min(int(Inches(0.060)), int(step_box.height * 0.04)))

    # Number Badge Header Box
    badge_box = BoundingBox(
        left=Length(step_box.left + pad_x),
        top=Length(step_box.top + pad_y),
        width=inner_width,
        height=badge_h,
    )
    tb_badge = ShapeBuilder.add_textbox(slide, badge_box)
    p_badge = tb_badge.text_frame.paragraphs[0]
    badge_text = f"STEP {step.step_number:02d}"
    p_badge.text = badge_text
    p_badge.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_badge.font.size = FontEngine.calculate_font_size(
        badge_text,
        badge_box.width,
        badge_box.height,
        max_pt=11.0,
        min_pt=8.0,
    )
    p_badge.font.bold = True
    p_badge.font.color.rgb = SlideMaster.COLORS["secondary"]

    # Step Title & Description Box
    content_top = Length(step_box.top + pad_y + badge_h + inner_gap)
    content_height = Length(max(int(Inches(0.300)), int(step_box.bottom - content_top - pad_y)))
    content_box = BoundingBox(
        left=Length(step_box.left + pad_x),
        top=content_top,
        width=inner_width,
        height=content_height,
    )
    tb_content = ShapeBuilder.add_textbox(slide, content_box)
    tf_content = tb_content.text_frame

    title_avail_h = Length(int(content_height * 0.35))
    desc_avail_h = Length(int(content_height * 0.60))

    p_title = tf_content.paragraphs[0]
    p_title.text = step.title
    p_title.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_title.font.size = FontEngine.calculate_font_size(
        step.title,
        content_box.width,
        title_avail_h,
        max_pt=12.5,
        min_pt=8.5,
    )
    p_title.font.bold = True
    p_title.font.color.rgb = SlideMaster.COLORS["primary"]
    p_title.space_after = FontEngine.pt(4.0)

    p_desc = tf_content.add_paragraph()
    p_desc.text = step.description
    p_desc.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_desc.font.size = FontEngine.calculate_font_size(
        step.description,
        content_box.width,
        desc_avail_h,
        max_pt=10.0,
        min_pt=8.0,
    )
    p_desc.font.color.rgb = SlideMaster.COLORS["text_dark"]


def _render_process_step_card(
    slide: Slide,
    step: ProcessStep,
    step_box: BoundingBox,
) -> None:
    """Render an individual process step card and its typography."""
    ShapeBuilder.add_card(
        slide,
        step_box,
        bg_color=SlideMaster.COLORS["bg_light"],
        border_color=SlideMaster.COLORS["border"],
        border_width_pt=1.0,
    )

    pad_x = Length(min(int(CARD_PADDING_X), int(step_box.width * 0.08)))
    pad_y = Length(min(int(CARD_PADDING_Y), int(step_box.height * 0.08)))

    if step_box.height < COMPACT_CARD_HEIGHT_THRESHOLD:
        _render_compact_step_card(slide, step, step_box, pad_x, pad_y)
    else:
        _render_standard_step_card(slide, step, step_box, pad_x, pad_y)


def render_core_20(slide: Slide, payload: Core20Payload, box: BoundingBox) -> None:
    """Render sequential steps using adaptive row/column grid layout.

    Args:
        slide: Target slide.
        payload: Core20Payload instance.
        box: Content BoundingBox.
    """
    num_steps = len(payload.steps)
    cols = max(1, min(num_steps, int(box.width / MIN_STEP_COL_WIDTH)))
    rows = -(-num_steps // cols)

    gap_x = Length(min(int(GAP_BETWEEN_STEPS), int(box.width * 0.03)))
    gap_y = Length(min(int(GAP_BETWEEN_STEPS), int(box.height * 0.04)))

    step_width = Length(int((box.width - gap_x * (cols - 1)) / cols))
    step_height = Length(int((box.height - gap_y * (rows - 1)) / rows))

    for idx, step in enumerate(payload.steps):
        r_idx = idx // cols
        c_idx = idx % cols
        step_box = BoundingBox(
            left=Length(box.left + c_idx * (step_width + gap_x)),
            top=Length(box.top + r_idx * (step_height + gap_y)),
            width=step_width,
            height=step_height,
        )
        _render_process_step_card(slide, step, step_box)
