"""Renderer for KPI Summary Cards slide (CORE-60).

Renders high-impact metric cards with primary numbers, units,
descriptive labels, and positive/negative trend badges.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.util import Inches

from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import Core60Payload

if TYPE_CHECKING:
    from pptx.slide import Slide


CARD_GAP = Inches(0.250)
CARD_INNER_PAD = Inches(0.250)


def render_core_60(slide: Slide, payload: Core60Payload, box: BoundingBox) -> None:
    """Render horizontal KPI cards dynamically.

    Args:
        slide: Target slide.
        payload: Core60Payload instance.
        box: Content BoundingBox.
    """
    num_cards = len(payload.metrics)
    total_gaps = CARD_GAP * (num_cards - 1)
    available_width = box.width - total_gaps
    card_width = Inches(available_width / num_cards)

    current_left = box.left

    for metric in payload.metrics:
        card_box = BoundingBox(
            left=current_left,
            top=box.top,
            width=card_width,
            height=box.height,
        )

        # 1. Card Container
        ShapeBuilder.add_card(
            slide,
            card_box,
            bg_color=SlideMaster.COLORS["bg_white"],
            border_color=SlideMaster.COLORS["border"],
            border_width_pt=1.0,
        )

        # 2. Content Box
        inner_box = BoundingBox(
            left=Inches(card_box.left + CARD_INNER_PAD),
            top=Inches(card_box.top + CARD_INNER_PAD),
            width=Inches(card_box.width - (CARD_INNER_PAD * 2)),
            height=Inches(card_box.height - (CARD_INNER_PAD * 2)),
        )
        tb = ShapeBuilder.add_textbox(slide, inner_box)
        tf = tb.text_frame

        # Metric Label
        p_label = tf.paragraphs[0]
        p_label.text = metric.label.upper()
        p_label.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p_label.font.size = FontEngine.pt(11.0)
        p_label.font.bold = True
        p_label.font.color.rgb = SlideMaster.COLORS["text_muted"]
        p_label.space_after = FontEngine.pt(16.0)

        # Big Number Value + Unit
        p_val = tf.add_paragraph()
        p_val.text = metric.value
        p_val.font.name = SlideMaster.FONT_FAMILY_SECONDARY
        p_val.font.size = FontEngine.pt(40.0)
        p_val.font.bold = True
        p_val.font.color.rgb = SlideMaster.COLORS["primary"]

        if metric.unit:
            run_unit = p_val.add_run()
            run_unit.text = f" {metric.unit}"
            run_unit.font.name = SlideMaster.FONT_FAMILY_PRIMARY
            run_unit.font.size = FontEngine.pt(16.0)
            run_unit.font.bold = True
            run_unit.font.color.rgb = SlideMaster.COLORS["secondary"]

        p_val.space_after = FontEngine.pt(12.0)

        # Trend Indicator
        if metric.trend:
            p_trend = tf.add_paragraph()
            p_trend.text = metric.trend
            p_trend.font.name = SlideMaster.FONT_FAMILY_PRIMARY
            p_trend.font.size = FontEngine.pt(10.5)
            p_trend.font.bold = True

            if metric.is_positive is True:
                p_trend.font.color.rgb = SlideMaster.COLORS["accent"]
            elif metric.is_positive is False:
                p_trend.font.color.rgb = SlideMaster.COLORS["fallback_text"]
            else:
                p_trend.font.color.rgb = SlideMaster.COLORS["text_muted"]

        current_left = Inches(current_left + card_width + CARD_GAP)