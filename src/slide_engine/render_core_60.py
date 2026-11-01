"""Renderer for KPI Summary Cards slide (CORE-60).

Renders high-impact metric cards with primary numbers, units,
descriptive labels, and positive/negative trend badges. Automatically
switches to a 2x2 or vertical stack layout inside narrow split panes.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.util import Inches, Length

from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import Core60Payload, KpiMetric

if TYPE_CHECKING:
    from pptx.slide import Slide


CARD_GAP = Inches(0.200)
CARD_INNER_PAD = Inches(0.220)
MIN_KPI_COL_WIDTH = Inches(1.850)


def _render_kpi_card(
    slide: Slide,
    metric: KpiMetric,
    card_box: BoundingBox,
) -> None:
    """Render an individual KPI card with dynamically scaled typography."""
    ShapeBuilder.add_card(
        slide,
        card_box,
        bg_color=SlideMaster.COLORS["bg_white"],
        border_color=SlideMaster.COLORS["border"],
        border_width_pt=1.0,
    )

    pad_x = Length(min(int(CARD_INNER_PAD), int(card_box.width * 0.08)))
    pad_y = Length(min(int(CARD_INNER_PAD), int(card_box.height * 0.10)))

    inner_box = BoundingBox(
        left=Length(card_box.left + pad_x),
        top=Length(card_box.top + pad_y),
        width=Length(max(int(Inches(0.400)), int(card_box.width - (pad_x * 2)))),
        height=Length(max(int(Inches(0.300)), int(card_box.height - (pad_y * 2)))),
    )
    tb = ShapeBuilder.add_textbox(slide, inner_box)
    tf = tb.text_frame

    label_avail_h = Length(int(inner_box.height * 0.22))
    val_avail_h = Length(int(inner_box.height * 0.45))
    trend_avail_h = Length(int(inner_box.height * 0.22))
    space_gap_pt = min(10.0, max(2.0, inner_box.height.pt * 0.04))

    # Metric Label
    p_label = tf.paragraphs[0]
    label_text = metric.label.upper()
    p_label.text = label_text
    p_label.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_label.font.size = FontEngine.calculate_font_size(
        label_text,
        inner_box.width,
        label_avail_h,
        max_pt=11.0,
        min_pt=8.0,
    )
    p_label.font.bold = True
    p_label.font.color.rgb = SlideMaster.COLORS["text_muted"]
    p_label.space_after = FontEngine.pt(space_gap_pt)

    # Big Number Value + Unit
    full_val_text = f"{metric.value} {metric.unit}" if metric.unit else metric.value
    val_font_len = FontEngine.calculate_font_size(
        full_val_text,
        inner_box.width,
        val_avail_h,
        max_pt=38.0,
        min_pt=13.0,
    )

    p_val = tf.add_paragraph()
    p_val.text = metric.value
    p_val.font.name = SlideMaster.FONT_FAMILY_SECONDARY
    p_val.font.size = val_font_len
    p_val.font.bold = True
    p_val.font.color.rgb = SlideMaster.COLORS["primary"]

    if metric.unit:
        unit_pt = max(9.0, min(16.0, val_font_len.pt * 0.45))
        run_unit = p_val.add_run()
        run_unit.text = f" {metric.unit}"
        run_unit.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        run_unit.font.size = FontEngine.pt(unit_pt)
        run_unit.font.bold = True
        run_unit.font.color.rgb = SlideMaster.COLORS["secondary"]

    p_val.space_after = FontEngine.pt(space_gap_pt * 0.8)

    # Trend Indicator
    if metric.trend:
        p_trend = tf.add_paragraph()
        p_trend.text = metric.trend
        p_trend.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p_trend.font.size = FontEngine.calculate_font_size(
            metric.trend,
            inner_box.width,
            trend_avail_h,
            max_pt=10.5,
            min_pt=8.0,
        )
        p_trend.font.bold = True
        if metric.is_positive is True:
            p_trend.font.color.rgb = SlideMaster.COLORS["secondary"]
        elif metric.is_positive is False:
            p_trend.font.color.rgb = SlideMaster.COLORS["accent"]
        else:
            p_trend.font.color.rgb = SlideMaster.COLORS["text_muted"]


def render_core_60(slide: Slide, payload: Core60Payload, box: BoundingBox) -> None:
    """Render KPI cards dynamically using single-row or grid layout.

    Args:
        slide: Target slide.
        payload: Core60Payload instance.
        box: Content BoundingBox.
    """
    num_cards = len(payload.metrics)
    cols = max(1, min(num_cards, int(box.width / MIN_KPI_COL_WIDTH)))
    if cols == 3 and num_cards == 4:
        cols = 2
    rows = -(-num_cards // cols)

    gap_x = Length(min(int(CARD_GAP), int(box.width * 0.03)))
    gap_y = Length(min(int(CARD_GAP), int(box.height * 0.04)))

    card_width = Length(int((box.width - gap_x * (cols - 1)) / cols))
    card_height = Length(int((box.height - gap_y * (rows - 1)) / rows))

    for idx, metric in enumerate(payload.metrics):
        r_idx = idx // cols
        c_idx = idx % cols
        card_box = BoundingBox(
            left=Length(box.left + c_idx * (card_width + gap_x)),
            top=Length(box.top + r_idx * (card_height + gap_y)),
            width=card_width,
            height=card_height,
        )
        _render_kpi_card(slide, metric, card_box)
