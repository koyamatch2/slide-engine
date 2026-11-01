"""Renderer for KPI Summary Cards slide (CORE-60).

Renders high-impact metric cards with primary numbers, units,
descriptive labels, and positive/negative trend badges. Automatically
switches to a 2x2 or vertical stack layout inside narrow split panes.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.util import Inches
from pydantic import Field

from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import BasePayload, ShortText, TemplateSpec

if TYPE_CHECKING:
    from pptx.slide import Slide


class KpiMetric(BasePayload):
    """Metric item for KPI summary card."""

    value: ShortText
    unit: ShortText | None = None
    label: ShortText
    trend: ShortText | None = None
    is_positive: bool | None = None


class Core60Payload(BasePayload):
    """Payload for KPI Summary Cards slide (CORE-60)."""

    metrics: list[KpiMetric] = Field(min_length=2, max_length=4)


_SAMPLE_CORE60: dict[str, Any] = {
    "metrics": [
        {
            "label": f"重要業績評価指標 {i}",
            "value": f"{i * 128.5}",
            "unit": "億円",
            "trend": f"前年同期比 +{i * 6.4}% 達成",
            "is_positive": True,
        }
        for i in range(1, 5)
    ]
}

_INVALID_SAMPLES_CORE60: list[dict[str, Any]] = [
    {"metrics": [{"label": "A", "value": "1"}]},  # Violates min_length=2
]

SPEC = TemplateSpec(
    payload_type=Core60Payload,
    supports_split=True,
    description="KPI Summary Cards",
    sample_payload=_SAMPLE_CORE60,
    invalid_sample_payloads=_INVALID_SAMPLES_CORE60,
)


CARD_GAP = Inches(0.200)
CARD_INNER_PAD = Inches(0.220)
MIN_KPI_COL_WIDTH = Inches(1.850)


def _resolve_trend_color(is_positive: bool | None) -> str:
    """Resolve trend badge color key based on polarity."""
    if is_positive is True:
        return "secondary"
    if is_positive is False:
        return "accent"
    return "text_muted"


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

    inner_box = card_box.inset(CARD_INNER_PAD, CARD_INNER_PAD, max_ratio_x=0.08, max_ratio_y=0.10)
    tb = ShapeBuilder.add_textbox(slide, inner_box)
    tf = tb.text_frame

    label_avail_h = inner_box.scale_height(0.22)
    val_avail_h = inner_box.scale_height(0.45)
    trend_avail_h = inner_box.scale_height(0.22)
    space_gap_pt = min(10.0, max(2.0, inner_box.height.pt * 0.04))

    # Metric Label
    ShapeBuilder.add_styled_paragraph(
        tf,
        metric.label.upper(),
        inner_box.width,
        label_avail_h,
        max_pt=11.0,
        min_pt=8.0,
        color_key="text_muted",
        bold=True,
        space_after_pt=space_gap_pt,
        is_first=True,
    )

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
        ShapeBuilder.add_styled_paragraph(
            tf,
            metric.trend,
            inner_box.width,
            trend_avail_h,
            max_pt=10.5,
            min_pt=8.0,
            color_key=_resolve_trend_color(metric.is_positive),
            bold=True,
        )


def render_core_60(slide: Slide, payload: Core60Payload, box: BoundingBox) -> None:
    """Render KPI cards dynamically using single-row or grid layout.

    Args:
        slide: Target slide.
        payload: Core60Payload instance.
        box: Content BoundingBox.
    """
    grid = box.adaptive_grid(
        num_items=len(payload.metrics),
        min_col_width=MIN_KPI_COL_WIDTH,
        gap_x=CARD_GAP,
        gap_y=CARD_GAP,
        balance_rows=True,
    )

    for metric, card_box in zip(payload.metrics, grid, strict=False):
        _render_kpi_card(slide, metric, card_box)
