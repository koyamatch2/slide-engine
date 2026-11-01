"""Renderer for Waterfall Chart slide (CORE-80).

Renders a coordinate-based waterfall chart using standard shapes,
bypassing the limitations of native OpenXML waterfall support.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Length
from pydantic import Field

from slide_engine.core_base import BoundingBox, ShapeBuilder, SlideMaster
from slide_engine.models import BasePayload, ShortText, TemplateSpec

if TYPE_CHECKING:
    from pptx.slide import Slide


class WaterfallStep(BasePayload):
    """Single step in a Waterfall Chart."""

    label: ShortText
    value: float
    is_total: bool = False


class Core80Payload(BasePayload):
    """Payload for Waterfall Chart slide (CORE-80)."""

    unit: ShortText | None = None
    steps: list[WaterfallStep] = Field(min_length=2, max_length=8)

    def validate_split_density(self, pane_name: str, template_id: str) -> None:
        """Enforce step count limits for split layout."""
        if len(self.steps) > 5:
            raise ValueError(
                f"{pane_name} waterfall chart '{template_id}' in split layout exceeds legibility limit: "
                f"{len(self.steps)} steps > 5 steps maximum."
            )


_SAMPLE_CORE80: dict[str, Any] = {
    "unit": "億円",
    "steps": [
        {"label": "前期売上高", "value": 500.0, "is_total": True},
        {"label": "国内既存事業伸長", "value": 120.0, "is_total": False},
        {"label": "新規クラウド事業", "value": 85.0, "is_total": False},
        {"label": "原材料調達コスト高騰", "value": -45.0, "is_total": False},
        {"label": "為替変動影響", "value": -25.0, "is_total": False},
        {"label": "構造改革人件費見直し", "value": -30.0, "is_total": False},
        {"label": "業務効率化推進効果", "value": 40.0, "is_total": False},
        {"label": "当期売上高着地予測", "value": 645.0, "is_total": True},
    ],
}

_SAMPLE_CORE80_SPLIT: dict[str, Any] = {
    "unit": "億円",
    "steps": [
        {"label": "前期売上", "value": 500.0, "is_total": True},
        {"label": "既存成長", "value": 120.0, "is_total": False},
        {"label": "新規参入", "value": 85.0, "is_total": False},
        {"label": "為替影響", "value": -25.0, "is_total": False},
        {"label": "当期実績", "value": 680.0, "is_total": True},
    ],
}

_INVALID_SAMPLES_CORE80: list[dict[str, Any]] = [
    {"steps": [{"label": "S", "value": 10.0}]},  # Violates min_length=2
]

SPEC = TemplateSpec(
    payload_type=Core80Payload,
    supports_split=True,
    description="Waterfall Chart",
    sample_payload=_SAMPLE_CORE80,
    split_sample_payload=_SAMPLE_CORE80_SPLIT,
    invalid_sample_payloads=_INVALID_SAMPLES_CORE80,
)

# Local Geometry Constants (No Magic Numbers)
GAP_X = Inches(0.180)
VAL_H = Inches(0.350)
LABEL_H = Inches(0.450)
MIN_BAR_H = Inches(0.200)
MIN_BAR_W = Inches(0.350)
CONNECTOR_THICKNESS = Inches(0.020)


def _calculate_spans(
    steps: list[WaterfallStep],
) -> tuple[list[tuple[float, float, WaterfallStep]], float, float]:
    """Calculate start/end values and min/max boundaries for scaling.

    Returns:
        tuple: (list of (start_val, end_val, step_obj), max_value, min_value)
    """
    current = 0.0
    spans: list[tuple[float, float, WaterfallStep]] = []
    max_v = 0.0
    min_v = 0.0

    for step in steps:
        if step.is_total:
            s_val = 0.0
            e_val = step.value
            current = step.value
        else:
            s_val = current
            e_val = current + step.value
            current = e_val

        spans.append((s_val, e_val, step))
        max_v = max(max_v, s_val, e_val)
        min_v = min(min_v, s_val, e_val)

    return spans, max_v, min_v


def _format_step_value(step: WaterfallStep, unit: str | None) -> str:
    """Format step value as signed string with unit without precision loss."""
    formatted_num = f"{int(step.value)}" if step.value.is_integer() else f"{step.value:g}"
    val_str = f"{formatted_num}{unit}" if unit else formatted_num
    if not step.is_total and step.value > 0:
        return f"+{val_str}"
    return val_str


def _render_step(
    slide: Slide,
    step: WaterfallStep,
    unit: str | None,
    bar_box: BoundingBox,
    is_positive: bool,
    lbl_top: Length,
    box_top: Length,
) -> None:
    """Render physical bar and value/category labels for a single waterfall step."""
    if step.is_total:
        bg_col = SlideMaster.COLORS["primary"]
    elif step.value >= 0:
        bg_col = SlideMaster.COLORS["secondary"]
    else:
        bg_col = SlideMaster.COLORS["accent"]

    ShapeBuilder.add_card(slide, bar_box, bg_color=bg_col, border_width_pt=0.0)

    # Positive changes/totals place value above bar; negative changes place below bar
    val_y = Length(max(int(box_top), int(bar_box.top - VAL_H))) if is_positive else Length(min(int(lbl_top - VAL_H), int(bar_box.bottom)))
    val_box = BoundingBox(bar_box.left, val_y, bar_box.width, VAL_H)
    ShapeBuilder.add_styled_textbox(
        slide,
        val_box,
        _format_step_value(step, unit),
        max_pt=11.5,
        min_pt=8.0,
        color_key="primary",
        bold=True,
        alignment=PP_ALIGN.CENTER,
        font_family=SlideMaster.FONT_FAMILY_SECONDARY,
    )

    lbl_box = BoundingBox(bar_box.left, lbl_top, bar_box.width, LABEL_H)
    ShapeBuilder.add_styled_textbox(
        slide,
        lbl_box,
        step.label,
        max_pt=10.5,
        min_pt=8.0,
        color_key="text_dark",
        bold=True,
        alignment=PP_ALIGN.CENTER,
    )


def _render_connector(slide: Slide, curr_x: Length, bar_w: Length, gap_x: Length, conn_y: Length) -> None:
    """Render horizontal connector line to the next waterfall bar."""
    conn_box = BoundingBox(
        left=Length(curr_x + bar_w),
        top=conn_y,
        width=gap_x,
        height=CONNECTOR_THICKNESS,
    )
    ShapeBuilder.add_card(slide, conn_box, bg_color=SlideMaster.COLORS["border"], border_width_pt=0.0)


def render_core_80(slide: Slide, payload: Core80Payload, box: BoundingBox) -> None:
    """Render dynamic waterfall chart using relative shape coordinates.

    Args:
        slide: Target slide.
        payload: Core80Payload instance.
        box: Content BoundingBox.
    """
    spans, max_v, min_v = _calculate_spans(payload.steps)

    lbl_top = Length(box.bottom - LABEL_H)
    gap_y = Inches(0.080)
    chart_area_top = Length(box.top + VAL_H)
    chart_area_bottom = Length(lbl_top - VAL_H - gap_y)
    usable_h = Length(max(int(Inches(1.000)), int(chart_area_bottom - chart_area_top)))

    val_range = max_v - min_v
    scale = float(usable_h / val_range) if val_range > 0 else 1.0

    num_steps = len(payload.steps)
    gap_x = Length(min(int(GAP_X), int(box.width * 0.03)))
    total_gap = Length(int(gap_x * (num_steps - 1)))
    avail_bar_space = max(int(Inches(0.500)), int(box.width - total_gap))
    bar_w = Length(int(avail_bar_space / num_steps))

    curr_x = box.left
    for idx, (s_val, e_val, step) in enumerate(spans):
        y_start = chart_area_top + int(round((max_v - s_val) * scale))
        y_end = chart_area_top + int(round((max_v - e_val) * scale))
        bar_top = Length(min(y_start, y_end))
        bar_h = Length(max(int(MIN_BAR_H), abs(y_start - y_end)))

        if bar_top + bar_h > chart_area_bottom:
            bar_top = Length(max(int(chart_area_top), int(chart_area_bottom - bar_h)))

        bar_box = BoundingBox(curr_x, bar_top, bar_w, bar_h)
        is_pos = (e_val >= s_val) if not step.is_total else (step.value >= 0)
        _render_step(slide, step, payload.unit, bar_box, is_pos, lbl_top, box.top)

        if idx < num_steps - 1:
            conn_y = bar_top if e_val >= s_val else Length(bar_top + bar_h)
            _render_connector(slide, curr_x, bar_w, gap_x, conn_y)

        curr_x = Length(curr_x + bar_w + gap_x)
