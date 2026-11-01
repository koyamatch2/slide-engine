"""Renderer for Native Chart slide (CORE-30).

Renders native OpenXML charts (column, bar, line, pie) directly onto the slide
canvas, ensuring embedded Excel sheet data integrity and client-side editability.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Literal

from pptx.enum.chart import XL_CHART_TYPE
from pydantic import Field, model_validator

from slide_engine.core_base import BoundingBox, ShapeBuilder, SlideMaster
from slide_engine.models import BasePayload, ShortText, TemplateSpec

if TYPE_CHECKING:
    from pptx.slide import Slide


class ChartSeries(BasePayload):
    """Data series for native charts."""

    name: ShortText
    values: list[float] = Field(min_length=1, max_length=20)


class Core30Payload(BasePayload):
    """Payload for Native OpenXML Chart slide (CORE-30)."""

    chart_type: Literal["column", "bar", "line", "pie"]
    categories: list[ShortText] = Field(min_length=1, max_length=20)
    series: list[ChartSeries] = Field(min_length=1, max_length=6)

    @model_validator(mode="after")
    def validate_series_length(self) -> Core30Payload:
        """Validate that all series have exactly the same length as categories."""
        expected_len = len(self.categories)
        for s in self.series:
            if len(s.values) != expected_len:
                msg = (
                    f"Series '{s.name}' length ({len(s.values)}) does not match "
                    f"categories length ({expected_len})."
                )
                raise ValueError(msg)
        return self

    def validate_split_density(self, pane_name: str, template_id: str) -> None:
        """Enforce category and series limits for split layout."""
        if len(self.categories) > 8:
            raise ValueError(
                f"{pane_name} chart '{template_id}' in split layout exceeds legibility limit: "
                f"{len(self.categories)} categories > 8 categories maximum."
            )
        if len(self.series) > 3:
            raise ValueError(
                f"{pane_name} chart '{template_id}' in split layout exceeds legibility limit: "
                f"{len(self.series)} series > 3 series maximum."
            )


_SAMPLE_CORE30: dict[str, Any] = {
    "chart_type": "column",
    "categories": [f"M{q}" for q in range(1, 21)],
    "series": [
        {"name": f"事業部 {s}", "values": [float(s * 10 + q) for q in range(1, 21)]}
        for s in range(1, 7)
    ],
}

_SAMPLE_CORE30_SPLIT: dict[str, Any] = {
    "chart_type": "column",
    "categories": [f"Q{q}" for q in range(1, 9)],
    "series": [
        {"name": f"事業部 {s}", "values": [float(s * 10 + q) for q in range(1, 9)]}
        for s in range(1, 4)
    ],
}

_INVALID_SAMPLES_CORE30: list[dict[str, Any]] = [
    {
        "chart_type": "column",
        "categories": ["Q1", "Q2", "Q3"],
        "series": [{"name": "Revenue", "values": [10.0, 20.0]}],
    }
]

SPEC = TemplateSpec(
    payload_type=Core30Payload,
    supports_split=True,
    description="Native Chart",
    sample_payload=_SAMPLE_CORE30,
    split_sample_payload=_SAMPLE_CORE30_SPLIT,
    invalid_sample_payloads=_INVALID_SAMPLES_CORE30,
)


CHART_TYPE_MAP: dict[str, XL_CHART_TYPE] = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
}


def render_core_30(slide: Slide, payload: Core30Payload, box: BoundingBox) -> None:
    """Render a native OpenXML chart within the bounding box.

    Args:
        slide: Target slide.
        payload: Core30Payload instance.
        box: Content BoundingBox.
    """
    series_items = [(s.name, s.values) for s in payload.series]
    chart_data = ShapeBuilder.build_chart_data(payload.categories, series_items)

    xl_chart_type = CHART_TYPE_MAP.get(payload.chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

    color_keys = ["primary", "secondary", "accent", "border", "fallback_text"]
    series_colors = [SlideMaster.COLORS[k] for k in color_keys]

    ShapeBuilder.add_styled_chart(
        slide=slide,
        chart_type=xl_chart_type,
        box=box,
        chart_data=chart_data,
        font_family=SlideMaster.FONT_FAMILY_PRIMARY,
        series_colors=series_colors,
    )
