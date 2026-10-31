"""Renderer for Native Chart slide (CORE-30).

Renders native OpenXML charts (column, bar, line, pie) directly onto the slide
canvas, ensuring embedded Excel sheet data integrity and client-side editability.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

from slide_engine.core_base import BoundingBox, SlideMaster
from slide_engine.models import Core30Payload

if TYPE_CHECKING:
    from pptx.slide import Slide


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
    chart_data = CategoryChartData()
    chart_data.categories = payload.categories

    for series in payload.series:
        chart_data.add_series(series.name, series.values)

    xl_chart_type = CHART_TYPE_MAP.get(payload.chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

    chart_shape = slide.shapes.add_chart(
        xl_chart_type,
        box.left,
        box.top,
        box.width,
        box.height,
        chart_data,  # type: ignore[arg-type]
    )
    chart = getattr(chart_shape, "chart", chart_shape)

    # Configure corporate chart aesthetics
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.TOP
    chart.legend.include_in_layout = False
    chart.legend.font.name = SlideMaster.FONT_FAMILY_PRIMARY

    # Apply corporate color scheme across series
    color_keys = ["primary", "secondary", "accent", "border", "fallback_text"]
    for idx, series_obj in enumerate(chart.series):
        color_key = color_keys[idx % len(color_keys)]
        rgb_color = SlideMaster.COLORS[color_key]
        series_obj.format.fill.solid()
        series_obj.format.fill.fore_color.rgb = rgb_color
