"""Renderer for Native Chart slide (CORE-30).

Renders native OpenXML charts (column, bar, line, pie) directly onto the slide
canvas, ensuring embedded Excel sheet data integrity and client-side editability.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.enum.chart import XL_CHART_TYPE

from slide_engine.core_base import BoundingBox, ShapeBuilder, SlideMaster
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
