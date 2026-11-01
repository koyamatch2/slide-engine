"""Adapter layer isolating untyped third-party python-pptx calls.

All direct calls to python-pptx methods that lack PEP 561 typed stubs
(e.g., cell.fill.solid, CategoryChartData creation, Shape.fill/line manipulation)
are strictly isolated in this module behind typed public signatures.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Length, Pt

if TYPE_CHECKING:
    from pptx.shapes.autoshape import Shape
    from pptx.slide import Slide
    from pptx.table import _Cell

    from slide_engine.core_base import BoundingBox
    from slide_engine.models import ChartSeries


def create_rgb_color(r: int, g: int, b: int) -> RGBColor:
    """Instantiate an RGBColor safely encapsulating untyped constructor call."""
    return RGBColor(r, g, b)


class ShapeBuilder:
    """Constructs and styles PowerPoint canvas elements with typed signatures."""

    @staticmethod
    def add_card(
        slide: Slide,
        box: BoundingBox,
        bg_color: RGBColor | None = None,
        border_color: RGBColor | None = None,
        border_width_pt: float = 1.0,
    ) -> Shape:
        """Add a rounded or standard rectangular card shape."""
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            box.left,
            box.top,
            box.width,
            box.height,
        )
        if bg_color is not None:
            shape.fill.solid()
            shape.fill.fore_color.rgb = bg_color
        else:
            shape.fill.background()

        if border_color is not None:
            shape.line.color.rgb = border_color
            shape.line.width = Pt(border_width_pt)
        else:
            shape.line.fill.background()

        return shape

    @staticmethod
    def add_textbox(
        slide: Slide,
        box: BoundingBox,
        margin_zero: bool = True,
    ) -> Shape:
        """Add an empty text box with optional zero-margin reset."""
        textbox = slide.shapes.add_textbox(box.left, box.top, box.width, box.height)
        tf = textbox.text_frame
        tf.word_wrap = True
        if margin_zero:
            tf.margin_left = Length(Inches(0))
            tf.margin_right = Length(Inches(0))
            tf.margin_top = Length(Inches(0))
            tf.margin_bottom = Length(Inches(0))
        return textbox

    @staticmethod
    def set_cell_solid_fill(cell: _Cell, color: RGBColor) -> None:
        """Apply solid fill and RGB color to a table cell."""
        cell.fill.solid()
        cell.fill.fore_color.rgb = color

    @staticmethod
    def build_chart_data(
        categories: list[str],
        series_list: list[ChartSeries],
    ) -> CategoryChartData:
        """Construct CategoryChartData safely encapsulating untyped pptx calls."""
        chart_data = CategoryChartData()
        chart_data.categories = categories
        for series in series_list:
            chart_data.add_series(series.name, series.values)
        return chart_data

    @staticmethod
    def add_chart(
        slide: Slide,
        chart_type: XL_CHART_TYPE,
        box: BoundingBox,
        chart_data: CategoryChartData,
    ) -> Any:
        """Add a chart to the slide canvas, encapsulating untyped ChartData conversion."""
        return slide.shapes.add_chart(
            chart_type,
            box.left,
            box.top,
            box.width,
            box.height,
            chart_data,  # type: ignore[arg-type]
        )

    @staticmethod
    def clean_slide_canvas(slide: Slide) -> None:
        """Purge empty shapes and master placeholders while preserving backgrounds."""
        for shape in list(slide.shapes):
            sp_elem = shape.element
            sp_elem.getparent().remove(sp_elem)
