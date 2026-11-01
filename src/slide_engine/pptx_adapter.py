"""Adapter layer isolating untyped third-party python-pptx calls.

All direct calls to python-pptx methods that lack PEP 561 typed stubs
(e.g., cell.fill.solid, CategoryChartData creation, Shape.fill/line manipulation,
and chart series formatting) are strictly isolated in this module behind typed
public signatures. This module sits at the bottom of the dependency DAG and
imports no other slide_engine modules.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Length, Pt

if TYPE_CHECKING:
    from collections.abc import Sequence

    from pptx.shapes.autoshape import Shape
    from pptx.slide import Slide
    from pptx.table import _Cell


@dataclass(frozen=True)
class BoundingBox:
    """Represents a 2D rectangular bounding box on a slide canvas."""

    left: Length
    top: Length
    width: Length
    height: Length

    @property
    def right(self) -> Length:
        """Calculate right boundary."""
        return Length(self.left + self.width)

    @property
    def bottom(self) -> Length:
        """Calculate bottom boundary."""
        return Length(self.top + self.height)


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
    def set_cell_solid_fill(
        cell: _Cell,
        color: RGBColor,
        compact_margins: bool = True,
    ) -> None:
        """Apply solid fill, RGB color, and compact internal margins to a table cell."""
        cell.fill.solid()
        cell.fill.fore_color.rgb = color
        if compact_margins:
            cell.margin_left = Length(Pt(4))
            cell.margin_right = Length(Pt(4))
            cell.margin_top = Length(Pt(2))
            cell.margin_bottom = Length(Pt(2))

    @staticmethod
    def build_chart_data(
        categories: list[str],
        series_items: Sequence[tuple[str, Sequence[float]]],
    ) -> CategoryChartData:
        """Construct CategoryChartData safely encapsulating untyped pptx calls."""
        chart_data = CategoryChartData()
        chart_data.categories = categories
        for name, values in series_items:
            chart_data.add_series(name, list(values))
        return chart_data

    @staticmethod
    def add_styled_chart(
        slide: Slide,
        chart_type: XL_CHART_TYPE,
        box: BoundingBox,
        chart_data: CategoryChartData,
        font_family: str,
        series_colors: Sequence[RGBColor],
    ) -> None:
        """Add and style a chart on the slide canvas, encapsulating all untyped chart APIs."""
        chart_shape = slide.shapes.add_chart(
            chart_type,
            box.left,
            box.top,
            box.width,
            box.height,
            chart_data,  # type: ignore[arg-type]
        )
        chart = getattr(chart_shape, "chart", chart_shape)
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.TOP
        chart.legend.include_in_layout = False
        chart.legend.font.name = font_family

        if series_colors:
            for idx, series_obj in enumerate(chart.series):
                rgb_color = series_colors[idx % len(series_colors)]
                series_obj.format.fill.solid()
                series_obj.format.fill.fore_color.rgb = rgb_color

    @staticmethod
    def clean_slide_canvas(slide: Slide) -> None:
        """Purge empty shapes and master placeholders while preserving backgrounds."""
        for shape in list(slide.shapes):
            sp_elem = shape.element
            sp_elem.getparent().remove(sp_elem)
