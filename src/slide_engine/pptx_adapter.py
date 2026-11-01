"""Low-level adapter isolating untyped third-party python-pptx calls.

All direct calls to python-pptx constructors and methods that lack PEP 561
typed stubs (e.g., RGBColor, CategoryChartData, cell.fill.solid, shape fill/line
formatting, and chart styling) are strictly isolated in this module behind
typed function signatures.
"""

from __future__ import annotations

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


def create_rgb_color(r: int, g: int, b: int) -> RGBColor:
    """Instantiate an RGBColor safely encapsulating untyped constructor call."""
    return RGBColor(r, g, b)


def add_rounded_card_shape(
    slide: Slide,
    left: Length,
    top: Length,
    width: Length,
    height: Length,
    bg_color: RGBColor | None,
    border_color: RGBColor | None,
    border_width: Length,
) -> Shape:
    """Add a rounded rectangle shape and configure fill and border styling."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left,
        top,
        width,
        height,
    )
    if bg_color is not None:
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg_color
    else:
        shape.fill.background()

    if border_color is not None:
        shape.line.color.rgb = border_color
        shape.line.width = border_width
    else:
        shape.line.fill.background()

    return shape


def add_textbox_shape(
    slide: Slide,
    left: Length,
    top: Length,
    width: Length,
    height: Length,
    margin_zero: bool,
) -> Shape:
    """Add a textbox shape with word wrap and optional zero-margin reset."""
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    tf.word_wrap = True
    if margin_zero:
        tf.margin_left = Inches(0)
        tf.margin_right = Inches(0)
        tf.margin_top = Inches(0)
        tf.margin_bottom = Inches(0)
    return textbox


def set_cell_solid_fill(cell: _Cell, color: RGBColor) -> None:
    """Apply solid fill, RGB color, and standardized internal margins to a table cell."""
    cell.fill.solid()
    cell.fill.fore_color.rgb = color
    cell.margin_left = Length(Pt(3.0))
    cell.margin_right = Length(Pt(3.0))
    cell.margin_top = Length(Pt(1.5))
    cell.margin_bottom = Length(Pt(1.5))


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


def add_styled_chart(
    slide: Slide,
    chart_type: XL_CHART_TYPE,
    left: Length,
    top: Length,
    width: Length,
    height: Length,
    chart_data: CategoryChartData,
    font_family: str,
    series_colors: Sequence[RGBColor],
) -> None:
    """Add and style a chart on the slide canvas, encapsulating untyped chart APIs."""
    chart_shape = slide.shapes.add_chart(
        chart_type,
        left,
        top,
        width,
        height,
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


def clean_slide_shapes(slide: Slide) -> None:
    """Purge slide-level shapes while preserving slide master backgrounds."""
    for shape in list(slide.shapes):
        sp_elem = shape.element
        sp_elem.getparent().remove(sp_elem)
