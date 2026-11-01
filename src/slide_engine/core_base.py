"""Core base module providing layout, geometry, typography, and styling primitives.

This module acts as the Single Source of Truth (SSOT) for design tokens,
coordinates, slide master layout resolution, and shape canvas manipulations.
It maintains zero external application dependencies to preserve a strict DAG.
"""

from __future__ import annotations

import math
import unicodedata
from collections.abc import Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Length, Pt

ColorType = RGBColor

if TYPE_CHECKING:
    from collections.abc import Sequence

    from pptx.presentation import Presentation as PresentationType
    from pptx.shapes.autoshape import Shape
    from pptx.slide import Slide, SlideLayout
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

    def inset(
        self,
        pad_x: Length,
        pad_y: Length | None = None,
        *,
        max_ratio_x: float = 1.0,
        max_ratio_y: float = 1.0,
        min_w: Length | None = None,
        min_h: Length | None = None,
    ) -> BoundingBox:
        """Create a new inner bounding box reduced by padding with safety limits."""
        raw_pad_y = pad_x if pad_y is None else pad_y
        eff_pad_x = Length(min(int(pad_x), int(self.width * max_ratio_x)))
        eff_pad_y = Length(min(int(raw_pad_y), int(self.height * max_ratio_y)))

        default_min_w = Inches(0.400) if min_w is None else min_w
        default_min_h = Inches(0.300) if min_h is None else min_h

        inner_w = Length(max(int(default_min_w), int(self.width - (eff_pad_x * 2))))
        inner_h = Length(max(int(default_min_h), int(self.height - (eff_pad_y * 2))))

        return BoundingBox(
            left=Length(self.left + eff_pad_x),
            top=Length(self.top + eff_pad_y),
            width=inner_w,
            height=inner_h,
        )

    def scale_height(self, ratio: float) -> Length:
        """Scale box height by a floating-point multiplier."""
        return Length(int(self.height * ratio))

    def scale_width(self, ratio: float) -> Length:
        """Scale box width by a floating-point multiplier."""
        return Length(int(self.width * ratio))

    def sub_grid(
        self,
        rows: int,
        cols: int,
        gap_x: Length,
        gap_y: Length | None = None,
        *,
        max_gap_x_ratio: float = 1.0,
        max_gap_y_ratio: float = 1.0,
    ) -> list[BoundingBox]:
        """Compute sub-bounding boxes for a regular grid in row-major order."""
        eff_gap_x = Length(min(int(gap_x), int(self.width * max_gap_x_ratio)))
        raw_gap_y = gap_x if gap_y is None else gap_y
        eff_gap_y = Length(min(int(raw_gap_y), int(self.height * max_gap_y_ratio)))

        cell_w = Length(int((self.width - eff_gap_x * (cols - 1)) / cols))
        cell_h = Length(int((self.height - eff_gap_y * (rows - 1)) / rows))

        grid: list[BoundingBox] = []
        for r_idx in range(rows):
            for c_idx in range(cols):
                c_left = Length(self.left + c_idx * (cell_w + eff_gap_x))
                c_top = Length(self.top + r_idx * (cell_h + eff_gap_y))
                grid.append(BoundingBox(c_left, c_top, cell_w, cell_h))
        return grid

    def adaptive_grid(
        self,
        num_items: int,
        min_col_width: Length,
        gap_x: Length,
        gap_y: Length | None = None,
        *,
        balance_rows: bool = True,
        max_gap_x_ratio: float = 0.03,
        max_gap_y_ratio: float = 0.04,
    ) -> list[BoundingBox]:
        """Compute adaptive sub-grid bounding boxes that automatically adapt columns and rows.

        Args:
            num_items: Total number of items to arrange.
            min_col_width: Minimum allowable column width before reducing columns.
            gap_x: Horizontal gap between grid cells.
            gap_y: Optional vertical gap between cells (defaults to gap_x).
            balance_rows: If True, balances column count across rows (e.g. 4 items in max 3 cols -> 2x2).
            max_gap_x_ratio: Maximum allowable horizontal gap ratio relative to width.
            max_gap_y_ratio: Maximum allowable vertical gap ratio relative to height.

        Returns:
            List of BoundingBox objects for each cell up to num_items.
        """
        if num_items <= 0:
            return []

        max_cols = max(1, min(num_items, int(self.width / min_col_width)))
        rows = math.ceil(num_items / max_cols)
        cols = math.ceil(num_items / rows) if balance_rows else max_cols

        eff_gap_y = gap_x if gap_y is None else gap_y
        all_cells = self.sub_grid(
            rows=rows,
            cols=cols,
            gap_x=gap_x,
            gap_y=eff_gap_y,
            max_gap_x_ratio=max_gap_x_ratio,
            max_gap_y_ratio=max_gap_y_ratio,
        )
        return all_cells[:num_items]


def clamp_length(
    value: int | float | Length,
    min_val: int | float | Length | None = None,
    max_val: int | float | Length | None = None,
) -> Length:
    """Clamp an EMU or integer value within optional minimum and maximum bounds."""
    val_int = int(value)
    if min_val is not None:
        val_int = max(val_int, int(min_val))
    if max_val is not None:
        val_int = min(val_int, int(max_val))
    return Length(val_int)


class SlideMaster:
    """Design system constants, slide dimensions, and master layout resolver."""

    # 16:9 Slide Dimensions
    SLIDE_WIDTH = Inches(13.333)
    SLIDE_HEIGHT = Inches(7.500)

    # Margins and content bounds
    MARGIN_LEFT = Inches(0.800)
    MARGIN_RIGHT = Inches(0.800)
    MARGIN_TOP = Inches(0.600)
    MARGIN_BOTTOM = Inches(0.600)

    # Standard body content box (without lead text)
    CONTENT_BOX = BoundingBox(
        left=Inches(0.800),
        top=Inches(1.800),
        width=Inches(11.733),
        height=Inches(5.100),
    )

    # Lead text offset parameters
    LEAD_TEXT_HEIGHT = Inches(0.600)
    LEAD_TEXT_GAP = Inches(0.250)

    # Color Palette (Enterprise Navy & Neutral Standard)
    COLORS: dict[str, RGBColor] = {
        "primary": RGBColor(16, 44, 87),       # Deep Navy
        "secondary": RGBColor(53, 95, 142),    # Steel Blue
        "accent": RGBColor(218, 98, 43),       # Burnt Orange
        "text_dark": RGBColor(33, 37, 41),     # Off-Black
        "text_light": RGBColor(248, 249, 250), # Off-White
        "text_muted": RGBColor(108, 117, 125), # Gray
        "bg_light": RGBColor(245, 247, 250),   # Light Cool Gray
        "bg_white": RGBColor(255, 255, 255),   # Pure White
        "border": RGBColor(222, 226, 230),     # Border Gray
        "fallback_bg": RGBColor(255, 243, 205),# Soft Amber
        "fallback_text": RGBColor(133, 100, 4),# Amber Dark
    }

    # Typography
    FONT_FAMILY_PRIMARY = "Meiryo UI"
    FONT_FAMILY_SECONDARY = "Arial"

    @classmethod
    def get_body_box_with_lead(cls, has_lead_text: bool) -> tuple[BoundingBox | None, BoundingBox]:
        """Compute bounding boxes for optional lead text and remaining body content.

        Args:
            has_lead_text: Whether lead text is present.

        Returns:
            A tuple of (lead_box, remaining_body_box).
        """
        if not has_lead_text:
            return None, cls.CONTENT_BOX

        lead_box = BoundingBox(
            left=cls.CONTENT_BOX.left,
            top=cls.CONTENT_BOX.top,
            width=cls.CONTENT_BOX.width,
            height=cls.LEAD_TEXT_HEIGHT,
        )

        offset = cls.LEAD_TEXT_HEIGHT + cls.LEAD_TEXT_GAP
        body_box = BoundingBox(
            left=cls.CONTENT_BOX.left,
            top=Length(cls.CONTENT_BOX.top + offset),
            width=cls.CONTENT_BOX.width,
            height=Length(cls.CONTENT_BOX.height - offset),
        )
        return lead_box, body_box

    DEFAULT_SPLIT_GAP = Inches(0.400)

    @classmethod
    def split_horizontal(
        cls,
        box: BoundingBox,
        split_ratio: float = 0.5,
        gap: Length | None = None,
    ) -> tuple[BoundingBox, BoundingBox]:
        """Split a bounding box horizontally into left and right panes.

        Args:
            box: Parent bounding box.
            split_ratio: Width ratio for the left pane (0.0 to 1.0).
            gap: Space between left and right panes.

        Returns:
            Tuple of (left_box, right_box).
        """
        effective_gap = cls.DEFAULT_SPLIT_GAP if gap is None else gap
        available_width = box.width - effective_gap
        left_width = Length(int(available_width * split_ratio))
        right_width = Length(int(available_width * (1.0 - split_ratio)))
        right_left = Length(box.left + left_width + effective_gap)

        left_box = BoundingBox(box.left, box.top, left_width, box.height)
        right_box = BoundingBox(right_left, box.top, right_width, box.height)
        return left_box, right_box

    @staticmethod
    def resolve_slide_layout(prs: PresentationType, slide_type: str) -> SlideLayout:
        """Resolve the appropriate master slide layout by name or fallback index.

        Args:
            prs: Presentation instance.
            slide_type: 'cover', 'agenda', 'body', or 'blank'.

        Returns:
            The resolved SlideLayout object.
        """
        name_targets: dict[str, list[str]] = {
            "cover": ["title", "cover", "タイトル", "表紙"],
            "agenda": ["section", "header", "中扉", "セクション"],
            "body": ["blank", "title and content", "白紙", "コンテンツ"],
            "blank": ["blank", "白紙"],
        }
        targets = name_targets.get(slide_type, ["blank", "白紙"])

        # Name-based match (case-insensitive)
        for layout in prs.slide_layouts:
            layout_name = layout.name.lower()
            if any(target in layout_name for target in targets):
                return layout

        # Fallback index logic: 0 for cover, 6 for blank, otherwise default first available
        if slide_type == "cover" and len(prs.slide_layouts) > 0:
            return prs.slide_layouts[0]
        if len(prs.slide_layouts) > 6:
            return prs.slide_layouts[6]
        return prs.slide_layouts[0]


@dataclass(frozen=True)
class ParagraphSpec:
    """Specification for an individual paragraph inside an adaptive text stack."""

    text: str | None
    weight: float = 1.0
    max_pt: float = 12.0
    min_pt: float = 8.0
    color_key: str = "text_dark"
    bold: bool = False
    alignment: PP_ALIGN | None = None
    font_family: str = SlideMaster.FONT_FAMILY_PRIMARY
    space_after_pt: float | None = None


class FontEngine:
    """Calculates text wrapping and dynamic point scaling."""

    # Approximate width per character relative to font size (Heuristic factor)
    CHAR_WIDTH_RATIO = 0.58
    FULLWIDTH_CHAR_RATIO = 1.0

    @staticmethod
    def pt(size: float) -> Length:
        """Safely construct a Pt instance conforming to banned-api rule."""
        return Length(Pt(size))

    @classmethod
    def _effective_text_em_width(cls, text: str) -> float:
        """Compute effective width in em units using Unicode East Asian Width properties."""
        total_em = 0.0
        for char in text:
            if unicodedata.east_asian_width(char) in ("F", "W", "A"):
                total_em += cls.FULLWIDTH_CHAR_RATIO
            else:
                total_em += cls.CHAR_WIDTH_RATIO
        return total_em

    @classmethod
    def calculate_font_size(
        cls,
        text: str,
        box_width: Length,
        box_height: Length,
        max_pt: float = 14.0,
        min_pt: float = 9.0,
        line_height_multiplier: float = 1.3,
        space_after_pt: float = 0.0,
    ) -> Length:
        """Compute the optimal font size to fit text within the target bounding box.

        Args:
            text: Text content to render.
            box_width: Available width.
            box_height: Available height.
            max_pt: Maximum font size in points.
            min_pt: Minimum font size in points.
            line_height_multiplier: Inter-line spacing multiplier.
            space_after_pt: Optional extra space after paragraph in points.

        Returns:
            Optimal Pt value.
        """
        if not text:
            return cls.pt(max_pt)

        # Precompute paragraph em widths once to avoid expensive repetitive scans in search loop
        paragraphs = text.split("\n")
        para_specs = [(cls._effective_text_em_width(p) if p else 0.0, bool(p)) for p in paragraphs]

        target_pt = max_pt
        box_width_pt = box_width.pt
        box_height_pt = box_height.pt

        while target_pt >= min_pt:
            avail_em = max(1.0, box_width_pt / target_pt)
            lines_count = cls._estimate_line_count(para_specs, avail_em)
            total_text_height = (lines_count * target_pt * line_height_multiplier) + space_after_pt
            if total_text_height <= box_height_pt:
                return cls.pt(target_pt)
            target_pt -= 0.5

        return cls.pt(min_pt)

    @staticmethod
    def _estimate_line_count(para_specs: list[tuple[float, bool]], avail_em: float) -> int:
        """Estimate lines needed across all paragraphs with zero deep nesting."""
        lines = 0
        for eff_em, is_non_empty in para_specs:
            if not is_non_empty:
                lines += 1
            else:
                lines += max(1, math.ceil(eff_em / avail_em))
        return lines

    @classmethod
    def estimate_text_lines(cls, text: str, width: Length, font_pt: float) -> int:
        """Estimate lines required to wrap text within available width at given font size."""
        if not text:
            return 0
        paragraphs = text.split("\n")
        para_specs = [(cls._effective_text_em_width(p) if p else 0.0, bool(p)) for p in paragraphs]
        avail_em = max(1.0, width.pt / max(1.0, font_pt))
        return cls._estimate_line_count(para_specs, avail_em)


def _compute_spec_min_lines(spec: ParagraphSpec, width: Length) -> int:
    """Calculate minimum line count for a paragraph spec using its min_pt."""
    if not spec.text or not spec.text.strip():
        return 0
    return max(1, FontEngine.estimate_text_lines(spec.text, width, spec.min_pt))


def _resolve_stack_spaces(
    specs: Sequence[ParagraphSpec],
    avail_h_pt: float,
    min_heights: list[float],
) -> list[float]:
    """Dynamically compress paragraph space_after values when vertical space is constrained."""
    num_specs = len(specs)
    raw_spaces = [
        (spec.space_after_pt or 0.0) if idx < num_specs - 1 else 0.0
        for idx, spec in enumerate(specs)
    ]
    total_req_space = sum(raw_spaces)
    if total_req_space <= 0.0:
        return raw_spaces

    total_min_text = sum(min_heights)
    available_for_space = max(0.0, avail_h_pt - total_min_text)
    factor = min(1.0, available_for_space / total_req_space)
    return [s * factor for s in raw_spaces]


def _allocate_stack_budgets(
    specs: Sequence[ParagraphSpec],
    width: Length,
    height: Length,
) -> list[tuple[Length, float]]:
    """Allocate optimal height budgets and space_after_pt for each paragraph in a stack.

    Returns:
        List of tuples: (budget_h, effective_space_after_pt).
    """
    total_weight = sum(s.weight for s in specs)
    if total_weight <= 0:
        total_weight = float(len(specs))

    avail_h_pt = height.pt
    min_heights = [
        _compute_spec_min_lines(spec, width) * spec.min_pt * 1.3
        for spec in specs
    ]
    eff_spaces = _resolve_stack_spaces(specs, avail_h_pt, min_heights)

    base_heights = [h + s for h, s in zip(min_heights, eff_spaces, strict=True)]
    total_base = sum(base_heights)
    slack_pt = max(0.0, avail_h_pt - total_base)

    results: list[tuple[Length, float]] = []
    for idx, spec in enumerate(specs):
        ratio = spec.weight / total_weight
        budget_pt = base_heights[idx] + (slack_pt * ratio)
        results.append((FontEngine.pt(budget_pt), eff_spaces[idx]))

    return results


class ShapeBuilder:
    """Helper methods for constructing PowerPoint shapes, borders, and text boxes."""

    @staticmethod
    def add_card(
        slide: Slide,
        box: BoundingBox,
        bg_color: RGBColor | None = None,
        border_color: RGBColor | None = None,
        border_width_pt: float = 1.0,
    ) -> Shape:
        """Add a rounded or standard rectangular card shape.

        Args:
            slide: Target slide.
            box: BoundingBox coordinates.
            bg_color: Fill color.
            border_color: Line border color.
            border_width_pt: Border width in points.

        Returns:
            The created Shape instance.
        """
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
            shape.line.width = FontEngine.pt(border_width_pt)
        else:
            shape.line.fill.background()

        return shape

    @staticmethod
    def add_textbox(
        slide: Slide,
        box: BoundingBox,
        margin_zero: bool = True,
    ) -> Shape:
        """Add an empty text box with optional zero-margin reset.

        Args:
            slide: Target slide.
            box: BoundingBox coordinates.
            margin_zero: Whether to set text frame internal margins to 0.

        Returns:
            The created Shape instance.
        """
        textbox = slide.shapes.add_textbox(box.left, box.top, box.width, box.height)
        tf = textbox.text_frame
        tf.word_wrap = True
        if margin_zero:
            tf.margin_left = Inches(0)
            tf.margin_right = Inches(0)
            tf.margin_top = Inches(0)
            tf.margin_bottom = Inches(0)
        return textbox

    @staticmethod
    def set_cell_solid_fill(cell: _Cell, color: RGBColor) -> None:
        """Apply solid fill and RGB color to a table cell."""
        cell.fill.solid()
        cell.fill.fore_color.rgb = color
        cell.margin_left = Length(Pt(3.0))
        cell.margin_right = Length(Pt(3.0))
        cell.margin_top = Length(Pt(1.5))
        cell.margin_bottom = Length(Pt(1.5))

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
        """Add and style a chart within the target bounding box."""
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
        """Purge empty shapes and master placeholders while preserving backgrounds.

        Safely iterates over shapes in reverse to delete raw `sp` elements without
        affecting the master presentation background or static master graphics.

        Args:
            slide: Target slide.
        """
        for shape in list(slide.shapes):
            sp_elem = shape.element
            sp_elem.getparent().remove(sp_elem)

    @staticmethod
    def purge_existing_slides(prs: PresentationType) -> None:
        """Purge all pre-existing dummy slides from the master presentation."""
        for i in range(len(prs.slides) - 1, -1, -1):
            r_id = prs.slides._sldIdLst[i].rId
            prs.part.drop_rel(r_id)
            del prs.slides._sldIdLst[i]

    @classmethod
    def add_styled_paragraph(
        cls,
        tf: Any,
        text: str,
        width: Length,
        height: Length,
        *,
        max_pt: float,
        min_pt: float = 8.0,
        color_key: str = "text_dark",
        bold: bool = False,
        alignment: PP_ALIGN | None = None,
        font_family: str = SlideMaster.FONT_FAMILY_PRIMARY,
        space_after_pt: float | None = None,
        is_first: bool = False,
    ) -> Any:
        """Add and format a styled paragraph with automated font scaling.

        Args:
            tf: Target TextFrame.
            text: Text content.
            width: Available width for font fitting.
            height: Available height budget for font fitting.
            max_pt: Maximum font size in points.
            min_pt: Minimum font size in points.
            color_key: Color name in SlideMaster.COLORS.
            bold: Bold styling.
            alignment: Optional text alignment.
            font_family: Font family name.
            space_after_pt: Optional space after paragraph in points.
            is_first: Whether to reuse existing first paragraph if empty.

        Returns:
            The created or reused Paragraph.
        """
        p = (
            tf.paragraphs[0]
            if (is_first and len(tf.paragraphs) > 0 and not tf.paragraphs[0].text)
            else tf.add_paragraph()
        )
        p.text = text
        p.font.name = font_family
        p.font.size = FontEngine.calculate_font_size(
            text,
            width,
            height,
            max_pt=max_pt,
            min_pt=min_pt,
            space_after_pt=space_after_pt or 0.0,
        )
        p.font.bold = bold
        p.font.color.rgb = SlideMaster.COLORS[color_key]
        if alignment is not None:
            p.alignment = alignment
        if space_after_pt is not None:
            p.space_after = FontEngine.pt(space_after_pt)
        return p

    @classmethod
    def add_styled_textbox(
        cls,
        slide: Slide,
        box: BoundingBox,
        text: str,
        *,
        max_pt: float,
        min_pt: float = 8.0,
        color_key: str = "text_dark",
        bold: bool = False,
        alignment: PP_ALIGN | None = None,
        font_family: str = SlideMaster.FONT_FAMILY_PRIMARY,
        margin_zero: bool = True,
        space_after_pt: float | None = None,
    ) -> Shape:
        """Add a complete textbox with styled single-paragraph text."""
        tb = cls.add_textbox(slide, box, margin_zero=margin_zero)
        cls.add_styled_paragraph(
            tb.text_frame,
            text,
            box.width,
            box.height,
            max_pt=max_pt,
            min_pt=min_pt,
            color_key=color_key,
            bold=bold,
            alignment=alignment,
            font_family=font_family,
            space_after_pt=space_after_pt,
            is_first=True,
        )
        return tb

    @classmethod
    def populate_paragraph_stack(
        cls,
        tf: Any,
        width: Length,
        height: Length,
        specs: Sequence[ParagraphSpec],
    ) -> None:
        """Render a vertical stack of styled paragraphs into a TextFrame with automated height budgeting."""
        valid_specs = [s for s in specs if s.text and s.text.strip()]
        if not valid_specs:
            return

        budgets_and_spaces = _allocate_stack_budgets(valid_specs, width, height)

        for idx, (spec, (budget_h, eff_space_pt)) in enumerate(
            zip(valid_specs, budgets_and_spaces, strict=True)
        ):
            cls.add_styled_paragraph(
                tf,
                spec.text,  # type: ignore[arg-type]
                width,
                budget_h,
                max_pt=spec.max_pt,
                min_pt=spec.min_pt,
                color_key=spec.color_key,
                bold=spec.bold,
                alignment=spec.alignment,
                font_family=spec.font_family,
                space_after_pt=eff_space_pt if eff_space_pt > 0 else None,
                is_first=(idx == 0),
            )

    @classmethod
    def add_styled_paragraph_stack(
        cls,
        slide: Slide,
        box: BoundingBox,
        specs: Sequence[ParagraphSpec],
        *,
        margin_zero: bool = True,
    ) -> Shape:
        """Create a textbox and render a vertical stack of styled paragraphs with automated height budgeting."""
        tb = cls.add_textbox(slide, box, margin_zero=margin_zero)
        cls.populate_paragraph_stack(tb.text_frame, box.width, box.height, specs)
        return tb


def render_standard_header(
    slide: Slide,
    title: str,
    subtitle: str | None = None,
    lead_text: str | None = None,
) -> None:
    """Render standardized corporate slide header with optional lead text.

    Args:
        slide: Target slide.
        title: Main slide title.
        subtitle: Optional tracker/category subtitle.
        lead_text: Optional one-sentence executive summary lead text.
    """
    # Header zone: Top 0.600 to 1.600
    header_box = BoundingBox(
        left=SlideMaster.MARGIN_LEFT,
        top=SlideMaster.MARGIN_TOP,
        width=Length(SlideMaster.SLIDE_WIDTH - (SlideMaster.MARGIN_LEFT * 2)),
        height=Inches(1.000),
    )
    ShapeBuilder.add_styled_paragraph_stack(
        slide,
        header_box,
        [
            ParagraphSpec(
                text=subtitle.upper() if subtitle else None,
                weight=0.28,
                max_pt=10.0,
                min_pt=10.0,
                color_key="secondary",
                bold=True,
                space_after_pt=2.0,
            ),
            ParagraphSpec(
                text=title,
                weight=0.72,
                max_pt=20.0,
                min_pt=14.0,
                color_key="primary",
                bold=True,
            ),
        ],
    )

    # 3. Optional Lead Text Box
    if lead_text:
        lead_box, _ = SlideMaster.get_body_box_with_lead(has_lead_text=True)
        if lead_box:
            ShapeBuilder.add_card(
                slide,
                lead_box,
                bg_color=SlideMaster.COLORS["bg_light"],
                border_color=SlideMaster.COLORS["border"],
                border_width_pt=0.75,
            )
            inner_box = lead_box.inset(Inches(0.200), Inches(0.120))
            ShapeBuilder.add_styled_textbox(
                slide,
                inner_box,
                lead_text,
                max_pt=12.0,
                min_pt=8.0,
                color_key="secondary",
                bold=True,
            )
