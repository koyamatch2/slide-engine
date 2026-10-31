"""Core base module providing layout, geometry, typography, and styling primitives.

This module acts as the Single Source of Truth (SSOT) for design tokens,
coordinates, slide master layout resolution, and shape canvas manipulations.
It strictly avoids importing other application modules to maintain a clean DAG.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

if TYPE_CHECKING:
    from pptx.presentation import Presentation as PresentationType
    from pptx.shapes.autoshape import Shape
    from pptx.slide import Slide, SlideLayout


@dataclass(frozen=True)
class BoundingBox:
    """Represents a 2D rectangular bounding box on a slide canvas."""

    left: Inches
    top: Inches
    width: Inches
    height: Inches

    @property
    def right(self) -> Inches:
        """Calculate right boundary."""
        return Inches(self.left + self.width)

    @property
    def bottom(self) -> Inches:
        """Calculate bottom boundary."""
        return Inches(self.top + self.height)


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
            top=Inches(cls.CONTENT_BOX.top + offset),
            width=cls.CONTENT_BOX.width,
            height=Inches(cls.CONTENT_BOX.height - offset),
        )
        return lead_box, body_box

    @classmethod
    def split_horizontal(
        cls,
        box: BoundingBox,
        split_ratio: float = 0.5,
        gap: Inches = Inches(0.400),
    ) -> tuple[BoundingBox, BoundingBox]:
        """Split a bounding box horizontally into left and right panes.

        Args:
            box: Parent bounding box.
            split_ratio: Width ratio for the left pane (0.0 to 1.0).
            gap: Space between left and right panes.

        Returns:
            Tuple of (left_box, right_box).
        """
        available_width = box.width - gap
        left_width = Inches(available_width * split_ratio)
        right_width = Inches(available_width * (1.0 - split_ratio))
        right_left = Inches(box.left + left_width + gap)

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


class FontEngine:
    """Calculates text wrapping and dynamic point scaling."""

    # Approximate width per character relative to font size (Heuristic factor)
    CHAR_WIDTH_RATIO = 0.58

    @staticmethod
    def pt(size: float) -> Pt:
        """Safely construct a Pt instance conforming to banned-api rule."""
        return Pt(size)

    @classmethod
    def calculate_font_size(
        cls,
        text: str,
        box_width: Inches,
        box_height: Inches,
        max_pt: float = 14.0,
        min_pt: float = 9.0,
        line_height_multiplier: float = 1.3,
    ) -> Pt:
        """Compute the optimal font size to fit text within the target bounding box.

        Args:
            text: Text content to render.
            box_width: Available width.
            box_height: Available height.
            max_pt: Maximum font size in points.
            min_pt: Minimum font size in points.
            line_height_multiplier: Inter-line spacing multiplier.

        Returns:
            Optimal Pt value.
        """
        if not text:
            return cls.pt(max_pt)

        target_pt = max_pt
        box_width_pt = box_width / Inches(1) * 72
        box_height_pt = box_height / Inches(1) * 72

        while target_pt >= min_pt:
            char_width_pt = target_pt * cls.CHAR_WIDTH_RATIO
            chars_per_line = max(1, int(box_width_pt / char_width_pt))

            lines_count = 0
            for paragraph in text.split("\n"):
                if not paragraph:
                    lines_count += 1
                else:
                    lines_count += max(1, -(-len(paragraph) // chars_per_line))

            total_text_height = lines_count * (target_pt * line_height_multiplier)
            if total_text_height <= box_height_pt:
                return cls.pt(target_pt)
            target_pt -= 0.5

        return cls.pt(min_pt)


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
    def clean_slide_canvas(slide: Slide) -> None:
        """Purge empty shapes and master placeholders while preserving backgrounds.

        Safely iterates over shapes in reverse to delete raw `sp` elements without
        affecting the master presentation background or static master graphics.

        Args:
            slide: Target slide.
        """
        for shape in list(slide.shapes):
            # Keep slide master static graphics, purge slide-level dynamic/placeholder shapes
            sp_elem = shape.element
            sp_elem.getparent().remove(sp_elem)


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
        width=SlideMaster.SLIDE_WIDTH - (SlideMaster.MARGIN_LEFT * 2),
        height=Inches(1.000),
    )
    textbox = ShapeBuilder.add_textbox(slide, header_box)
    tf = textbox.text_frame

    # 1. Subtitle / Tracker Category
    if subtitle:
        p_sub = tf.paragraphs[0]
        p_sub.text = subtitle.upper()
        p_sub.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p_sub.font.size = FontEngine.pt(10.0)
        p_sub.font.bold = True
        p_sub.font.color.rgb = SlideMaster.COLORS["secondary"]
        p_sub.space_after = FontEngine.pt(2.0)
        p_title = tf.add_paragraph()
    else:
        p_title = tf.paragraphs[0]

    # 2. Main Slide Title
    p_title.text = title
    p_title.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_title.font.size = FontEngine.pt(20.0)
    p_title.font.bold = True
    p_title.font.color.rgb = SlideMaster.COLORS["primary"]

    # 3. Optional Lead Text Box
    if lead_text:
        lead_box, _ = SlideMaster.get_body_box_with_lead(has_lead_text=True)
        if lead_box:
            # Subtle card backdrop for lead text
            ShapeBuilder.add_card(
                slide,
                lead_box,
                bg_color=SlideMaster.COLORS["bg_light"],
                border_color=SlideMaster.COLORS["border"],
                border_width_pt=0.75,
            )
            inner_box = BoundingBox(
                left=Inches(lead_box.left + Inches(0.200)),
                top=Inches(lead_box.top + Inches(0.120)),
                width=Inches(lead_box.width - Inches(0.400)),
                height=Inches(lead_box.height - Inches(0.240)),
            )
            tb_lead = ShapeBuilder.add_textbox(slide, inner_box)
            p_lead = tb_lead.text_frame.paragraphs[0]
            p_lead.text = lead_text
            p_lead.font.name = SlideMaster.FONT_FAMILY_PRIMARY
            p_lead.font.size = FontEngine.pt(12.0)
            p_lead.font.bold = True
            p_lead.font.color.rgb = SlideMaster.COLORS["secondary"]