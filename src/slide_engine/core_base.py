"""Core base module providing layout, geometry, typography, and styling primitives.

This module acts as the Single Source of Truth (SSOT) for design tokens,
coordinates, slide master layout resolution, and shape canvas manipulations.
It strictly avoids importing other application modules to maintain a clean DAG.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from pptx.dml.color import RGBColor
from pptx.util import Inches, Length, Pt

from slide_engine.pptx_adapter import (
    ShapeBuilder as ShapeBuilder,
)
from slide_engine.pptx_adapter import (
    create_rgb_color,
)

ColorType = RGBColor

if TYPE_CHECKING:
    from pptx.presentation import Presentation as PresentationType
    from pptx.slide import Slide, SlideLayout


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
        "primary": create_rgb_color(16, 44, 87),       # Deep Navy
        "secondary": create_rgb_color(53, 95, 142),    # Steel Blue
        "accent": create_rgb_color(218, 98, 43),       # Burnt Orange
        "text_dark": create_rgb_color(33, 37, 41),     # Off-Black
        "text_light": create_rgb_color(248, 249, 250), # Off-White
        "text_muted": create_rgb_color(108, 117, 125), # Gray
        "bg_light": create_rgb_color(245, 247, 250),   # Light Cool Gray
        "bg_white": create_rgb_color(255, 255, 255),   # Pure White
        "border": create_rgb_color(222, 226, 230),     # Border Gray
        "fallback_bg": create_rgb_color(255, 243, 205),# Soft Amber
        "fallback_text": create_rgb_color(133, 100, 4),# Amber Dark
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


class FontEngine:
    """Calculates text wrapping and dynamic point scaling."""

    # Approximate width per character relative to font size (Heuristic factor)
    CHAR_WIDTH_RATIO = 0.58

    @staticmethod
    def pt(size: float) -> Length:
        """Safely construct a Length point value conforming to banned-api rule."""
        return Length(Pt(size))

    @classmethod
    def calculate_font_size(
        cls,
        text: str,
        box_width: Length,
        box_height: Length,
        max_pt: float = 14.0,
        min_pt: float = 9.0,
        line_height_multiplier: float = 1.3,
    ) -> Length:
        """Compute the optimal font size to fit text within the target bounding box.

        Args:
            text: Text content to render.
            box_width: Available width Length.
            box_height: Available height Length.
            max_pt: Maximum font size in points.
            min_pt: Minimum font size in points.
            line_height_multiplier: Inter-line spacing multiplier.

        Returns:
            Optimal Length point value.
        """
        if not text:
            return cls.pt(max_pt)

        target_pt = max_pt
        box_width_pt = box_width.pt
        box_height_pt = box_height.pt

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
                left=Length(lead_box.left + Inches(0.200)),
                top=Length(lead_box.top + Inches(0.120)),
                width=Length(lead_box.width - Inches(0.400)),
                height=Length(lead_box.height - Inches(0.240)),
            )
            tb_lead = ShapeBuilder.add_textbox(slide, inner_box)
            p_lead = tb_lead.text_frame.paragraphs[0]
            p_lead.text = lead_text
            p_lead.font.name = SlideMaster.FONT_FAMILY_PRIMARY
            p_lead.font.size = FontEngine.pt(12.0)
            p_lead.font.bold = True
            p_lead.font.color.rgb = SlideMaster.COLORS["secondary"]
