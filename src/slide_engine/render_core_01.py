"""Renderer for Cover slide (CORE-01).

Renders presentation title, subtitle, presenter, and date.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.util import Inches, Length

from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import Core01Payload

if TYPE_CHECKING:
    from pptx.slide import Slide


# Cover Layout Constants
COVER_LEFT = Inches(1.200)
COVER_TOP = Inches(2.200)
COVER_WIDTH = Inches(10.933)
COVER_HEIGHT = Inches(2.400)

META_OFFSET_TOP = Inches(4.900)
META_HEIGHT = Inches(1.200)


def render_core_01(slide: Slide, payload: Core01Payload, box: BoundingBox) -> None:
    """Render presentation cover slide elements.

    Args:
        slide: Target slide.
        payload: Core01Payload instance.
        box: Content BoundingBox (overridden by cover-specific geometry).
    """
    # 1. Main Title & Subtitle Box
    title_box = BoundingBox(
        left=COVER_LEFT,
        top=COVER_TOP,
        width=COVER_WIDTH,
        height=COVER_HEIGHT,
    )
    tb_title = ShapeBuilder.add_textbox(slide, title_box)
    tf_title = tb_title.text_frame

    title_budget_h = Length(int(COVER_HEIGHT * 0.55)) if payload.subtitle else COVER_HEIGHT

    # Main Title
    p_title = tf_title.paragraphs[0]
    p_title.text = payload.title
    p_title.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_title.font.size = FontEngine.calculate_font_size(
        payload.title,
        COVER_WIDTH,
        title_budget_h,
        max_pt=32.0,
        min_pt=18.0,
    )
    p_title.font.bold = True
    p_title.font.color.rgb = SlideMaster.COLORS["primary"]
    p_title.space_after = FontEngine.pt(14.0)

    # Subtitle
    if payload.subtitle:
        sub_budget_h = Length(int(COVER_HEIGHT * 0.35))
        p_sub = tf_title.add_paragraph()
        p_sub.text = payload.subtitle
        p_sub.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p_sub.font.size = FontEngine.calculate_font_size(
            payload.subtitle,
            COVER_WIDTH,
            sub_budget_h,
            max_pt=16.0,
            min_pt=10.0,
        )
        p_sub.font.color.rgb = SlideMaster.COLORS["secondary"]

    # 2. Presenter and Date Metadata Box
    meta_box = BoundingBox(
        left=COVER_LEFT,
        top=META_OFFSET_TOP,
        width=COVER_WIDTH,
        height=META_HEIGHT,
    )
    tb_meta = ShapeBuilder.add_textbox(slide, meta_box)
    tf_meta = tb_meta.text_frame

    meta_parts: list[str] = []
    if payload.presenter:
        meta_parts.append(payload.presenter)
    if payload.date_str:
        meta_parts.append(payload.date_str)

    if meta_parts:
        meta_str = "  |  ".join(meta_parts)
        p_meta = tf_meta.paragraphs[0]
        p_meta.text = meta_str
        p_meta.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p_meta.font.size = FontEngine.calculate_font_size(
            meta_str,
            COVER_WIDTH,
            META_HEIGHT,
            max_pt=11.0,
            min_pt=9.0,
        )
        p_meta.font.bold = True
        p_meta.font.color.rgb = SlideMaster.COLORS["text_muted"]
