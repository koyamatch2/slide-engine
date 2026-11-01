"""Renderer for Visual Guide Frame slide (CORE-50).

Renders a manual placement guide box (placeholder frame) with dashed border,
subtle background, placeholder label, and caption. Completely eliminates
external image I/O and external image parsing dependencies.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Length

from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import Core50Payload

if TYPE_CHECKING:
    from pptx.slide import Slide


CAPTION_HEIGHT = Inches(0.500)
CAPTION_GAP = Inches(0.150)


def render_core_50(slide: Slide, payload: Core50Payload, box: BoundingBox) -> None:
    """Render visual guide frame for manual screenshot/diagram paste.

    Args:
        slide: Target slide.
        payload: Core50Payload instance.
        box: Content BoundingBox.
    """
    has_caption = bool(payload.caption)
    caption_offset = CAPTION_HEIGHT + CAPTION_GAP if has_caption else Inches(0)
    frame_height = Length(box.height - caption_offset)

    frame_box = BoundingBox(
        left=box.left,
        top=box.top,
        width=box.width,
        height=frame_height,
    )

    # 1. Guide Frame Shape
    ShapeBuilder.add_card(
        slide,
        frame_box,
        bg_color=SlideMaster.COLORS["bg_light"],
        border_color=SlideMaster.COLORS["border"],
        border_width_pt=1.5,
    )

    # 2. Centered Helper Text
    tb_center = ShapeBuilder.add_textbox(slide, frame_box)
    tf_center = tb_center.text_frame
    p_center = tf_center.paragraphs[0]
    center_text = f"[ 画像・図解配置エリア: {payload.placeholder_label} ]"
    p_center.text = center_text
    p_center.alignment = PP_ALIGN.CENTER
    p_center.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_center.font.size = FontEngine.calculate_font_size(
        center_text,
        frame_box.width,
        frame_box.height,
        max_pt=13.0,
        min_pt=9.0,
    )
    p_center.font.bold = True
    p_center.font.color.rgb = SlideMaster.COLORS["text_muted"]

    # 3. Optional Bottom Caption
    if has_caption and payload.caption:
        caption_box = BoundingBox(
            left=box.left,
            top=Length(box.top + frame_height + CAPTION_GAP),
            width=box.width,
            height=CAPTION_HEIGHT,
        )
        tb_caption = ShapeBuilder.add_textbox(slide, caption_box)
        p_cap = tb_caption.text_frame.paragraphs[0]
        cap_text = f"※ {payload.caption}"
        p_cap.text = cap_text
        p_cap.alignment = PP_ALIGN.LEFT
        p_cap.font.name = SlideMaster.FONT_FAMILY_PRIMARY
        p_cap.font.size = FontEngine.calculate_font_size(
            cap_text,
            caption_box.width,
            caption_box.height,
            max_pt=9.5,
            min_pt=8.0,
        )
        p_cap.font.color.rgb = SlideMaster.COLORS["text_muted"]
