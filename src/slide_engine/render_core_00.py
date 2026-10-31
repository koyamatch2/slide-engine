"""Renderer for System Fallback slide (CORE-00).

Provides guaranteed plain-text fail-soft rendering when unhandled
exceptions occur during normal slide generation.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.util import Inches, Length

from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import Core00Payload

if TYPE_CHECKING:
    from pptx.slide import Slide


# Fallback Card Geometry Constants
PADDING_X = Inches(0.400)
PADDING_Y = Inches(0.350)
HEADER_OFFSET_Y = Inches(0.600)


def render_core_00(slide: Slide, payload: Core00Payload, box: BoundingBox) -> None:
    """Render fail-soft fallback error container and diagnostics.

    Args:
        slide: Target slide.
        payload: Core00Payload instance containing error details.
        box: Content BoundingBox.
    """
    # 1. Background warning container card
    ShapeBuilder.add_card(
        slide,
        box,
        bg_color=SlideMaster.COLORS["fallback_bg"],
        border_color=SlideMaster.COLORS["fallback_text"],
        border_width_pt=1.5,
    )

    # 2. Error Title & Message Text Box
    text_box = BoundingBox(
        left=Length(box.left + PADDING_X),
        top=Length(box.top + PADDING_Y),
        width=Length(box.width - (PADDING_X * 2)),
        height=Length(box.height - (PADDING_Y * 2)),
    )
    tb = ShapeBuilder.add_textbox(slide, text_box)
    tf = tb.text_frame

    # Error Badge / Title
    p_title = tf.paragraphs[0]
    p_title.text = f"[SYSTEM FALLBACK] {payload.error_title}"
    p_title.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_title.font.size = FontEngine.pt(16.0)
    p_title.font.bold = True
    p_title.font.color.rgb = SlideMaster.COLORS["fallback_text"]
    p_title.space_after = FontEngine.pt(8.0)

    # Error Description
    p_desc = tf.add_paragraph()
    p_desc.text = f"Cause: {payload.error_message}"
    p_desc.font.name = SlideMaster.FONT_FAMILY_PRIMARY
    p_desc.font.size = FontEngine.pt(11.0)
    p_desc.font.bold = True
    p_desc.font.color.rgb = SlideMaster.COLORS["text_dark"]
    p_desc.space_after = FontEngine.pt(12.0)

    # Raw Payload Diagnostics Dump
    if payload.raw_payload_dump:
        p_dump_label = tf.add_paragraph()
        p_dump_label.text = "Payload Diagnostics Snapshot:"
        p_dump_label.font.name = SlideMaster.FONT_FAMILY_SECONDARY
        p_dump_label.font.size = FontEngine.pt(9.5)
        p_dump_label.font.bold = True
        p_dump_label.font.color.rgb = SlideMaster.COLORS["text_muted"]

        p_dump = tf.add_paragraph()
        p_dump.text = payload.raw_payload_dump[:1000]
        p_dump.font.name = SlideMaster.FONT_FAMILY_SECONDARY
        p_dump.font.size = FontEngine.pt(8.5)
        p_dump.font.color.rgb = SlideMaster.COLORS["text_muted"]
