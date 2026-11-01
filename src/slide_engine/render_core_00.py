"""Renderer for System Fallback slide (CORE-00).

Provides guaranteed plain-text fail-soft rendering when unhandled
exceptions occur during normal slide generation.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.util import Inches
from pydantic import Field

from slide_engine.core_base import BoundingBox, ShapeBuilder, SlideMaster
from slide_engine.models import BasePayload, SafeText, ShortText, TemplateSpec

if TYPE_CHECKING:
    from pptx.slide import Slide


class Core00Payload(BasePayload):
    """Payload for system fallback slide (CORE-00)."""

    error_title: ShortText = Field(default="Slide Rendering Error")
    error_message: SafeText = Field(default="An unhandled exception occurred during rendering.")
    raw_payload_dump: str = Field(default="")


_SAMPLE_CORE00: dict[str, Any] = {
    "error_title": "システムエラー",
    "error_message": "スライドレンダリング処理中に例外が発生しました。",
    "raw_payload_dump": "{}",
}

SPEC = TemplateSpec(
    payload_type=Core00Payload,
    supports_split=False,
    description="System Fallback",
    sample_payload=_SAMPLE_CORE00,
    renders_header=False,
    layout_type="blank",
    is_fallback=True,
)

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
    text_box = box.inset(PADDING_X, PADDING_Y)
    tb = ShapeBuilder.add_textbox(slide, text_box)
    tf = tb.text_frame

    title_h = text_box.scale_height(0.18)
    has_dump = bool(payload.raw_payload_dump)
    desc_h = text_box.scale_height(0.22) if has_dump else text_box.scale_height(0.75)

    # Error Badge / Title
    ShapeBuilder.add_styled_paragraph(
        tf,
        f"[SYSTEM FALLBACK] {payload.error_title}",
        text_box.width,
        title_h,
        max_pt=16.0,
        min_pt=12.0,
        color_key="fallback_text",
        bold=True,
        space_after_pt=8.0,
        is_first=True,
    )

    # Error Description
    ShapeBuilder.add_styled_paragraph(
        tf,
        f"Cause: {payload.error_message}",
        text_box.width,
        desc_h,
        max_pt=11.0,
        min_pt=9.0,
        color_key="text_dark",
        bold=True,
        space_after_pt=12.0,
    )

    # Raw Payload Diagnostics Dump
    if has_dump:
        diag_hdr_h = text_box.scale_height(0.10)
        diag_body_h = text_box.scale_height(0.50)
        ShapeBuilder.add_styled_paragraph(
            tf,
            "Payload Diagnostics Snapshot:",
            text_box.width,
            diag_hdr_h,
            max_pt=9.5,
            min_pt=9.5,
            color_key="text_muted",
            bold=True,
            font_family=SlideMaster.FONT_FAMILY_SECONDARY,
        )
        ShapeBuilder.add_styled_paragraph(
            tf,
            payload.raw_payload_dump[:1000],
            text_box.width,
            diag_body_h,
            max_pt=8.5,
            min_pt=8.0,
            color_key="text_muted",
            font_family=SlideMaster.FONT_FAMILY_SECONDARY,
        )
