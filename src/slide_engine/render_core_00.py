"""Renderer for System Fallback slide (CORE-00).

Provides guaranteed plain-text fail-soft rendering when unhandled
exceptions occur during normal slide generation.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.util import Inches
from pydantic import Field

from slide_engine.core_base import BoundingBox, ParagraphSpec, ShapeBuilder, SlideMaster
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
    dump_text = payload.raw_payload_dump[:1000] if payload.raw_payload_dump else None
    ShapeBuilder.add_styled_paragraph_stack(
        slide,
        text_box,
        [
            ParagraphSpec(
                text=f"[SYSTEM FALLBACK] {payload.error_title}",
                weight=0.18,
                max_pt=16.0,
                min_pt=12.0,
                color_key="fallback_text",
                bold=True,
                space_after_pt=8.0,
            ),
            ParagraphSpec(
                text=f"Cause: {payload.error_message}",
                weight=0.30,
                max_pt=11.0,
                min_pt=9.0,
                color_key="text_dark",
                bold=True,
                space_after_pt=12.0,
            ),
            ParagraphSpec(
                text="Payload Diagnostics Snapshot:" if dump_text else None,
                weight=0.10,
                max_pt=9.5,
                min_pt=9.5,
                color_key="text_muted",
                bold=True,
                font_family=SlideMaster.FONT_FAMILY_SECONDARY,
            ),
            ParagraphSpec(
                text=dump_text,
                weight=0.42,
                max_pt=8.5,
                min_pt=8.0,
                color_key="text_muted",
                font_family=SlideMaster.FONT_FAMILY_SECONDARY,
            ),
        ],
    )
