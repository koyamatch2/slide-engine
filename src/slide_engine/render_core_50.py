"""Renderer for Visual Guide Frame slide (CORE-50).

Renders a manual placement guide box (placeholder frame) with dashed border,
subtle background, placeholder label, and caption. Completely eliminates
external image I/O and external image parsing dependencies.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Length

from slide_engine.core_base import BoundingBox, ShapeBuilder, SlideMaster
from slide_engine.models import BasePayload, SafeText, ShortText, TemplateSpec

if TYPE_CHECKING:
    from pptx.slide import Slide


class Core50Payload(BasePayload):
    """Payload for Visual Guide Frame slide (CORE-50).

    Used for rendering a manual screenshot/diagram guide frame and caption.
    Completely eliminates external file I/O and Pillow dependency.
    """

    placeholder_label: ShortText
    caption: SafeText | None = None


_SAMPLE_CORE50: dict[str, Any] = {
    "placeholder_label": "ハイブリッドクラウド冗長化構成および閉域網ネットワークトポロジー図",
    "caption": "図1: 東西リージョン間でのアクティブ・アクティブ構成と自動フェイルオーバー経路の定義",
}

_INVALID_SAMPLES_CORE50: list[dict[str, Any]] = [
    {"placeholder_label": "A" * 150},  # Exceeds MAX_SHORT_TEXT_LEN (120)
]

SPEC = TemplateSpec(
    payload_type=Core50Payload,
    supports_split=True,
    description="Visual Guide Frame",
    sample_payload=_SAMPLE_CORE50,
    invalid_sample_payloads=_INVALID_SAMPLES_CORE50,
)


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
    ShapeBuilder.add_styled_textbox(
        slide,
        frame_box,
        f"[ 画像・図解配置エリア: {payload.placeholder_label} ]",
        max_pt=13.0,
        min_pt=9.0,
        color_key="text_muted",
        bold=True,
        alignment=PP_ALIGN.CENTER,
    )

    # 3. Optional Bottom Caption
    if has_caption and payload.caption:
        caption_box = BoundingBox(
            left=box.left,
            top=Length(box.top + frame_height + CAPTION_GAP),
            width=box.width,
            height=CAPTION_HEIGHT,
        )
        ShapeBuilder.add_styled_textbox(
            slide,
            caption_box,
            f"※ {payload.caption}",
            max_pt=9.5,
            min_pt=8.0,
            color_key="text_muted",
            alignment=PP_ALIGN.LEFT,
        )
