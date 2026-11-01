"""Renderer for Cover slide (CORE-01).

Renders presentation title, subtitle, presenter, and date.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.util import Inches

from slide_engine.core_base import BoundingBox, ShapeBuilder
from slide_engine.models import BasePayload, SafeText, ShortText, TemplateSpec

if TYPE_CHECKING:
    from pptx.slide import Slide


class Core01Payload(BasePayload):
    """Payload for Cover slide (CORE-01)."""

    title: ShortText
    subtitle: SafeText | None = None
    presenter: ShortText | None = None
    date_str: ShortText | None = None


_SAMPLE_CORE01: dict[str, Any] = {
    "title": "次世代エンタープライズ統合データプラットフォーム刷新計画および移行戦略ロードマップ策定",
    "subtitle": "2026年度〜2028年度 中期経営計画における基幹系データベースとデータレイクハウスの完全統合アーキテクチャ設計書",
    "presenter": "デジタル変革推進本部 エンタープライズアーキテクチャグループ",
    "date_str": "2026年11月1日 (最終版)",
}

_INVALID_SAMPLES_CORE01: list[dict[str, Any]] = [
    {"title": "A" * 150},  # Exceeds MAX_SHORT_TEXT_LEN (120)
]

SPEC = TemplateSpec(
    payload_type=Core01Payload,
    supports_split=False,
    description="Cover Slide",
    sample_payload=_SAMPLE_CORE01,
    invalid_sample_payloads=_INVALID_SAMPLES_CORE01,
    renders_header=False,
    layout_type="cover",
    is_fallback=False,
)

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

    title_budget_h = title_box.scale_height(0.55) if payload.subtitle else COVER_HEIGHT

    # Main Title
    ShapeBuilder.add_styled_paragraph(
        tf_title,
        payload.title,
        COVER_WIDTH,
        title_budget_h,
        max_pt=32.0,
        min_pt=18.0,
        color_key="primary",
        bold=True,
        space_after_pt=14.0,
        is_first=True,
    )

    # Subtitle
    if payload.subtitle:
        sub_budget_h = title_box.scale_height(0.35)
        ShapeBuilder.add_styled_paragraph(
            tf_title,
            payload.subtitle,
            COVER_WIDTH,
            sub_budget_h,
            max_pt=16.0,
            min_pt=10.0,
            color_key="secondary",
        )

    # 2. Presenter and Date Metadata Box
    meta_parts: list[str] = []
    if payload.presenter:
        meta_parts.append(payload.presenter)
    if payload.date_str:
        meta_parts.append(payload.date_str)

    if meta_parts:
        meta_box = BoundingBox(
            left=COVER_LEFT,
            top=META_OFFSET_TOP,
            width=COVER_WIDTH,
            height=META_HEIGHT,
        )
        ShapeBuilder.add_styled_textbox(
            slide,
            meta_box,
            "  |  ".join(meta_parts),
            max_pt=11.0,
            min_pt=9.0,
            color_key="text_muted",
            bold=True,
        )
