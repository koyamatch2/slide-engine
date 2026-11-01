"""Renderer for Agenda and Section Divider slide (CORE-02).

Renders numbered agenda items with dynamic vertical distribution
and active section highlighting.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Length
from pydantic import Field

from slide_engine.core_base import BoundingBox, ShapeBuilder, SlideMaster
from slide_engine.models import BasePayload, SafeText, ShortText, TemplateSpec

if TYPE_CHECKING:
    from pptx.slide import Slide


class AgendaItem(BasePayload):
    """Item for Agenda slide."""

    number: ShortText
    title: ShortText
    description: SafeText | None = None
    is_active: bool = False


class Core02Payload(BasePayload):
    """Payload for Agenda / Section slide (CORE-02)."""

    items: list[AgendaItem] = Field(min_length=1, max_length=8)


_SAMPLE_CORE02: dict[str, Any] = {
    "items": [
        {
            "number": f"0{i}",
            "title": f"戦略フェーズ {i}: 全社データ基盤の統合とガバナンス強化",
            "description": "レガシーシステムの段階的廃止とクラウドネイティブ基盤への完全移行",
            "is_active": i == 1,
        }
        for i in range(1, 9)
    ]
}

_INVALID_SAMPLES_CORE02: list[dict[str, Any]] = [
    {"items": []},  # Violates min_length=1
]

SPEC = TemplateSpec(
    payload_type=Core02Payload,
    supports_split=False,
    description="Agenda / Section",
    sample_payload=_SAMPLE_CORE02,
    invalid_sample_payloads=_INVALID_SAMPLES_CORE02,
    renders_header=True,
    layout_type="agenda",
)


# Agenda Geometry Constants
BADGE_WIDTH = Inches(0.850)
GAP_BETWEEN_ROWS = Inches(0.180)
PADDING_CONTENT = Inches(0.220)


def _render_agenda_row(
    slide: Slide,
    item: AgendaItem,
    row_box: BoundingBox,
) -> None:
    """Render an individual numbered agenda item card."""
    bg_col = SlideMaster.COLORS["primary"] if item.is_active else SlideMaster.COLORS["bg_light"]
    border_col = SlideMaster.COLORS["primary"] if item.is_active else SlideMaster.COLORS["border"]
    ShapeBuilder.add_card(slide, row_box, bg_color=bg_col, border_color=border_col, border_width_pt=1.0)

    pad_x = Length(min(int(PADDING_CONTENT), int(row_box.width * 0.05)))
    pad_y = Length(min(int(Inches(0.100)), int(row_box.height * 0.12)))
    badge_h = Length(min(int(Inches(0.420)), int(row_box.height - pad_y * 2)))
    badge_w = Length(min(int(BADGE_WIDTH), int(row_box.width * 0.15)))

    # 1. Number Badge
    badge_box = BoundingBox(
        left=Length(row_box.left + pad_x),
        top=Length(int(row_box.top + (row_box.height - badge_h) / 2)),
        width=badge_w,
        height=badge_h,
    )
    badge_col_key = "accent" if item.is_active else "secondary"
    ShapeBuilder.add_styled_textbox(
        slide,
        badge_box,
        str(item.number),
        max_pt=18.0,
        min_pt=10.0,
        color_key=badge_col_key,
        bold=True,
        alignment=PP_ALIGN.CENTER,
    )

    # 2. Title & Description Box
    content_left = Length(row_box.left + pad_x + badge_w + Inches(0.180))
    content_width = Length(max(int(Inches(0.500)), int(row_box.right - content_left - pad_x)))
    content_height = Length(max(int(Inches(0.200)), int(row_box.height - (pad_y * 2))))
    content_box = BoundingBox(
        left=content_left,
        top=Length(row_box.top + pad_y),
        width=content_width,
        height=content_height,
    )
    tb_content = ShapeBuilder.add_textbox(slide, content_box)
    tf_content = tb_content.text_frame

    title_budget_h = content_box.scale_height(0.55) if item.description else content_height
    title_col_key = "text_light" if item.is_active else "text_dark"
    ShapeBuilder.add_styled_paragraph(
        tf_content,
        item.title,
        content_box.width,
        title_budget_h,
        max_pt=14.0,
        min_pt=9.0,
        color_key=title_col_key,
        bold=True,
        is_first=True,
    )

    if item.description:
        desc_budget_h = content_box.scale_height(0.42)
        desc_col_key = "bg_light" if item.is_active else "text_muted"
        ShapeBuilder.add_styled_paragraph(
            tf_content,
            item.description,
            content_box.width,
            desc_budget_h,
            max_pt=10.0,
            min_pt=8.0,
            color_key=desc_col_key,
        )


def render_core_02(slide: Slide, payload: Core02Payload, box: BoundingBox) -> None:
    """Render numbered agenda items with active highlighting.

    Args:
        slide: Target slide.
        payload: Core02Payload instance.
        box: Content BoundingBox.
    """
    num_items = len(payload.items)
    gap = Length(min(int(GAP_BETWEEN_ROWS), int(box.height * 0.03)))
    total_gaps = gap * (num_items - 1)
    available_height = box.height - total_gaps
    row_height = Length(int(available_height / num_items))

    current_top = box.top

    for item in payload.items:
        row_box = BoundingBox(
            left=box.left,
            top=current_top,
            width=box.width,
            height=row_height,
        )
        _render_agenda_row(slide, item, row_box)
        current_top = Length(current_top + row_height + gap)
