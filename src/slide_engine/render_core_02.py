"""Renderer for Agenda and Section Divider slide (CORE-02).

Renders numbered agenda items with dynamic vertical distribution
and active section highlighting.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Length
from pydantic import Field, model_validator

from slide_engine.core_base import (
    BoundingBox,
    ParagraphSpec,
    ShapeBuilder,
    SlideMaster,
    clamp_length,
)
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

    @model_validator(mode="after")
    def validate_agenda_density(self) -> Core02Payload:
        """Validate character density across items to guarantee legibility and physical containment."""
        if len(self.items) >= 7:
            for idx, item in enumerate(self.items):
                if item.description and len(item.description) > 100:
                    raise ValueError(
                        f"Agenda item {idx + 1} description exceeds 100 characters ({len(item.description)}) "
                        f"for dense {len(self.items)}-item agenda."
                    )
        return self


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
    {
        "items": [
            {
                "number": f"0{i}",
                "title": f"フェーズ {i}",
                "description": "長文" * 60,  # 120 chars > 100 chars limit for 7+ items
            }
            for i in range(1, 8)
        ]
    },
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
    ShapeBuilder.add_card(
        slide, row_box, bg_color=bg_col, border_color=border_col, border_width_pt=1.0
    )

    pad_x = clamp_length(PADDING_CONTENT, max_val=row_box.width * 0.05)
    pad_y = clamp_length(Inches(0.020), max_val=row_box.height * 0.035)
    badge_h = clamp_length(Inches(0.420), max_val=row_box.height - pad_y * 2)
    badge_w = clamp_length(BADGE_WIDTH, max_val=row_box.width * 0.15)

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
    title_col_key = "text_light" if item.is_active else "text_dark"
    desc_col_key = "bg_light" if item.is_active else "text_muted"

    ShapeBuilder.add_styled_paragraph_stack(
        slide,
        content_box,
        [
            ParagraphSpec(
                text=item.title,
                weight=0.55,
                max_pt=14.0,
                min_pt=8.0,
                color_key=title_col_key,
                bold=True,
            ),
            ParagraphSpec(
                text=item.description,
                weight=0.42,
                max_pt=10.0,
                min_pt=8.0,
                color_key=desc_col_key,
            ),
        ],
    )


def render_core_02(slide: Slide, payload: Core02Payload, box: BoundingBox) -> None:
    """Render numbered agenda items with active highlighting.

    Args:
        slide: Target slide.
        payload: Core02Payload instance.
        box: Content BoundingBox.
    """
    num_items = len(payload.items)
    max_gap_ratio = 0.008 if num_items >= 7 else (0.015 if num_items >= 5 else 0.03)
    gap = Length(min(int(GAP_BETWEEN_ROWS), int(box.height * max_gap_ratio)))
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
