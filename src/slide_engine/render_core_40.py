"""Renderer for 2x2 Matrix slide (CORE-40).

Renders a four-quadrant positioning map with custom X/Y axis labels,
distinct card backgrounds, and category badges.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Length

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


class MatrixQuadrant(BasePayload):
    """Quadrant definition for 2x2 Matrix."""

    title: ShortText
    description: SafeText
    badge: ShortText | None = None


class Core40Payload(BasePayload):
    """Payload for 2x2 Matrix slide (CORE-40)."""

    x_axis_label: ShortText
    y_axis_label: ShortText
    top_left: MatrixQuadrant
    top_right: MatrixQuadrant
    bottom_left: MatrixQuadrant
    bottom_right: MatrixQuadrant


_SAMPLE_CORE40: dict[str, Any] = {
    "x_axis_label": "業務定着コスト・移行難易度",
    "y_axis_label": "全社財務インパクト・戦略貢献度",
    "top_left": {
        "title": "重点投資領域",
        "badge": "P1",
        "description": "中長期的な競争優位性を確立するための次世代AIプラットフォーム投資。",
    },
    "top_right": {
        "title": "最優先推進（コア）",
        "badge": "P0",
        "description": "直ちに全社展開し収益拡大を牽引するミッションクリティカル施策。",
    },
    "bottom_left": {
        "title": "段階的縮退・廃止",
        "badge": "P3",
        "description": "保守コストが高止まりしている老朽化オンプレミス資産の計画的廃棄。",
    },
    "bottom_right": {
        "title": "クイックウィン",
        "badge": "P2",
        "description": "既存業務プロセスの自動化により短期で投資回収が可能な改善施策。",
    },
}

_INVALID_SAMPLES_CORE40: list[dict[str, Any]] = [
    {"x_axis_label": "X"},  # Missing required quadrant fields
]

SPEC = TemplateSpec(
    payload_type=Core40Payload,
    supports_split=True,
    description="2x2 Matrix",
    sample_payload=_SAMPLE_CORE40,
    invalid_sample_payloads=_INVALID_SAMPLES_CORE40,
)


# Matrix Geometry Constants
AXIS_LABEL_SIZE = Inches(0.350)
GRID_GAP = Inches(0.180)
QUADRANT_PADDING = Inches(0.180)


def _render_quadrant_card(
    slide: Slide,
    quadrant: MatrixQuadrant,
    box: BoundingBox,
) -> None:
    """Render a single quadrant card with title, badge, and description."""
    ShapeBuilder.add_card(
        slide,
        box,
        bg_color=SlideMaster.COLORS["bg_light"],
        border_color=SlideMaster.COLORS["border"],
        border_width_pt=1.0,
    )

    inner_box = box.inset(QUADRANT_PADDING, QUADRANT_PADDING, max_ratio_x=0.08, max_ratio_y=0.08)
    header_text = f"[{quadrant.badge}] {quadrant.title}" if quadrant.badge else quadrant.title
    ShapeBuilder.add_styled_paragraph_stack(
        slide,
        inner_box,
        [
            ParagraphSpec(
                text=header_text,
                weight=0.35,
                max_pt=12.5,
                min_pt=8.5,
                color_key="primary",
                bold=True,
                space_after_pt=4.0,
            ),
            ParagraphSpec(
                text=quadrant.description,
                weight=0.60,
                max_pt=10.0,
                min_pt=8.0,
                color_key="text_dark",
            ),
        ],
    )


def render_core_40(slide: Slide, payload: Core40Payload, box: BoundingBox) -> None:
    """Render 2x2 matrix with axis headers.

    Args:
        slide: Target slide.
        payload: Core40Payload instance.
        box: Content BoundingBox.
    """
    # 1. Axis Label Zones
    axis_h = clamp_length(AXIS_LABEL_SIZE, max_val=box.height * 0.10)
    axis_box = BoundingBox(box.left, box.top, box.width, axis_h)
    ShapeBuilder.add_styled_textbox(
        slide,
        axis_box,
        f"◀ {payload.x_axis_label}  |  {payload.y_axis_label} ▶",
        max_pt=11.0,
        min_pt=8.0,
        color_key="secondary",
        bold=True,
        alignment=PP_ALIGN.CENTER,
    )

    # 2. Quadrants Layout Calculation via sub_grid
    gap_x = clamp_length(GRID_GAP, max_val=box.width * 0.04)
    gap_y = clamp_length(GRID_GAP, max_val=box.height * 0.04)
    top_offset = clamp_length(Inches(0.080), max_val=box.height * 0.02)

    quad_top = Length(box.top + axis_h + top_offset)
    quad_area_h = Length(box.height - axis_h - top_offset)
    quad_area = BoundingBox(box.left, quad_top, box.width, quad_area_h)
    grid = quad_area.sub_grid(2, 2, gap_x, gap_y)

    quadrants = [payload.top_left, payload.top_right, payload.bottom_left, payload.bottom_right]
    for quad_data, cell_box in zip(quadrants, grid, strict=True):
        _render_quadrant_card(slide, quad_data, cell_box)
