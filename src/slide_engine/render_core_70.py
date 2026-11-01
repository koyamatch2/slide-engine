"""Renderer for Comparison slide (CORE-70).

Renders two contrasting columns (Before/After, Pros/Cons, Current/Target)
with distinct headers, bullets, and optional highlight accents. Automatically
switches to vertical top/bottom stacking inside narrow split panes.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.util import Inches, Length
from pydantic import Field

from slide_engine.core_base import BoundingBox, ShapeBuilder, SlideMaster
from slide_engine.models import BasePayload, SafeText, ShortText, TemplateSpec

if TYPE_CHECKING:
    from pptx.slide import Slide


class ComparisonColumn(BasePayload):
    """Column definition for Comparison slide."""

    header: ShortText
    sub_header: ShortText | None = None
    points: list[SafeText] = Field(min_length=1, max_length=6)
    is_highlighted: bool = False


class Core70Payload(BasePayload):
    """Payload for Comparison / Before-After slide (CORE-70)."""

    left_column: ComparisonColumn
    right_column: ComparisonColumn

    def validate_split_density(self, pane_name: str, template_id: str) -> None:
        """Enforce point count limits for split layout."""
        max_pts = max(len(self.left_column.points), len(self.right_column.points))
        if max_pts > 4:
            raise ValueError(
                f"{pane_name} comparison '{template_id}' in split layout exceeds legibility limit: "
                f"{max_pts} points > 4 points maximum."
            )


_SAMPLE_CORE70: dict[str, Any] = {
    "left_column": {
        "header": "現状課題 (As-Is)",
        "sub_header": "サイロ化された個別最適システム",
        "points": [f"部門間でのデータ連携が手動バッチに依存し遅延が発生 (課題{i})" for i in range(1, 7)],
        "is_highlighted": False,
    },
    "right_column": {
        "header": "理想姿 (To-Be)",
        "sub_header": "イベント駆動型リアルタイム統合基盤",
        "points": [f"ゼロトラストセキュリティ下での全社リアルタイムAPI連携の実現 (施策{i})" for i in range(1, 7)],
        "is_highlighted": True,
    },
}

_SAMPLE_CORE70_SPLIT: dict[str, Any] = {
    "left_column": {
        "header": "現状課題 (As-Is)",
        "sub_header": "サイロ化された個別最適",
        "points": [f"部門間データ連携が手動バッチに依存 (課題{i})" for i in range(1, 5)],
        "is_highlighted": False,
    },
    "right_column": {
        "header": "理想姿 (To-Be)",
        "sub_header": "リアルタイム統合基盤",
        "points": [f"全社リアルタイムAPI連携の実現 (施策{i})" for i in range(1, 5)],
        "is_highlighted": True,
    },
}

_INVALID_SAMPLES_CORE70: list[dict[str, Any]] = [
    {
        "left_column": {"header": "H", "points": []},  # Violates min_length=1
        "right_column": {"header": "H", "points": ["P"]},
    }
]

SPEC = TemplateSpec(
    payload_type=Core70Payload,
    supports_split=True,
    description="Comparison Box",
    sample_payload=_SAMPLE_CORE70,
    split_sample_payload=_SAMPLE_CORE70_SPLIT,
    invalid_sample_payloads=_INVALID_SAMPLES_CORE70,
)


COL_GAP = Inches(0.250)
COL_INNER_PAD = Inches(0.240)
MIN_SIDE_BY_SIDE_WIDTH = Inches(4.200)


def _render_comparison_column(
    slide: Slide,
    column: ComparisonColumn,
    box: BoundingBox,
) -> None:
    """Render a single comparison column card with adaptive typography."""
    if column.is_highlighted:
        bg_color = SlideMaster.COLORS["bg_white"]
        border_color = SlideMaster.COLORS["accent"]
        border_width = 1.5
    else:
        bg_color = SlideMaster.COLORS["bg_light"]
        border_color = SlideMaster.COLORS["border"]
        border_width = 1.0

    ShapeBuilder.add_card(
        slide,
        box,
        bg_color=bg_color,
        border_color=border_color,
        border_width_pt=border_width,
    )

    inner_box = box.inset(COL_INNER_PAD, COL_INNER_PAD, max_ratio_x=0.08, max_ratio_y=0.08)
    tb = ShapeBuilder.add_textbox(slide, inner_box)
    tf = tb.text_frame

    has_sub = bool(column.sub_header)
    header_budget_h = inner_box.scale_height(0.15 if has_sub else 0.20)
    sub_budget_h = inner_box.scale_height(0.12)
    points_budget_ratio = 0.60 if has_sub else 0.70
    points_count = max(1, len(column.points))
    point_budget_h = Length(int(inner_box.scale_height(points_budget_ratio) / points_count))
    space_gap_pt = min(6.0, max(1.5, inner_box.height.pt * 0.02))

    # Header Title
    header_col_key = "accent" if column.is_highlighted else "primary"
    ShapeBuilder.add_styled_paragraph(
        tf,
        column.header,
        inner_box.width,
        header_budget_h,
        max_pt=15.0,
        min_pt=9.0,
        color_key=header_col_key,
        bold=True,
        space_after_pt=None if column.sub_header else (space_gap_pt * 1.2),
        is_first=True,
    )

    # Sub-header
    if column.sub_header:
        ShapeBuilder.add_styled_paragraph(
            tf,
            column.sub_header,
            inner_box.width,
            sub_budget_h,
            max_pt=10.5,
            min_pt=8.0,
            color_key="text_muted",
            space_after_pt=space_gap_pt * 1.2,
        )

    # Point Bullet List
    for point in column.points:
        ShapeBuilder.add_styled_paragraph(
            tf,
            f"•  {point}",
            inner_box.width,
            point_budget_h,
            max_pt=11.0,
            min_pt=8.0,
            color_key="text_dark",
            space_after_pt=space_gap_pt * 0.6,
        )


def render_core_70(slide: Slide, payload: Core70Payload, box: BoundingBox) -> None:
    """Render dual-column or vertically stacked comparison boxes.

    Args:
        slide: Target slide.
        payload: Core70Payload instance.
        box: Content BoundingBox.
    """
    if box.width >= MIN_SIDE_BY_SIDE_WIDTH:
        gap = Length(min(int(COL_GAP), int(box.width * 0.05)))
        col_width = Length(int((box.width - gap) / 2))
        first_box = BoundingBox(box.left, box.top, col_width, box.height)
        second_box = BoundingBox(Length(box.left + col_width + gap), box.top, col_width, box.height)
    else:
        gap = Length(min(int(COL_GAP), int(box.height * 0.05)))
        row_height = Length(int((box.height - gap) / 2))
        first_box = BoundingBox(box.left, box.top, box.width, row_height)
        second_box = BoundingBox(box.left, Length(box.top + row_height + gap), box.width, row_height)

    _render_comparison_column(slide, payload.left_column, first_box)
    _render_comparison_column(slide, payload.right_column, second_box)
