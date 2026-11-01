"""Renderer for Process Flow slide (CORE-20).

Renders sequential timeline or procedural steps with badges,
connector indicators, and cards. Automatically adapts from single-row
horizontal flow to multi-row grid when rendered inside narrow split panes.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pptx.util import Inches, Length
from pydantic import Field

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


class ProcessStep(BasePayload):
    """Step definition for Process Flow."""

    step_number: int = Field(ge=1, le=10)
    title: ShortText
    description: SafeText


class Core20Payload(BasePayload):
    """Payload for Process Flow slide (CORE-20)."""

    steps: list[ProcessStep] = Field(min_length=2, max_length=6)

    def validate_split_density(self, pane_name: str, template_id: str) -> None:
        """Enforce step count limits for split layout."""
        if len(self.steps) > 4:
            raise ValueError(
                f"{pane_name} process flow '{template_id}' in split layout exceeds legibility limit: "
                f"{len(self.steps)} steps > 4 steps maximum."
            )


_SAMPLE_CORE20: dict[str, Any] = {
    "steps": [
        {
            "step_number": i,
            "title": f"工程 {i}: 基盤構築",
            "description": f"マイクロサービス第{i}群の自動デプロイと統合監視の確立",
        }
        for i in range(1, 7)
    ]
}

_SAMPLE_CORE20_SPLIT: dict[str, Any] = {
    "steps": [
        {
            "step_number": i,
            "title": f"工程 {i}: 基盤構築",
            "description": f"マイクロサービス第{i}群の自動デプロイ",
        }
        for i in range(1, 5)
    ]
}

_INVALID_SAMPLES_CORE20: list[dict[str, Any]] = [
    {"steps": [{"step_number": 1, "title": "A", "description": "B"}]},  # Violates min_length=2
]

SPEC = TemplateSpec(
    payload_type=Core20Payload,
    supports_split=True,
    description="Process Flow",
    sample_payload=_SAMPLE_CORE20,
    split_sample_payload=_SAMPLE_CORE20_SPLIT,
    invalid_sample_payloads=_INVALID_SAMPLES_CORE20,
)


# Process Flow Layout Constants
GAP_BETWEEN_STEPS = Inches(0.180)
CARD_PADDING_X = Inches(0.180)
CARD_PADDING_Y = Inches(0.150)
BADGE_HEIGHT = Inches(0.320)
MIN_STEP_COL_WIDTH = Inches(1.650)
COMPACT_CARD_HEIGHT_THRESHOLD = Inches(1.800)


def _render_step_content(
    tf: Any,
    title: str,
    description: str,
    width: Length,
    height: Length,
    *,
    title_max_pt: float,
    space_after_pt: float,
) -> None:
    """Render standardized step card title and description."""
    ShapeBuilder.populate_paragraph_stack(
        tf,
        width,
        height,
        [
            ParagraphSpec(
                text=title,
                weight=0.40,
                max_pt=title_max_pt,
                min_pt=8.0,
                color_key="primary",
                bold=True,
                space_after_pt=space_after_pt,
            ),
            ParagraphSpec(
                text=description,
                weight=0.55,
                max_pt=10.0,
                min_pt=8.0,
                color_key="text_dark",
            ),
        ],
    )


def _render_compact_step_card(
    slide: Slide,
    step: ProcessStep,
    step_box: BoundingBox,
) -> None:
    """Render a compact step card with inline badge header to prevent shape height collapse."""
    inner_box = step_box.inset(CARD_PADDING_X, CARD_PADDING_Y, max_ratio_x=0.08, max_ratio_y=0.08)
    tb = ShapeBuilder.add_textbox(slide, inner_box)
    full_title = f"[STEP {step.step_number:02d}] {step.title}"
    _render_step_content(
        tb.text_frame,
        full_title,
        step.description,
        inner_box.width,
        inner_box.height,
        title_max_pt=11.5,
        space_after_pt=2.5,
    )


def _render_standard_step_card(
    slide: Slide,
    step: ProcessStep,
    step_box: BoundingBox,
) -> None:
    """Render a standard step card with dedicated badge box."""
    pad_x = clamp_length(CARD_PADDING_X, max_val=step_box.width * 0.08)
    pad_y = clamp_length(CARD_PADDING_Y, max_val=step_box.height * 0.08)
    inner_w = clamp_length(step_box.width - (pad_x * 2), min_val=Inches(0.400))
    badge_h = clamp_length(BADGE_HEIGHT, max_val=step_box.height * 0.22)
    inner_gap = clamp_length(Inches(0.060), max_val=step_box.height * 0.04)

    # Number Badge Header Box
    badge_box = BoundingBox(
        left=Length(step_box.left + pad_x),
        top=Length(step_box.top + pad_y),
        width=inner_w,
        height=badge_h,
    )
    ShapeBuilder.add_styled_textbox(
        slide,
        badge_box,
        f"STEP {step.step_number:02d}",
        max_pt=11.0,
        min_pt=8.0,
        color_key="secondary",
        bold=True,
    )

    # Step Title & Description Box
    content_top = Length(step_box.top + pad_y + badge_h + inner_gap)
    content_h = clamp_length(step_box.bottom - content_top - pad_y, min_val=Inches(0.300))
    content_box = BoundingBox(
        left=Length(step_box.left + pad_x),
        top=content_top,
        width=inner_w,
        height=content_h,
    )
    tb = ShapeBuilder.add_textbox(slide, content_box)
    _render_step_content(
        tb.text_frame,
        step.title,
        step.description,
        content_box.width,
        content_box.height,
        title_max_pt=12.5,
        space_after_pt=4.0,
    )


def _render_process_step_card(
    slide: Slide,
    step: ProcessStep,
    step_box: BoundingBox,
) -> None:
    """Render an individual process step card and its typography."""
    ShapeBuilder.add_card(
        slide,
        step_box,
        bg_color=SlideMaster.COLORS["bg_light"],
        border_color=SlideMaster.COLORS["border"],
        border_width_pt=1.0,
    )

    if step_box.height < COMPACT_CARD_HEIGHT_THRESHOLD:
        _render_compact_step_card(slide, step, step_box)
    else:
        _render_standard_step_card(slide, step, step_box)


def render_core_20(slide: Slide, payload: Core20Payload, box: BoundingBox) -> None:
    """Render sequential steps using adaptive row/column grid layout.

    Args:
        slide: Target slide.
        payload: Core20Payload instance.
        box: Content BoundingBox.
    """
    grid = box.adaptive_grid(
        num_items=len(payload.steps),
        min_col_width=MIN_STEP_COL_WIDTH,
        gap_x=GAP_BETWEEN_STEPS,
        gap_y=GAP_BETWEEN_STEPS,
        balance_rows=False,
    )

    for step, step_box in zip(payload.steps, grid, strict=False):
        _render_process_step_card(slide, step, step_box)
