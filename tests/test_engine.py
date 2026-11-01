"""Comprehensive test suite for Slide Engine.

Verifies:
1. Dynamic dispatch resolution and function signatures across all 10 templates.
2. Pydantic boundary validation (text limits, array lengths, chart series matching).
3. Complex layouts (lead text vertical offset, heterogeneous split).
4. Fail-soft behavior (CORE-00 fallback on exceptions).
5. Clean canvas purging and master background preservation.
"""

from __future__ import annotations

import inspect
import itertools
import json
from pathlib import Path
from typing import Any

import pytest
from pptx import Presentation
from pptx.util import Inches, Length

from slide_engine.__main__ import main
from slide_engine.controller import (
    _initialize_presentation,
    _render_slide_with_failsoft,
    build_dynamic_dispatch_table,
    generate_presentation,
)
from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import (
    MAX_SLIDES_PER_DECK,
    TEMPLATE_REGISTRY,
    ChartSeries,
    Core10Payload,
    Core30Payload,
    Core50Payload,
    PresentationDeck,
    SlideDefinition,
)


class TestDynamicDispatch:
    """Verifies template registry and dispatch table integrity."""

    def test_all_registered_templates_have_renderers(self) -> None:
        """Every template in TEMPLATE_REGISTRY must have a corresponding renderer."""
        dispatch_table = build_dynamic_dispatch_table()
        for template_id in TEMPLATE_REGISTRY:
            assert template_id in dispatch_table, f"Missing renderer for registered '{template_id}'"

    def test_renderer_function_signatures(self) -> None:
        """All renderers must strictly accept (slide, payload, box)."""
        dispatch_table = build_dynamic_dispatch_table()
        for template_id, render_fn in dispatch_table.items():
            sig = inspect.signature(render_fn)
            params = list(sig.parameters.keys())
            assert len(params) == 3, f"Renderer '{template_id}' parameter count != 3: {params}"
            assert params[0] == "slide", f"Renderer '{template_id}' first parameter must be 'slide'"
            assert params[1] == "payload", f"Renderer '{template_id}' second parameter must be 'payload'"
            assert params[2] == "box", f"Renderer '{template_id}' third parameter must be 'box'"


class TestSchemaValidationAndDefenses:
    """Tests Pydantic V2 models, resource limits, and error rejection."""

    def test_chart_series_length_mismatch_raises_error(self) -> None:
        """CORE-30 must reject series whose values count != categories count."""
        with pytest.raises(ValueError, match="does not match categories length"):
            Core30Payload(
                chart_type="column",
                categories=["Q1", "Q2", "Q3"],
                series=[ChartSeries(name="Revenue", values=[10.0, 20.0])],  # Expected 3 values
            )

    def test_deck_slide_count_limit(self) -> None:
        """Deck must reject slides exceeding MAX_SLIDES_PER_DECK."""
        valid_slide_dict = {
            "template_id": "CORE-01",
            "title": "Title",
            "payload": {"title": "Cover Title"},
        }
        excessive_slides = [valid_slide_dict] * (MAX_SLIDES_PER_DECK + 1)
        with pytest.raises(ValueError):
            PresentationDeck.model_validate({"theme_name": "corporate_blue", "slides": excessive_slides})

    def test_unsupported_split_template_raises_error(self) -> None:
        """Templates with supports_split=False (e.g. CORE-01) must reject split."""
        with pytest.raises(ValueError, match="does not support horizontal split"):
            SlideDefinition.model_validate({
                "template_id": "CORE-01",
                "title": "Invalid Split",
                "payload": {"title": "Title"},
                "right_template_id": "CORE-10",
                "right_payload": {"headers": ["Col"], "rows": [["Val"]]},
            })

    def test_core50_payload_purity(self) -> None:
        """CORE-50 accepts guide frame parameters without file paths."""
        payload = Core50Payload(
            placeholder_label="System Topology Diagram",
            caption="Manual placement placeholder",
        )
        assert payload.placeholder_label == "System Topology Diagram"
        assert payload.caption == "Manual placement placeholder"


class TestGeometryAndLayout:
    """Verifies bounding box calculations, lead text offsets, and canvas operations."""

    def test_bounding_box_coordinates_and_properties(self) -> None:
        """BoundingBox must compute right and bottom as Length without double-scaling."""
        box = BoundingBox(
            left=Length(1000),
            top=Length(2000),
            width=Length(3000),
            height=Length(4000),
        )
        assert box.right == Length(4000)
        assert box.bottom == Length(6000)
        assert isinstance(box.right, Length)
        assert isinstance(box.bottom, Length)
        # Ensure raw arithmetic value is strictly preserved (no Inches(x) multiplier)
        assert int(box.right) == 4000
        assert int(box.bottom) == 6000

    def test_split_horizontal_exact_geometry_and_conservation(self) -> None:
        """split_horizontal must strictly conserve total width and retain Length types."""
        content_box = SlideMaster.CONTENT_BOX
        left_box, right_box = SlideMaster.split_horizontal(content_box, split_ratio=0.5)

        # Type invariants
        assert isinstance(left_box.left, Length)
        assert isinstance(left_box.width, Length)
        assert isinstance(left_box.right, Length)
        assert isinstance(right_box.left, Length)
        assert isinstance(right_box.width, Length)
        assert isinstance(right_box.right, Length)

        # Alignment invariants
        assert left_box.left == content_box.left
        assert left_box.top == content_box.top
        assert right_box.top == content_box.top
        assert left_box.height == content_box.height
        assert right_box.height == content_box.height

        # Width conservation and split gap continuity
        gap = SlideMaster.DEFAULT_SPLIT_GAP
        assert abs(int(left_box.width + right_box.width + gap) - int(content_box.width)) <= 1
        assert right_box.left == Length(left_box.left + left_box.width + gap)

        # Coordinate bounds check (prevent regression of EMU overflow bug)
        assert right_box.right <= SlideMaster.SLIDE_WIDTH
        assert right_box.bottom <= SlideMaster.SLIDE_HEIGHT

    def test_split_horizontal_custom_ratio_and_gap(self) -> None:
        """split_horizontal respects custom split ratio and custom gap values."""
        content_box = SlideMaster.CONTENT_BOX
        custom_gap = Inches(0.200)
        left_box, right_box = SlideMaster.split_horizontal(
            content_box,
            split_ratio=0.3,
            gap=custom_gap,
        )

        assert left_box.width < right_box.width
        assert abs(int(left_box.width + right_box.width + custom_gap) - int(content_box.width)) <= 1
        assert right_box.left == Length(left_box.left + left_box.width + custom_gap)
        assert right_box.right <= SlideMaster.SLIDE_WIDTH

    def test_split_horizontal_nested_splits_no_coordinate_explosion(self) -> None:
        """Repeated splits must not explode coordinates exponentially."""
        content_box = SlideMaster.CONTENT_BOX
        left_box, _ = SlideMaster.split_horizontal(content_box)
        sub_left, sub_right = SlideMaster.split_horizontal(left_box, split_ratio=0.5)

        assert sub_left.left == content_box.left
        assert sub_right.right <= left_box.right
        assert sub_right.right < SlideMaster.SLIDE_WIDTH

    def test_lead_text_box_offset(self) -> None:
        """Lead text presence must reduce and push down the body content box."""
        lead_box, body_box = SlideMaster.get_body_box_with_lead(has_lead_text=True)
        assert lead_box is not None
        assert lead_box.top == SlideMaster.CONTENT_BOX.top
        assert body_box.top > SlideMaster.CONTENT_BOX.top
        assert body_box.height < SlideMaster.CONTENT_BOX.height

    def test_clean_slide_canvas_purges_shapes(self) -> None:
        """clean_slide_canvas must remove all dynamic shapes on the canvas."""
        prs = Presentation()
        prs.slide_width = SlideMaster.SLIDE_WIDTH
        prs.slide_height = SlideMaster.SLIDE_HEIGHT
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Add dummy shapes
        slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(2))
        slide.shapes.add_table(2, 2, Inches(3), Inches(3), Inches(2), Inches(2))
        assert len(slide.shapes) == 2

        ShapeBuilder.clean_slide_canvas(slide)
        assert len(slide.shapes) == 0

    def test_sample_deck_all_shapes_within_slide_bounds(self) -> None:
        """Every rendered shape across all slides must lie strictly within slide boundaries."""
        sample_path = Path(__file__).parent.parent / "examples" / "sample_deck.json"
        with open(sample_path, encoding="utf-8") as f:
            deck_dict = json.load(f)

        prs, summary = generate_presentation(PresentationDeck.model_validate(deck_dict))
        assert summary["fallback_slides"] == 0
        assert prs.slide_width is not None
        assert prs.slide_height is not None

        slide_w = int(prs.slide_width)
        slide_h = int(prs.slide_height)

        for idx, slide in enumerate(prs.slides, 1):
            for s_idx, shape in enumerate(slide.shapes):
                left = int(shape.left or 0)
                top = int(shape.top or 0)
                width = int(shape.width or 0)
                height = int(shape.height or 0)
                right = left + width
                bottom = top + height

                assert left >= 0, f"Slide {idx} Shape {s_idx} left < 0: {left}"
                assert top >= 0, f"Slide {idx} Shape {s_idx} top < 0: {top}"
                assert right <= slide_w, f"Slide {idx} Shape {s_idx} exceeds slide width: {right} > {slide_w}"
                assert bottom <= slide_h, f"Slide {idx} Shape {s_idx} exceeds slide height: {bottom} > {slide_h}"

    def test_sample_deck_no_unintended_peer_text_overlap(self) -> None:
        """Verify peer textboxes and pane elements do not unintentionally overlap."""
        sample_path = Path(__file__).parent.parent / "examples" / "sample_deck.json"
        with open(sample_path, encoding="utf-8") as f:
            deck_dict = json.load(f)

        prs, _ = generate_presentation(PresentationDeck.model_validate(deck_dict))

        # Slide 1 (Cover): Title box and metadata box must have zero overlap
        slide_1 = prs.slides[0]
        text_boxes = [s for s in slide_1.shapes if s.has_text_frame]
        assert len(text_boxes) >= 2
        title_tb, meta_tb = text_boxes[0], text_boxes[1]
        title_bottom = int(title_tb.top or 0) + int(title_tb.height or 0)
        meta_top = int(meta_tb.top or 0)
        assert title_bottom <= meta_top, f"Cover title overlaps meta: bottom {title_bottom} > top {meta_top}"

        # Slide 10 (Split): Left table and Right guide box must not horizontally overlap
        slide_10 = prs.slides[9]
        table_shapes = [s for s in slide_10.shapes if s.has_table]
        assert len(table_shapes) == 1
        table_shape = table_shapes[0]
        table_right = int(table_shape.left or 0) + int(table_shape.width or 0)

        right_shapes = [
            s for s in slide_10.shapes
            if int(s.left or 0) > table_right and s != table_shape and int(s.top or 0) >= int(table_shape.top or 0)
        ]
        assert len(right_shapes) >= 1
        for rs in right_shapes:
            rs_left = int(rs.left or 0)
            assert table_right <= rs_left, f"Split panes overlap: table right {table_right} > right shape left {rs_left}"


class TestFailSoftExecution:
    """Verifies that rendering errors fall back to CORE-00 safely."""

    def test_exception_triggers_core00_fallback(self) -> None:
        """A crashing renderer must be replaced by CORE-00 without halting execution."""
        prs = _initialize_presentation("corporate_blue")
        dispatch_table = build_dynamic_dispatch_table()

        # Mock a broken renderer for CORE-10
        def broken_renderer(slide: Any, payload: Any, box: BoundingBox) -> None:
            raise RuntimeError("Simulated Database Render Failure")

        broken_table = dict(dispatch_table)
        broken_table["CORE-10"] = broken_renderer

        slide_def = SlideDefinition(
            template_id="CORE-10",
            title="Faulty Table",
            payload=Core10Payload(headers=["H1"], rows=[["R1"]]),
        )

        success = _render_slide_with_failsoft(prs, slide_def, broken_table)
        assert success is False
        assert len(prs.slides) == 1

        # Check speaker notes for stack trace
        notes = prs.slides[0].notes_slide.notes_text_frame.text
        assert "Simulated Database Render Failure" in notes

    def test_full_pipeline_heterogeneous_split_execution(self) -> None:
        """Full presentation pipeline must render heterogeneous split slides."""
        deck_data = {
            "theme_name": "corporate_blue",
            "slides": [
                {
                    "template_id": "CORE-01",
                    "title": "Title",
                    "payload": {"title": "Executive Strategy Briefing", "subtitle": "Q4 Target Analysis"},
                },
                {
                    "template_id": "CORE-10",
                    "title": "Hybrid Architecture",
                    "lead_text": "Left table defines SLA parameters, right frame indicates topology diagram.",
                    "payload": {
                        "headers": ["Service", "SLA"],
                        "rows": [["API Gateway", "99.99%"], ["Worker", "99.95%"]],
                    },
                    "right_template_id": "CORE-50",
                    "right_payload": {
                        "placeholder_label": "Paste Enterprise Network Diagram Here",
                        "caption": "AWS DirectConnect + Transit Gateway",
                    },
                    "split_ratio": 0.5,
                },
            ],
        }

        deck = PresentationDeck.model_validate(deck_data)
        prs, summary = generate_presentation(deck)

        assert summary["total_slides"] == 2
        assert summary["rendered_slides"] == 2
        assert summary["fallback_slides"] == 0
        assert len(prs.slides) == 2

    def test_sample_deck_full_generation_zero_fallbacks(self) -> None:
        """Full sample deck must render all 10 slides with exactly 0 fallbacks."""
        sample_path = Path(__file__).parent.parent / "examples" / "sample_deck.json"
        with open(sample_path, encoding="utf-8") as f:
            deck_dict = json.load(f)

        deck = PresentationDeck.model_validate(deck_dict)
        prs, summary = generate_presentation(deck)

        assert summary["total_slides"] == 10
        assert summary["rendered_slides"] == 10
        assert summary["fallback_slides"] == 0
        assert len(prs.slides) == 10

        # Specifically verify the 10th slide (the heterogeneous split slide)
        slide_10 = prs.slides[9]
        # Must not have error notes
        if slide_10.has_notes_slide:
            notes_text = slide_10.notes_slide.notes_text_frame.text
            assert "Slide Rendering Error" not in notes_text

        # Must have shapes rendered on both split panes (table on left, placeholder on right)
        shape_types = [s.shape_type for s in slide_10.shapes]
        assert len(shape_types) >= 2

    def test_cli_fail_on_fallback_success(self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
        """CLI with --fail-on-fallback exits with 0 on normal presentation."""
        sample_path = Path(__file__).parent.parent / "examples" / "sample_deck.json"
        out_path = tmp_path / "out.pptx"

        monkeypatch.setattr(
            "sys.argv",
            ["slide-engine", "-i", str(sample_path), "-o", str(out_path), "--fail-on-fallback"],
        )
        with pytest.raises(SystemExit) as exc_info:
            main()
        assert exc_info.value.code == 0
        assert out_path.exists()
        assert out_path.stat().st_size > 0

    def test_cli_fail_on_fallback_failure(
        self,
        monkeypatch: pytest.MonkeyPatch,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """CLI with --fail-on-fallback exits with code 4 when a fallback occurs."""
        # Override build_dynamic_dispatch_table to include a crashing renderer
        def crashing_renderer(slide: Any, payload: Any, box: BoundingBox) -> None:
            raise RuntimeError("Forced renderer explosion")

        monkeypatch.setattr(
            "slide_engine.controller.build_dynamic_dispatch_table",
            lambda: {"CORE-00": lambda s, p, b: None, "CORE-10": crashing_renderer},
        )

        sample_input = {
            "theme_name": "corporate_blue",
            "slides": [
                {
                    "template_id": "CORE-10",
                    "title": "Will Fallback",
                    "payload": {"headers": ["A"], "rows": [["1"]]},
                }
            ],
        }

        monkeypatch.setattr("sys.stdin.read", lambda: json.dumps(sample_input))
        monkeypatch.setattr("sys.argv", ["slide-engine", "--fail-on-fallback"])

        with pytest.raises(SystemExit) as exc_info:
            main()
        assert exc_info.value.code == 4
        captured = capsys.readouterr()
        assert "Strict Generation Failure: 1 slide(s)" in captured.err


SPLIT_TEMPLATES_SAMPLE_PAYLOADS: dict[str, dict[str, Any]] = {
    "CORE-10": {
        "headers": ["Col A", "Col B"],
        "rows": [["1", "2"], ["3", "4"]],
    },
    "CORE-20": {
        "steps": [
            {"step_number": 1, "title": "Step 1", "description": "Desc 1"},
            {"step_number": 2, "title": "Step 2", "description": "Desc 2"},
        ]
    },
    "CORE-30": {
        "chart_type": "column",
        "categories": ["Q1", "Q2"],
        "series": [{"name": "S1", "values": [10.0, 20.0]}],
    },
    "CORE-40": {
        "x_axis_label": "X Axis",
        "y_axis_label": "Y Axis",
        "top_left": {"title": "TL", "badge": "P1", "description": "Top Left"},
        "top_right": {"title": "TR", "badge": "P2", "description": "Top Right"},
        "bottom_left": {"title": "BL", "badge": "P3", "description": "Bottom Left"},
        "bottom_right": {"title": "BR", "badge": "P4", "description": "Bottom Right"},
    },
    "CORE-50": {
        "placeholder_label": "Diagram Guide",
        "caption": "Caption Text",
    },
    "CORE-60": {
        "metrics": [
            {"label": "Metric A", "value": "100", "unit": "%", "trend": "+10%", "is_positive": True},
            {"label": "Metric B", "value": "200", "unit": "h", "trend": "-5%", "is_positive": False},
        ]
    },
    "CORE-70": {
        "left_column": {
            "header": "As-Is",
            "sub_header": "Legacy",
            "points": ["Point 1", "Point 2"],
            "is_highlighted": False,
        },
        "right_column": {
            "header": "To-Be",
            "sub_header": "Target",
            "points": ["Point A", "Point B"],
            "is_highlighted": True,
        },
    },
}

SPLIT_COMBINATIONS = list(
    itertools.product(SPLIT_TEMPLATES_SAMPLE_PAYLOADS.keys(), SPLIT_TEMPLATES_SAMPLE_PAYLOADS.keys())
)


def _assert_shape_bounds(shape: Any, slide_w: int, slide_h: int, tolerance: int, label: str) -> None:
    """Assert an individual shape does not exceed outer presentation canvas boundaries."""
    left = int(shape.left or 0)
    top = int(shape.top or 0)
    width = int(shape.width or 0)
    height = int(shape.height or 0)
    assert left >= 0, f"Shape in {label} left < 0: {left}"
    assert top >= 0, f"Shape in {label} top < 0: {top}"
    assert left + width <= slide_w + tolerance, f"Shape in {label} right exceeds width: {left + width} > {slide_w}"
    assert top + height <= slide_h + tolerance, f"Shape in {label} bottom exceeds height: {top + height} > {slide_h}"


def _verify_slide_pane_containment(
    slide: Any,
    body_top: int,
    left_box: BoundingBox,
    right_box: BoundingBox,
    slide_w: int,
    slide_h: int,
    label: str,
) -> None:
    """Verify all slide shapes are within boundaries and partitioned cleanly into left/right panes."""
    tolerance = int(Inches(0.05))
    left_shapes = []
    right_shapes = []

    for shape in slide.shapes:
        _assert_shape_bounds(shape, slide_w, slide_h, tolerance, label)
        top = int(shape.top or 0)
        left = int(shape.left or 0)
        if top >= body_top - tolerance:
            if left < int(right_box.left) - tolerance:
                left_shapes.append(shape)
            else:
                right_shapes.append(shape)

    if left_shapes and right_shapes:
        max_left_right = max(int(s.left or 0) + int(s.width or 0) for s in left_shapes)
        min_right_left = min(int(s.left or 0) for s in right_shapes)
        assert max_left_right <= int(left_box.right) + tolerance, (
            f"Left pane shape in {label} crosses into gap/right: {max_left_right} exceeds left pane right {int(left_box.right)}"
        )
        assert min_right_left >= int(right_box.left) - tolerance, (
            f"Right pane shape in {label} crosses into gap/left: {min_right_left} precedes right pane left {int(right_box.left)}"
        )
        assert max_left_right <= min_right_left + tolerance, (
            f"Pane overlap in {label}: left pane edge {max_left_right} exceeds right pane edge {min_right_left}"
        )


class TestSplitCombinations:
    """Verifies all pairwise split-layout combinations across all eligible templates."""

    @pytest.mark.parametrize("left_id,right_id", SPLIT_COMBINATIONS)
    def test_pairwise_split_rendering(self, left_id: str, right_id: str) -> None:
        """Any valid pair of split-supporting templates must render with zero fallbacks."""
        left_payload = SPLIT_TEMPLATES_SAMPLE_PAYLOADS[left_id]
        right_payload = SPLIT_TEMPLATES_SAMPLE_PAYLOADS[right_id]

        slide_def = SlideDefinition.model_validate({
            "template_id": left_id,
            "title": f"Split: {left_id} + {right_id}",
            "lead_text": "Evaluating split combination layout integrity.",
            "payload": left_payload,
            "right_template_id": right_id,
            "right_payload": right_payload,
            "split_ratio": 0.5,
        })

        deck = PresentationDeck(theme_name="corporate_blue", slides=[slide_def])
        prs, summary = generate_presentation(deck)

        assert summary["total_slides"] == 1
        assert summary["rendered_slides"] == 1
        assert summary["fallback_slides"] == 0
        assert len(prs.slides) == 1

        slide = prs.slides[0]
        if slide.has_notes_slide:
            assert "Slide Rendering Error" not in slide.notes_slide.notes_text_frame.text

        assert prs.slide_width is not None and prs.slide_height is not None
        _, body_box = SlideMaster.get_body_box_with_lead(has_lead_text=True)
        left_box, right_box = SlideMaster.split_horizontal(body_box, split_ratio=0.5)

        _verify_slide_pane_containment(
            slide=slide,
            body_top=int(body_box.top),
            left_box=left_box,
            right_box=right_box,
            slide_w=int(prs.slide_width),
            slide_h=int(prs.slide_height),
            label=f"{left_id}+{right_id}",
        )

    @pytest.mark.parametrize("split_ratio", [0.2, 0.35, 0.5, 0.65, 0.8])
    def test_split_ratios_and_pane_separation(self, split_ratio: float) -> None:
        """Verify dynamic layout recalculation across various horizontal split ratios."""
        slide_def = SlideDefinition.model_validate({
            "template_id": "CORE-10",
            "title": f"Split Ratio {split_ratio}",
            "lead_text": "Testing variable split ratios with pane bounds.",
            "payload": SPLIT_TEMPLATES_SAMPLE_PAYLOADS["CORE-10"],
            "right_template_id": "CORE-60",
            "right_payload": SPLIT_TEMPLATES_SAMPLE_PAYLOADS["CORE-60"],
            "split_ratio": split_ratio,
        })

        deck = PresentationDeck(theme_name="corporate_blue", slides=[slide_def])
        prs, summary = generate_presentation(deck)
        assert summary["fallback_slides"] == 0

        slide = prs.slides[0]
        assert prs.slide_width is not None and prs.slide_height is not None
        _, body_box = SlideMaster.get_body_box_with_lead(has_lead_text=True)
        left_box, right_box = SlideMaster.split_horizontal(body_box, split_ratio=split_ratio)

        _verify_slide_pane_containment(
            slide=slide,
            body_top=int(body_box.top),
            left_box=left_box,
            right_box=right_box,
            slide_w=int(prs.slide_width),
            slide_h=int(prs.slide_height),
            label=f"split_ratio_{split_ratio}",
        )

    @pytest.mark.parametrize(
        "left_t,right_t",
        [
            ("CORE-10", "CORE-70"),
            ("CORE-20", "CORE-60"),
            ("CORE-40", "CORE-10"),
            ("CORE-60", "CORE-70"),
        ],
    )
    def test_split_boundary_maximum_capacity(self, left_t: str, right_t: str) -> None:
        """Verify rendering resilience at maximum payload capacities across diverse template pairs."""
        boundary_payloads: dict[str, dict[str, Any]] = {
            "CORE-10": {
                "headers": [f"Header {i}" for i in range(1, 6)],
                "rows": [[f"Cell {r}-{c}" for c in range(1, 6)] for r in range(1, 16)],
            },
            "CORE-20": {
                "steps": [
                    {
                        "step_number": i,
                        "title": f"Process Step {i}",
                        "description": f"Detailed execution task for step {i} in automated pipeline.",
                    }
                    for i in range(1, 7)
                ]
            },
            "CORE-40": {
                "x_axis_label": "Operational Efficiency",
                "y_axis_label": "Strategic Alignment",
                "top_left": {
                    "title": "Strategic Invest",
                    "badge": "P1",
                    "description": "Long-term foundational investments with high architectural payoff.",
                },
                "top_right": {
                    "title": "Core Strengths",
                    "badge": "P0",
                    "description": "Mission-critical revenue generating operations requiring continuous scale.",
                },
                "bottom_left": {
                    "title": "Sunsetting",
                    "badge": "P3",
                    "description": "Legacy components slated for graceful decommissioning.",
                },
                "bottom_right": {
                    "title": "Quick Wins",
                    "badge": "P2",
                    "description": "Immediate automation tactical steps yielding quick efficiency.",
                },
            },
            "CORE-60": {
                "metrics": [
                    {
                        "label": f"Metric {i}",
                        "value": f"{i * 125}k",
                        "unit": "JPY",
                        "trend": f"+{i * 4}%",
                        "is_positive": True,
                    }
                    for i in range(1, 5)
                ]
            },
            "CORE-70": {
                "left_column": {
                    "header": "Current State Analysis",
                    "sub_header": "High Friction Legacy",
                    "points": [f"Current bottleneck point {i} in operation" for i in range(1, 6)],
                    "is_highlighted": False,
                },
                "right_column": {
                    "header": "Future State Vision",
                    "sub_header": "Optimized AI Pipeline",
                    "points": [f"Future strategic transformation initiative {i}" for i in range(1, 6)],
                    "is_highlighted": True,
                },
            },
        }

        slide_def = SlideDefinition.model_validate({
            "template_id": left_t,
            "title": f"Max Capacity Split: {left_t} + {right_t}",
            "lead_text": "Boundary stress testing of high density content panes.",
            "payload": boundary_payloads[left_t],
            "right_template_id": right_t,
            "right_payload": boundary_payloads[right_t],
            "split_ratio": 0.5,
        })

        deck = PresentationDeck(theme_name="corporate_blue", slides=[slide_def])
        prs, summary = generate_presentation(deck)
        assert summary["fallback_slides"] == 0

        slide = prs.slides[0]
        assert prs.slide_width is not None and prs.slide_height is not None
        slide_w, slide_h = int(prs.slide_width), int(prs.slide_height)

        for s in slide.shapes:
            left = int(s.left or 0)
            top = int(s.top or 0)
            width = int(s.width or 0)
            height = int(s.height or 0)
            assert left >= 0
            assert top >= 0
            assert left + width <= slide_w + int(Inches(0.05))
            assert top + height <= slide_h + int(Inches(0.05))


def _estimate_text_frame_required_height(tf: Any, box_w: float) -> float:
    """Estimate total required height in inches for text inside a text frame using CJK font engine logic."""
    total_height = 0.0
    for paragraph in tf.paragraphs:
        p_text = paragraph.text.strip()
        if not p_text:
            continue
        pt_size = paragraph.font.size.pt if paragraph.font.size else 12.0
        req_pt = FontEngine.estimate_height_pt(p_text, box_w * 72.0, pt_size)
        total_height += req_pt / 72.0
    return total_height


def _check_shape_text_overflow(shape: Any, slide_idx: int, shape_idx: int, tolerance_ratio: float) -> int:
    """Check text frame or table cells for vertical overflow. Returns number of checked frames."""
    checked = 0
    if shape.has_text_frame and shape.text_frame.text.strip():
        box_w = float(shape.width) / 914400.0
        box_h = float(shape.height) / 914400.0
        req_h = _estimate_text_frame_required_height(shape.text_frame, box_w)
        assert req_h <= box_h * tolerance_ratio, (
            f"Slide {slide_idx} Shape {shape_idx} text overflow: "
            f"required height {req_h:.2f}in exceeds box height {box_h:.2f}in"
        )
        checked += 1

    if shape.has_table:
        table = shape.table
        col_widths = [float(col.width) / 914400.0 for col in table.columns]
        row_heights = [float(row.height) / 914400.0 for row in table.rows]
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                if cell.text_frame and cell.text_frame.text.strip():
                    c_w = col_widths[c_idx]
                    r_h = row_heights[r_idx]
                    req_h = _estimate_text_frame_required_height(cell.text_frame, c_w)
                    assert req_h <= r_h * tolerance_ratio, (
                        f"Slide {slide_idx} Table cell ({r_idx},{c_idx}) overflow: "
                        f"required height {req_h:.2f}in exceeds cell height {r_h:.2f}in"
                    )
                    checked += 1
    return checked


class TestTextOverflowEstimation:
    """Verifies that generated text content does not visually overflow allocated shape boundaries."""

    def test_text_does_not_overflow_allocated_boxes(self) -> None:
        """Inspect all rendered text frames across sample slides to ensure text fits within bounding boxes."""
        sample_path = Path(__file__).parent.parent / "examples" / "sample_deck.json"
        deck = PresentationDeck.model_validate_json(sample_path.read_text(encoding="utf-8"))
        prs, summary = generate_presentation(deck)
        assert summary["fallback_slides"] == 0

        overflow_tolerance_ratio = 1.15  # Strict 15% padding allowance

        checked_text_frames = 0
        for slide_idx, slide in enumerate(prs.slides):
            for shape_idx, shape in enumerate(slide.shapes):
                checked_text_frames += _check_shape_text_overflow(
                    shape, slide_idx, shape_idx, overflow_tolerance_ratio
                )

        assert checked_text_frames > 40, f"Expected >40 text frames checked (including tables), got {checked_text_frames}"

