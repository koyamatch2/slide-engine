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
from slide_engine.core_base import BoundingBox, ShapeBuilder, SlideMaster
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
