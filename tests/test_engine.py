"""Comprehensive test suite for Slide Engine.

Verifies:
1. Dynamic dispatch resolution and function signatures across all registered templates.
2. Pydantic boundary validation (text limits, array lengths, chart series matching, split density ceilings).
3. Complex layouts (lead text vertical offset, heterogeneous split geometry).
4. Fail-soft behavior (CORE-00 fallback on exceptions).
5. Clean canvas purging and master background preservation.
6. Direct unit verification of CJK-aware FontEngine scaling.
7. Adaptive grid layout, strict physical non-overflow, and zero peer collision under extreme split and max-capacity stress.
"""

from __future__ import annotations

import inspect
import json
import math
import unicodedata
from pathlib import Path
from typing import Any

import pytest
from pptx import Presentation
from pptx.util import Inches, Length
from pydantic import ValidationError

from slide_engine.__main__ import main
from slide_engine.controller import (
    _initialize_presentation,
    _render_slide_with_failsoft,
    _resolve_safe_asset_path,
    build_dynamic_dispatch_table,
    generate_presentation,
)
from slide_engine.core_base import BoundingBox, FontEngine, ShapeBuilder, SlideMaster
from slide_engine.models import (
    MAX_SLIDES_PER_DECK,
    TEMPLATE_REGISTRY,
    PresentationDeck,
    SlideDefinition,
)


class TestDynamicDispatch:
    """Verifies template registry and dispatch table integrity."""

    def test_all_registered_templates_have_renderers(self) -> None:
        """Bijective 1-to-1 matching across filesystem, TEMPLATE_REGISTRY, and dispatch_table."""
        dispatch_table = build_dynamic_dispatch_table()
        pkg_dir = Path(__file__).parent.parent / "src" / "slide_engine"
        expected_tids = {
            f"CORE-{p.stem.replace('render_core_', '').upper()}"
            for p in pkg_dir.glob("render_core_*.py")
        }
        assert set(TEMPLATE_REGISTRY.keys()) == expected_tids, "Mismatch between files and TEMPLATE_REGISTRY"
        assert set(dispatch_table.keys()) == expected_tids, "Mismatch between files and dispatch_table"

    def test_exactly_one_fallback_registered(self) -> None:
        """System must register exactly one fallback template."""
        fallbacks = [tid for tid, s in TEMPLATE_REGISTRY.items() if s.is_fallback]
        assert len(fallbacks) == 1, f"Expected exactly 1 fallback template, found: {fallbacks}"

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


def _get_field_max_length(field_info: Any) -> int | None:
    """Extract max_length constraint from field metadata if present."""
    for meta in field_info.metadata:
        max_len = getattr(meta, "max_length", None)
        if max_len is not None:
            return int(max_len)
    return None


class TestSchemaValidationAndDefenses:
    """Tests Pydantic V2 models, resource limits, and error rejection."""

    def test_template_custom_invalid_payloads_rejected(self) -> None:
        """Every non-fallback template must define invalid payloads and reject them via Pydantic validation."""
        for tid, spec in TEMPLATE_REGISTRY.items():
            if spec.is_fallback:
                continue
            assert len(spec.invalid_sample_payloads) >= 1, (
                f"Template '{tid}' must define at least one invalid_sample_payload to verify rejection defenses."
            )
            for invalid_payload in spec.invalid_sample_payloads:
                with pytest.raises(ValueError):
                    spec.payload_type.model_validate(invalid_payload)

    def test_deck_slide_count_limit(self) -> None:
        """Deck must reject slides exceeding MAX_SLIDES_PER_DECK."""
        valid_tid, valid_spec = next((t, s) for t, s in TEMPLATE_REGISTRY.items() if not s.is_fallback)
        valid_slide_dict = {
            "template_id": valid_tid,
            "title": "Title",
            "payload": valid_spec.sample_payload,
        }
        excessive_slides = [valid_slide_dict] * (MAX_SLIDES_PER_DECK + 1)
        with pytest.raises(ValueError):
            PresentationDeck.model_validate({"theme_name": "corporate_blue", "slides": excessive_slides})

    def test_theme_name_path_traversal_rejection(self) -> None:
        """PresentationDeck must reject path traversal sequences and invalid theme_name characters."""
        valid_tid, valid_spec = next((t, s) for t, s in TEMPLATE_REGISTRY.items() if not s.is_fallback)
        slide_dict = {
            "template_id": valid_tid,
            "title": "Title",
            "payload": valid_spec.sample_payload,
        }

        invalid_themes = [
            "../../../../etc/passwd",
            "/app/assets/malicious",
            "../corporate_blue",
            "corporate/blue",
            "theme\\sub",
            "theme name with space",
            "theme$special!",
            "",
            "a" * 81,
        ]

        for inv_theme in invalid_themes:
            with pytest.raises(ValidationError):
                PresentationDeck.model_validate({"theme_name": inv_theme, "slides": [slide_dict]})

    def test_resolve_safe_asset_path_containment(self) -> None:
        """Controller _resolve_safe_asset_path must never resolve paths outside allowed asset dirs."""
        assert _resolve_safe_asset_path("../../../../etc/passwd") is None
        assert _resolve_safe_asset_path("..\\something") is None
        assert _resolve_safe_asset_path("non_existent_random_theme_12345") is None

    def test_unsupported_split_template_raises_error(self) -> None:
        """Templates with supports_split=False must reject horizontal split."""
        unsupported = [tid for tid, s in TEMPLATE_REGISTRY.items() if not s.supports_split]
        split_capable = [tid for tid, s in TEMPLATE_REGISTRY.items() if s.supports_split]
        if not unsupported or not split_capable:
            pytest.skip("Requires both split-capable and non-split templates registered.")

        target_tid = unsupported[0]
        partner_tid = split_capable[0]
        with pytest.raises(ValueError, match="does not support horizontal split"):
            SlideDefinition.model_validate({
                "template_id": target_tid,
                "title": "Invalid Split",
                "payload": TEMPLATE_REGISTRY[target_tid].sample_payload,
                "right_template_id": partner_tid,
                "right_payload": TEMPLATE_REGISTRY[partner_tid].effective_split_payload,
            })

    def test_payload_purity_rejects_extra_fields(self) -> None:
        """All registered template payloads must reject forbidden extra attributes."""
        for _tid, spec in TEMPLATE_REGISTRY.items():
            sample_with_extra = dict(spec.sample_payload)
            sample_with_extra["__forbidden_injected_extra_attribute__"] = "malicious_payload"
            with pytest.raises(ValidationError, match="Extra inputs are not permitted"):
                spec.payload_type.model_validate(sample_with_extra)

    def test_template_registry_sample_payloads_valid_against_schema(self) -> None:
        """Every registered template's sample payloads must pass validation and enforce boundary invariants."""
        split_capable = [tid for tid, s in TEMPLATE_REGISTRY.items() if s.supports_split]

        for tid, spec in TEMPLATE_REGISTRY.items():
            assert spec.payload_type.model_validate(spec.sample_payload) is not None
            if spec.split_sample_payload is None:
                continue

            assert spec.payload_type.model_validate(spec.split_sample_payload) is not None

            # Automated Zero-Maintenance Test:
            # Any template with a distinct split_sample_payload MUST reject its full sample_payload
            # inside a split layout via validate_split_density.
            partner_candidates = [p for p in split_capable if p != tid]
            if partner_candidates:
                partner_tid = partner_candidates[0]
                with pytest.raises(ValidationError):
                    SlideDefinition.model_validate({
                        "template_id": tid,
                        "title": f"Split Density Overload: {tid}",
                        "payload": spec.sample_payload,
                        "right_template_id": partner_tid,
                        "right_payload": TEMPLATE_REGISTRY[partner_tid].effective_split_payload,
                    })

    def test_template_registry_sample_payloads_saturate_max_capacity(self) -> None:
        """All top-level collections in sample_payloads must fully saturate schema max_length limits."""
        for tid, spec in TEMPLATE_REGISTRY.items():
            for fname, field_info in spec.payload_type.model_fields.items():
                max_len = _get_field_max_length(field_info)
                field_val = spec.sample_payload.get(fname)
                if max_len is None or not isinstance(field_val, list):
                    continue
                assert len(field_val) == max_len, (
                    f"Template '{tid}' sample_payload field '{fname}' length {len(field_val)} "
                    f"does not saturate schema max_length={max_len}. Max capacity stress test requires full saturation."
                )

    def test_split_ratio_bounds_raise_errors(self) -> None:
        """Split layout must reject split_ratio below 0.25 or above 0.75."""
        split_capable = [tid for tid, s in TEMPLATE_REGISTRY.items() if s.supports_split]
        if len(split_capable) < 2:
            pytest.skip("Requires at least 2 split-capable templates registered.")

        t_left, t_right = split_capable[0], split_capable[1]
        p_left = TEMPLATE_REGISTRY[t_left].effective_split_payload
        p_right = TEMPLATE_REGISTRY[t_right].effective_split_payload

        for invalid_ratio, err_match in [
            (0.20, "greater than or equal to 0.25"),
            (0.80, "less than or equal to 0.75"),
        ]:
            with pytest.raises(ValidationError, match=err_match):
                SlideDefinition.model_validate({
                    "template_id": t_left,
                    "title": "Split Ratio Out of Bounds",
                    "payload": p_left,
                    "right_template_id": t_right,
                    "right_payload": p_right,
                    "split_ratio": invalid_ratio,
                })

    def test_core10_row_column_count_mismatch_rejected(self) -> None:
        """CORE-10 table must reject row data with column count mismatching headers."""
        spec = TEMPLATE_REGISTRY["CORE-10"]
        with pytest.raises(ValidationError, match="does not match headers count"):
            spec.payload_type.model_validate({
                "headers": ["Col 1", "Col 2"],
                "rows": [["Val 1"]],
            })

    def test_core10_highlight_index_out_of_bounds_rejected(self) -> None:
        """CORE-10 table must reject highlight_row_index exceeding row count."""
        spec = TEMPLATE_REGISTRY["CORE-10"]
        with pytest.raises(ValidationError, match="exceeds total row count"):
            spec.payload_type.model_validate({
                "headers": ["Col 1"],
                "rows": [["Val 1"]],
                "highlight_row_index": 5,
            })

    def test_core30_split_density_rejected(self) -> None:
        """CORE-30 chart in split pane must reject exceeding 8 categories or 3 series."""
        chart_spec = TEMPLATE_REGISTRY["CORE-30"]
        partner_tid = next(t for t, s in TEMPLATE_REGISTRY.items() if s.supports_split and t != "CORE-30")
        partner_spec = TEMPLATE_REGISTRY[partner_tid]

        with pytest.raises(ValidationError, match="exceeds legibility limit"):
            SlideDefinition.model_validate({
                "template_id": "CORE-30",
                "title": "Overloaded Split Chart",
                "payload": chart_spec.sample_payload,
                "right_template_id": partner_tid,
                "right_payload": partner_spec.effective_split_payload,
            })

    def test_split_partial_fields_rejected(self) -> None:
        """Providing only one of right_template_id or right_payload must raise ValidationError."""
        spec = TEMPLATE_REGISTRY["CORE-10"]
        with pytest.raises(ValidationError, match="must be provided together"):
            SlideDefinition.model_validate({
                "template_id": "CORE-10",
                "title": "Partial Split Missing Payload",
                "payload": spec.effective_split_payload,
                "right_template_id": "CORE-10",
            })
        with pytest.raises(ValidationError, match="must be provided together"):
            SlideDefinition.model_validate({
                "template_id": "CORE-10",
                "title": "Partial Split Missing ID",
                "payload": spec.effective_split_payload,
                "right_payload": spec.effective_split_payload,
            })


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
        assert int(box.right) == 4000
        assert int(box.bottom) == 6000

    def test_split_horizontal_exact_geometry_and_conservation(self) -> None:
        """split_horizontal must strictly conserve total width and retain Length types."""
        content_box = SlideMaster.CONTENT_BOX
        left_box, right_box = SlideMaster.split_horizontal(content_box, split_ratio=0.5)

        assert isinstance(left_box.left, Length)
        assert isinstance(left_box.width, Length)
        assert isinstance(left_box.right, Length)
        assert isinstance(right_box.left, Length)
        assert isinstance(right_box.width, Length)
        assert isinstance(right_box.right, Length)

        assert left_box.left == content_box.left
        assert left_box.top == content_box.top
        assert right_box.top == content_box.top
        assert left_box.height == content_box.height
        assert right_box.height == content_box.height

        gap = SlideMaster.DEFAULT_SPLIT_GAP
        assert abs(int(left_box.width + right_box.width + gap) - int(content_box.width)) <= 1
        assert right_box.left == Length(left_box.left + left_box.width + gap)

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

        slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(2))
        slide.shapes.add_table(2, 2, Inches(3), Inches(3), Inches(2), Inches(2))
        assert len(slide.shapes) == 2

        ShapeBuilder.clean_slide_canvas(slide)
        assert len(slide.shapes) == 0


class TestFailSoftExecution:
    """Verifies that rendering errors fall back to fallback template safely."""

    def test_exception_triggers_fallback(self) -> None:
        """A crashing renderer must be replaced by fallback template without halting execution."""
        prs = _initialize_presentation("corporate_blue")
        dispatch_table = build_dynamic_dispatch_table()

        def broken_renderer(slide: Any, payload: Any, box: BoundingBox) -> None:
            raise RuntimeError("Simulated Database Render Failure")

        any_tid, any_spec = next((t, s) for t, s in TEMPLATE_REGISTRY.items() if not s.is_fallback)
        broken_table = dict(dispatch_table)
        broken_table[any_tid] = broken_renderer

        slide_def = SlideDefinition.model_validate({
            "template_id": any_tid,
            "title": "Faulty Slide",
            "payload": any_spec.sample_payload,
        })

        success = _render_slide_with_failsoft(prs, slide_def, broken_table)
        assert success is False
        assert len(prs.slides) == 1

        notes = prs.slides[0].notes_slide.notes_text_frame.text
        assert "Simulated Database Render Failure" in notes

    def test_cli_fail_on_fallback_success(
        self,
        monkeypatch: pytest.MonkeyPatch,
        capsys: pytest.CaptureFixture[str],
        tmp_path: Path,
    ) -> None:
        """CLI with --fail-on-fallback exits with 0 and emits telemetry JSON to STDERR."""
        active_tid, active_spec = next((t, s) for t, s in TEMPLATE_REGISTRY.items() if not s.is_fallback)
        valid_deck_dict = {
            "theme_name": "corporate_blue",
            "slides": [
                {
                    "template_id": active_tid,
                    "title": "CLI Test Slide",
                    "payload": active_spec.sample_payload,
                }
            ],
        }
        test_deck_path = tmp_path / "valid_deck.json"
        test_deck_path.write_text(json.dumps(valid_deck_dict), encoding="utf-8")
        out_path = tmp_path / "out.pptx"

        monkeypatch.setattr(
            "sys.argv",
            ["slide-engine", "-i", str(test_deck_path), "-o", str(out_path), "--fail-on-fallback"],
        )
        with pytest.raises(SystemExit) as exc_info:
            main()
        assert exc_info.value.code == 0
        assert out_path.exists()
        assert out_path.stat().st_size > 0

        captured = capsys.readouterr()
        assert '"event": "deck_generation_summary"' in captured.err
        assert '"rendered_slides": 1' in captured.err

    def test_generate_presentation_is_pure_without_stderr_side_effect(
        self,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """Calling generate_presentation directly must not write to STDERR (pure library behavior)."""
        active_tid, active_spec = next((t, s) for t, s in TEMPLATE_REGISTRY.items() if not s.is_fallback)
        deck = PresentationDeck.model_validate({
            "theme_name": "corporate_blue",
            "slides": [
                {
                    "template_id": active_tid,
                    "title": "Pure Lib Call",
                    "payload": active_spec.sample_payload,
                }
            ],
        })
        prs, summary = generate_presentation(deck)
        assert summary["rendered_slides"] == 1
        assert len(prs.slides) == 1

        captured = capsys.readouterr()
        assert captured.err == "", f"generate_presentation polluted STDERR with: {captured.err!r}"

    def test_cli_fail_on_fallback_failure(
        self,
        monkeypatch: pytest.MonkeyPatch,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """CLI with --fail-on-fallback exits with code 4 when a fallback occurs."""
        def crashing_renderer(slide: Any, payload: Any, box: BoundingBox) -> None:
            raise RuntimeError("Forced renderer explosion")

        any_tid, any_spec = next((t, s) for t, s in TEMPLATE_REGISTRY.items() if not s.is_fallback)
        fallback_tid = next(t for t, s in TEMPLATE_REGISTRY.items() if s.is_fallback)
        monkeypatch.setattr(
            "slide_engine.controller.build_dynamic_dispatch_table",
            lambda: {fallback_tid: lambda s, p, b: None, any_tid: crashing_renderer},
        )

        sample_input = {
            "theme_name": "corporate_blue",
            "slides": [
                {
                    "template_id": any_tid,
                    "title": "Will Fallback",
                    "payload": any_spec.sample_payload,
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


class TestFontEngineUnit:
    """Direct unit tests for FontEngine without self-referential helpers."""

    def test_cjk_fullwidth_scales_smaller_than_ascii_for_same_char_count(self) -> None:
        """Japanese fullwidth characters must trigger font reduction earlier than ASCII of equal length."""
        box_w = Length(Inches(2.0))
        box_h = Length(Inches(0.60))

        ascii_text = "A" * 32
        cjk_text = "あ" * 32

        ascii_font = FontEngine.calculate_font_size(ascii_text, box_w, box_h, max_pt=14.0, min_pt=8.0)
        cjk_font = FontEngine.calculate_font_size(cjk_text, box_w, box_h, max_pt=14.0, min_pt=8.0)

        assert ascii_font.pt == 14.0
        assert cjk_font.pt <= 11.5
        assert cjk_font.pt >= 8.0

    def test_font_size_respects_max_and_min_bounds(self) -> None:
        """calculate_font_size must return max_pt for empty/short text and floor at min_pt for massive text."""
        box_w = Length(Inches(3.0))
        box_h = Length(Inches(0.5))

        assert FontEngine.calculate_font_size("", box_w, box_h, max_pt=16.0, min_pt=9.0).pt == 16.0
        assert FontEngine.calculate_font_size("Short", box_w, box_h, max_pt=16.0, min_pt=9.0).pt == 16.0

        huge_text = "大規模な基幹システム刷新における段階的移行アプローチとリスク緩和策の検討" * 10
        assert FontEngine.calculate_font_size(huge_text, box_w, box_h, max_pt=16.0, min_pt=9.0).pt == 9.0


# ----------------------------------------------------------------------
# Independent Physical Layout & Invariant Verification Helpers
# ----------------------------------------------------------------------


def _independent_paragraph_height_pt(paragraph: Any, avail_width_pt: float) -> float:
    """Independently compute physical height of a paragraph without calling FontEngine."""
    p_text = paragraph.text.strip()
    if not p_text:
        return 0.0

    pt_candidates = [paragraph.font.size.pt] if paragraph.font.size else []
    for run in paragraph.runs:
        if run.font and run.font.size:
            pt_candidates.append(run.font.size.pt)
    font_pt = max(pt_candidates) if pt_candidates else 12.0

    assert font_pt >= 8.0, f"Illegible micro-font detected ({font_pt}pt < 8.0pt) in text: '{p_text[:30]}'"

    em_width = sum(1.0 if unicodedata.east_asian_width(ch) in ("F", "W", "A") else 0.58 for ch in p_text)
    chars_per_line = max(1.0, avail_width_pt / font_pt)
    lines = max(1, math.ceil(em_width / chars_per_line))

    space_after_pt = paragraph.space_after.pt if paragraph.space_after else 0.0
    return (lines * font_pt * 1.3) + space_after_pt


def _assert_text_frame_fits_physically(
    tf: Any,
    box_w_pt: float,
    box_h_pt: float,
    label: str,
) -> int:
    """Verify text frame content fits physically inside (box_w_pt, box_h_pt) without artificial padding."""
    if not tf or not tf.text.strip():
        return 0

    assert box_w_pt >= 24.0, f"{label}: text frame width collapsed ({box_w_pt:.1f}pt < 24pt)"
    assert box_h_pt >= 12.0, f"{label}: text frame height collapsed ({box_h_pt:.1f}pt < 12pt)"

    total_req_pt = sum(_independent_paragraph_height_pt(p, box_w_pt) for p in tf.paragraphs)
    # Strict physical containment: text height must never exceed available box height (+0.5pt float epsilon)
    assert total_req_pt <= box_h_pt + 0.5, (
        f"{label} physical text overflow: required {total_req_pt:.1f}pt > available {box_h_pt:.1f}pt "
        f"(text={tf.text.strip()[:40]!r})"
    )
    return 1


def _assert_table_cells_fit_physically(table: Any, label: str) -> int:
    """Verify all cells within a table meet text containment limits with zero deep nesting."""
    col_w_pts = [col.width / 12700.0 for col in table.columns]
    row_h_pts = [row.height / 12700.0 for row in table.rows]
    checked = 0
    for r_idx, row in enumerate(table.rows):
        for c_idx, cell in enumerate(row.cells):
            avail_w_pt = col_w_pts[c_idx] - 6.0
            avail_h_pt = row_h_pts[r_idx] - 3.0
            checked += _assert_text_frame_fits_physically(
                cell.text_frame,
                avail_w_pt,
                avail_h_pt,
                f"{label} cell({r_idx},{c_idx})",
            )
    return checked


def _assert_shape_physical_invariants(shape: Any, slide_w: int, slide_h: int, label: str) -> int:
    """Verify a shape has positive non-collapsed dimensions, stays in bounds, and has zero text overflow."""
    left = int(shape.left or 0)
    top = int(shape.top or 0)
    width = int(shape.width or 0)
    height = int(shape.height or 0)
    tolerance = int(Inches(0.02))

    # All shapes must have positive non-zero dimensions and stay inside the slide canvas
    assert width > 0, f"{label} width must be positive: {width}"
    assert height > 0, f"{label} height must be positive: {height}"
    assert left >= 0 and top >= 0, f"{label} negative coordinates: ({left}, {top})"
    assert left + width <= slide_w + tolerance, f"{label} exceeds slide width"
    assert top + height <= slide_h + tolerance, f"{label} exceeds slide height"

    # Content-bearing shapes (text, tables, charts) must meet minimum legible bounds
    if _is_content_shape(shape):
        assert width >= int(Inches(0.35)), f"{label} content width collapsed: {width / 914400:.2f}in"
        assert height >= int(Inches(0.20)), f"{label} content height collapsed: {height / 914400:.2f}in"

    checked = 0
    if shape.has_text_frame:
        checked += _assert_text_frame_fits_physically(
            shape.text_frame,
            width / 12700.0,
            height / 12700.0,
            label,
        )

    if shape.has_table:
        checked += _assert_table_cells_fit_physically(shape.table, label)

    return checked


def _assert_pane_containment_and_separation(
    slide: Any,
    body_top: int,
    left_box: BoundingBox,
    right_box: BoundingBox,
    label: str,
) -> None:
    """Independently verify left pane and right pane shapes stay strictly inside their allocated boxes."""
    tolerance = int(Inches(0.02))
    left_shapes = [
        s for s in slide.shapes
        if int(s.top or 0) >= body_top - tolerance and int(s.left or 0) < int(right_box.left) - tolerance
    ]
    right_shapes = [
        s for s in slide.shapes
        if int(s.top or 0) >= body_top - tolerance and int(s.left or 0) >= int(right_box.left) - tolerance
    ]

    assert len(left_shapes) > 0, f"{label}: left pane rendered 0 shapes"
    assert len(right_shapes) > 0, f"{label}: right pane rendered 0 shapes"

    for s in left_shapes:
        s_left, s_right = int(s.left or 0), int(s.left or 0) + int(s.width or 0)
        assert s_left >= int(left_box.left) - tolerance, f"{label}: left shape leaks left"
        assert s_right <= int(left_box.right) + tolerance, f"{label}: left shape leaks into gap/right pane"

    for s in right_shapes:
        s_left, s_right = int(s.left or 0), int(s.left or 0) + int(s.width or 0)
        assert s_left >= int(right_box.left) - tolerance, f"{label}: right shape leaks into gap/left pane"
        assert s_right <= int(right_box.right) + tolerance, f"{label}: right shape leaks right"


def _is_content_shape(shape: Any) -> bool:
    """Determine whether a shape is a non-empty text, table, or chart shape."""
    if shape.has_table or shape.has_chart:
        return True
    return bool(shape.has_text_frame and shape.text_frame.text.strip())


def _check_shape_pair_overlap(s1: Any, s2: Any, tolerance: int, label: str) -> None:
    """Assert two distinct shapes do not physically collide."""
    l1, t1, w1, h1 = int(s1.left or 0), int(s1.top or 0), int(s1.width or 0), int(s1.height or 0)
    l2, t2, w2, h2 = int(s2.left or 0), int(s2.top or 0), int(s2.width or 0), int(s2.height or 0)

    overlap_w = min(l1 + w1, l2 + w2) - max(l1, l2)
    overlap_h = min(t1 + h1, t2 + h2) - max(t1, t2)

    if overlap_w > tolerance and overlap_h > tolerance:
        text1 = s1.text_frame.text[:30].replace("\n", " ") if s1.has_text_frame else f"<type {s1.shape_type}>"
        text2 = s2.text_frame.text[:30].replace("\n", " ") if s2.has_text_frame else f"<type {s2.shape_type}>"
        raise AssertionError(
            f"{label}: Content shapes collide by {overlap_w / 914400:.3f}in x {overlap_h / 914400:.3f}in.\n"
            f"  Shape A: ({l1/914400:.2f}, {t1/914400:.2f}, {w1/914400:.2f}, {h1/914400:.2f}) text={text1!r}\n"
            f"  Shape B: ({l2/914400:.2f}, {t2/914400:.2f}, {w2/914400:.2f}, {h2/914400:.2f}) text={text2!r}"
        )


def _assert_no_peer_text_overlap(slide: Any, label: str) -> None:
    """Independently verify that no two content-bearing shapes (text, tables, charts) collide/overlap."""
    content_shapes = [s for s in slide.shapes if _is_content_shape(s)]
    tolerance = int(Inches(0.01))
    for i in range(len(content_shapes)):
        for j in range(i + 1, len(content_shapes)):
            _check_shape_pair_overlap(content_shapes[i], content_shapes[j], tolerance, label)


def _assert_deck_physical_integrity(
    prs: Any,
    expected_slides: int,
    deck_label: str,
) -> None:
    """Verify physical containment, lack of collision, and text legibility across all slides in a deck."""
    assert prs.slide_width is not None and prs.slide_height is not None
    slide_w = int(prs.slide_width)
    slide_h = int(prs.slide_height)

    total_shapes_checked = 0
    total_text_elements_checked = 0
    total_shapes_in_deck = sum(len(slide.shapes) for slide in prs.slides)

    for s_idx, slide in enumerate(prs.slides):
        _assert_no_peer_text_overlap(slide, f"{deck_label}[{s_idx}]")
        for sh_idx, shape in enumerate(slide.shapes):
            total_shapes_checked += 1
            total_text_elements_checked += _assert_shape_physical_invariants(
                shape,
                slide_w,
                slide_h,
                f"{deck_label}[{s_idx}].Shape[{sh_idx}]",
            )

    assert total_shapes_checked == total_shapes_in_deck, "All shapes in all slides must be verified"
    assert total_shapes_checked >= expected_slides * 2, f"Expected average >= 2 shapes per slide, got {total_shapes_checked}"
    assert total_text_elements_checked >= expected_slides, f"Expected text elements >= slide count, got {total_text_elements_checked}"


class TestAdaptiveLayoutAndPhysicalInvariants:
    """Verifies physical layout integrity, card non-collapse, zero collision, and text fit under stress."""

    def test_sample_deck_comprehensive_physical_integrity(self) -> None:
        """All slides in sample_deck.json must pass zero-fallback, boundary, text-fit, and overlap checks."""
        sample_path = Path(__file__).parent.parent / "examples" / "sample_deck.json"
        raw_deck = json.loads(sample_path.read_text(encoding="utf-8"))
        registered_templates = set(TEMPLATE_REGISTRY.keys())
        active_slides = [
            s for s in raw_deck.get("slides", [])
            if s.get("template_id") in registered_templates
            and (not s.get("right_template_id") or s.get("right_template_id") in registered_templates)
        ]
        if not active_slides:
            pytest.skip("No slides in sample_deck.json match currently registered templates.")

        deck = PresentationDeck.model_validate({
            "theme_name": raw_deck.get("theme_name", "corporate_blue"),
            "slides": active_slides,
        })

        prs, summary = generate_presentation(deck)
        num_slides = len(deck.slides)
        assert summary["total_slides"] == num_slides
        assert summary["rendered_slides"] == num_slides
        assert summary["fallback_slides"] == 0
        assert len(prs.slides) == num_slides

        _assert_deck_physical_integrity(prs, num_slides, "SampleSlide")

        # Verify any split slides present in the deck render without error notes
        for slide_def, slide in zip(deck.slides, prs.slides, strict=True):
            if slide_def.right_template_id and slide.has_notes_slide:
                assert "Slide Rendering Error" not in slide.notes_slide.notes_text_frame.text

    def test_all_registered_templates_synthesized_deck_integrity(self) -> None:
        """Synthesized deck covering all registered non-fallback templates must pass physical invariants."""
        slides = [
            SlideDefinition.model_validate({
                "template_id": tid,
                "title": f"All-In-One: {tid}",
                "subtitle": "SYNTHESIZED INTEGRATION TEST",
                "lead_text": "全登録テンプレート自動合成による物理境界およびテキスト整合性検証。" if spec.renders_header else None,
                "payload": spec.sample_payload,
            })
            for tid, spec in TEMPLATE_REGISTRY.items()
            if not spec.is_fallback
        ]
        deck = PresentationDeck(theme_name="corporate_blue", slides=slides)
        prs, summary = generate_presentation(deck)
        assert summary["fallback_slides"] == 0
        assert summary["rendered_slides"] == len(slides)
        assert len(prs.slides) == len(slides)

        _assert_deck_physical_integrity(prs, len(slides), "SynthesizedSlide")

    @pytest.mark.parametrize(
        "template_id",
        [tid for tid, s in TEMPLATE_REGISTRY.items() if not s.is_fallback],
    )
    def test_single_pane_max_capacity_stress(self, template_id: str) -> None:
        """Every registered template at maximum schema capacity must fit physically without overflow or collapse."""
        spec = TEMPLATE_REGISTRY[template_id]
        slide_def = SlideDefinition.model_validate({
            "template_id": template_id,
            "title": f"最大容量ストレステスト: {template_id}",
            "subtitle": "MAX CAPACITY VERIFICATION",
            "lead_text": None if not spec.renders_header else "スキーマ上限の要素数および日本語長文テキスト投入検証。",
            "payload": spec.sample_payload,
        })
        deck = PresentationDeck(theme_name="corporate_blue", slides=[slide_def])
        prs, summary = generate_presentation(deck)
        assert summary["fallback_slides"] == 0
        assert prs.slide_width is not None and prs.slide_height is not None

        slide = prs.slides[0]
        _assert_no_peer_text_overlap(slide, f"MaxSingle[{template_id}]")
        for sh_idx, shape in enumerate(slide.shapes):
            _assert_shape_physical_invariants(
                shape,
                int(prs.slide_width),
                int(prs.slide_height),
                f"MaxSingle[{template_id}].Shape[{sh_idx}]",
            )

    @pytest.mark.parametrize(
        "template_id,is_left",
        [
            (tid, is_left)
            for tid, spec in TEMPLATE_REGISTRY.items()
            if spec.supports_split
            for is_left in (True, False)
        ],
    )
    def test_split_max_capacity_invariants(
        self,
        template_id: str,
        is_left: bool,
    ) -> None:
        """All split-capable templates must adapt layout and pass physical checks in narrow and wide panes."""
        target_spec = TEMPLATE_REGISTRY[template_id]
        other_split_candidates = [
            tid for tid, s in TEMPLATE_REGISTRY.items()
            if s.supports_split and tid != template_id
        ]
        partner_id = other_split_candidates[0] if other_split_candidates else template_id
        partner_spec = TEMPLATE_REGISTRY[partner_id]

        if is_left:
            left_id, left_payload = template_id, target_spec.effective_split_payload
            right_id, right_payload = partner_id, partner_spec.effective_split_payload
            split_ratio = 0.30
        else:
            left_id, left_payload = partner_id, partner_spec.effective_split_payload
            right_id, right_payload = template_id, target_spec.effective_split_payload
            split_ratio = 0.70

        slide_def = SlideDefinition.model_validate({
            "template_id": left_id,
            "title": f"極限スプリット検証: {left_id} + {right_id} ({split_ratio})",
            "lead_text": "狭小ペインにおけるグリッド・スタック自動切替およびフォント縮小の物理境界検証。",
            "payload": left_payload,
            "right_template_id": right_id,
            "right_payload": right_payload,
            "split_ratio": split_ratio,
        })
        deck = PresentationDeck(theme_name="corporate_blue", slides=[slide_def])
        prs, summary = generate_presentation(deck)
        assert summary["fallback_slides"] == 0
        assert prs.slide_width is not None and prs.slide_height is not None

        slide = prs.slides[0]
        _, body_box = SlideMaster.get_body_box_with_lead(has_lead_text=True)
        left_box, right_box = SlideMaster.split_horizontal(body_box, split_ratio=split_ratio)

        _assert_pane_containment_and_separation(
            slide,
            int(body_box.top),
            left_box,
            right_box,
            f"Split[{left_id}+{right_id}@{split_ratio}]",
        )

        _assert_no_peer_text_overlap(slide, f"Split[{left_id}+{right_id}@{split_ratio}]")

        for sh_idx, shape in enumerate(slide.shapes):
            _assert_shape_physical_invariants(
                shape,
                int(prs.slide_width),
                int(prs.slide_height),
                f"Split[{left_id}+{right_id}@{split_ratio}].Shape[{sh_idx}]",
            )

    def test_max_text_capacity_physical_containment(self) -> None:
        """Slides saturating text lengths and item limits must fit physically without overflow."""
        max_title = "日" * 80
        max_desc = "説" * 150

        slides_data: list[dict[str, Any]] = [
            # 1. CORE-01 Cover Slide saturated
            {
                "template_id": "CORE-01",
                "title": max_title,
                "payload": {
                    "title": max_title,
                    "subtitle": max_desc,
                    "presenter": "発" * 80,
                    "date_str": "日" * 30,
                },
            },
            # 2. CORE-02 Agenda with 8 saturated items
            {
                "template_id": "CORE-02",
                "title": max_title,
                "lead_text": max_desc,
                "payload": {
                    "items": [
                        {
                            "number": f"0{i}",
                            "title": f"標題{i}_" + ("目" * 70),
                            "description": "説" * 100,
                            "is_active": i == 1,
                        }
                        for i in range(1, 9)
                    ]
                },
            },
            # 3. CORE-20 Process Flow with 5 saturated steps
            {
                "template_id": "CORE-20",
                "title": max_title,
                "lead_text": max_desc,
                "payload": {
                    "steps": [
                        {
                            "step_number": i,
                            "title": f"工程{i}_" + ("題" * 70),
                            "description": max_desc,
                        }
                        for i in range(1, 6)
                    ]
                },
            },
            # 4. CORE-40 2x2 Matrix saturated
            {
                "template_id": "CORE-40",
                "title": max_title,
                "lead_text": max_desc,
                "payload": {
                    "x_axis_label": "軸" * 40,
                    "y_axis_label": "軸" * 40,
                    "top_left": {"title": "象" * 70, "badge": "B" * 20, "description": max_desc},
                    "top_right": {"title": "象" * 70, "badge": "B" * 20, "description": max_desc},
                    "bottom_left": {"title": "象" * 70, "badge": "B" * 20, "description": max_desc},
                    "bottom_right": {"title": "象" * 70, "badge": "B" * 20, "description": max_desc},
                },
            },
            # 5. CORE-70 Dual Column Comparison saturated
            {
                "template_id": "CORE-70",
                "title": max_title,
                "lead_text": max_desc,
                "payload": {
                    "left_column": {
                        "header": "見" * 70,
                        "sub_header": "副" * 70,
                        "points": ["点" * 120 for _ in range(6)],
                    },
                    "right_column": {
                        "header": "見" * 70,
                        "sub_header": "副" * 70,
                        "points": ["点" * 120 for _ in range(6)],
                        "is_highlighted": True,
                    },
                },
            },
            # 6. CORE-70 Horizontal Split with CORE-10 saturated
            {
                "template_id": "CORE-10",
                "title": max_title,
                "lead_text": max_desc,
                "payload": {
                    "headers": [f"見{c}" for c in range(1, 5)],
                    "rows": [["字" * 35 for _ in range(4)] for _ in range(8)],
                },
                "right_template_id": "CORE-70",
                "right_payload": {
                    "left_column": {
                        "header": "見" * 50,
                        "points": ["点" * 80 for _ in range(4)],
                    },
                    "right_column": {
                        "header": "見" * 50,
                        "points": ["点" * 80 for _ in range(4)],
                        "is_highlighted": True,
                    },
                },
                "split_ratio": 0.5,
            },
        ]

        deck_dict = {
            "theme_name": "corporate_blue",
            "slides": slides_data,
        }
        deck = PresentationDeck.model_validate(deck_dict)
        prs, summary = generate_presentation(deck)
        assert summary["fallback_slides"] == 0
        assert summary["rendered_slides"] == len(slides_data)
        _assert_deck_physical_integrity(prs, len(slides_data), "MaxTextCapacityMultiTemplate")
