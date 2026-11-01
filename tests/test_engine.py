"""Comprehensive test suite for Slide Engine.

Verifies:
1. Dynamic dispatch resolution and function signatures across all 10 templates.
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

    def test_core50_payload_purity_rejects_extra_fields(self) -> None:
        """CORE-50 must reject forbidden extra attributes like image_path or binary data."""
        with pytest.raises(ValidationError, match="Extra inputs are not permitted"):
            Core50Payload.model_validate({
                "placeholder_label": "System Diagram",
                "caption": "Figure 1",
                "image_path": "/var/tmp/diagram.png",
            })

    @pytest.mark.parametrize(
        "slide_data,expected_match",
        [
            # Left pane table > 8 rows
            (
                {
                    "template_id": "CORE-10",
                    "title": "Split Overdensity",
                    "payload": {"headers": ["A", "B"], "rows": [[f"R{r}-1", f"R{r}-2"] for r in range(9)]},
                    "right_template_id": "CORE-50",
                    "right_payload": {"placeholder_label": "P", "caption": "C"},
                },
                "rows > 8 rows maximum",
            ),
            # Left pane table > 4 columns
            (
                {
                    "template_id": "CORE-10",
                    "title": "Split Overdensity",
                    "payload": {"headers": ["A", "B", "C", "D", "E"], "rows": [["1", "2", "3", "4", "5"]]},
                    "right_template_id": "CORE-50",
                    "right_payload": {"placeholder_label": "P", "caption": "C"},
                },
                "columns > 4 columns maximum",
            ),
            # Right pane process flow > 4 steps
            (
                {
                    "template_id": "CORE-50",
                    "title": "Split Overdensity",
                    "payload": {"placeholder_label": "P", "caption": "C"},
                    "right_template_id": "CORE-20",
                    "right_payload": {
                        "steps": [
                            {"step_number": i, "title": f"S{i}", "description": f"D{i}"}
                            for i in range(1, 6)
                        ]
                    },
                },
                "steps > 4 steps maximum",
            ),
            # Right pane comparison > 4 points
            (
                {
                    "template_id": "CORE-50",
                    "title": "Split Overdensity",
                    "payload": {"placeholder_label": "P", "caption": "C"},
                    "right_template_id": "CORE-70",
                    "right_payload": {
                        "left_column": {"header": "L", "points": [f"P{i}" for i in range(5)]},
                        "right_column": {"header": "R", "points": ["P1"]},
                    },
                },
                "points > 4 points maximum",
            ),
            # Split ratio < 0.25
            (
                {
                    "template_id": "CORE-10",
                    "title": "Split Ratio Out of Bounds",
                    "payload": {"headers": ["A"], "rows": [["1"]]},
                    "right_template_id": "CORE-50",
                    "right_payload": {"placeholder_label": "P", "caption": "C"},
                    "split_ratio": 0.20,
                },
                "greater than or equal to 0.25",
            ),
            # Split ratio > 0.75
            (
                {
                    "template_id": "CORE-10",
                    "title": "Split Ratio Out of Bounds",
                    "payload": {"headers": ["A"], "rows": [["1"]]},
                    "right_template_id": "CORE-50",
                    "right_payload": {"placeholder_label": "P", "caption": "C"},
                    "split_ratio": 0.80,
                },
                "less than or equal to 0.75",
            ),
        ],
    )
    def test_split_overdensity_and_ratio_bounds_raise_errors(
        self,
        slide_data: dict[str, Any],
        expected_match: str,
    ) -> None:
        """Split layout must reject payloads exceeding density limits or ratio boundaries."""
        with pytest.raises(ValidationError, match=expected_match):
            SlideDefinition.model_validate(slide_data)


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
    """Verifies that rendering errors fall back to CORE-00 safely."""

    def test_exception_triggers_core00_fallback(self) -> None:
        """A crashing renderer must be replaced by CORE-00 without halting execution."""
        prs = _initialize_presentation("corporate_blue")
        dispatch_table = build_dynamic_dispatch_table()

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

        notes = prs.slides[0].notes_slide.notes_text_frame.text
        assert "Simulated Database Render Failure" in notes

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

MAX_CAPACITY_PAYLOADS: dict[str, dict[str, Any]] = {
    "CORE-01": {
        "title": "次世代エンタープライズ統合データプラットフォーム刷新計画および移行戦略ロードマップ策定",
        "subtitle": "2026年度〜2028年度 中期経営計画における基幹系データベースとデータレイクハウスの完全統合アーキテクチャ設計書",
        "presenter": "デジタル変革推進本部 エンタープライズアーキテクチャグループ",
        "date_str": "2026年11月1日 (最終版)",
    },
    "CORE-02": {
        "items": [
            {
                "number": f"0{i}",
                "title": f"戦略フェーズ {i}: 全社データ基盤の統合とガバナンス強化",
                "description": "レガシーシステムの段階的廃止とクラウドネイティブ基盤への完全移行",
                "is_active": i == 1,
            }
            for i in range(1, 9)  # Max 8 items
        ]
    },
    "CORE-10": {
        "headers": [f"評価指標項目 {c}" for c in range(1, 6)],
        "rows": [[f"データ値 {r}-{c} (検証済)" for c in range(1, 6)] for r in range(1, 16)],  # Max 15 rows
        "highlight_row_index": 0,
    },
    "CORE-20": {
        "steps": [
            {
                "step_number": i,
                "title": f"工程 {i}: 基盤構築",
                "description": f"マイクロサービス第{i}群の自動デプロイと統合監視の確立",
            }
            for i in range(1, 7)  # Max 6 steps
        ]
    },
    "CORE-30": {
        "chart_type": "column",
        "categories": [f"Q{q}" for q in range(1, 9)],
        "series": [
            {"name": f"事業部 {s}", "values": [float(s * 10 + q) for q in range(1, 9)]}
            for s in range(1, 5)
        ],
    },
    "CORE-40": {
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
    },
    "CORE-50": {
        "placeholder_label": "ハイブリッドクラウド冗長化構成および閉域網ネットワークトポロジー図",
        "caption": "図1: 東西リージョン間でのアクティブ・アクティブ構成と自動フェイルオーバー経路の定義",
    },
    "CORE-60": {
        "metrics": [
            {
                "label": f"重要業績評価指標 {i}",
                "value": f"{i * 128.5}",
                "unit": "億円",
                "trend": f"前年同期比 +{i * 6.4}% 達成",
                "is_positive": True,
            }
            for i in range(1, 5)  # Max 4 cards
        ]
    },
    "CORE-70": {
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
    },
}

MAX_SPLIT_CAPACITY_PAYLOADS: dict[str, dict[str, Any]] = {
    **MAX_CAPACITY_PAYLOADS,
    "CORE-10": {
        "headers": [f"評価指標 {c}" for c in range(1, 5)],
        "rows": [[f"データ {r}-{c}" for c in range(1, 5)] for r in range(1, 9)],  # Max 8 rows in split
        "highlight_row_index": 0,
    },
    "CORE-20": {
        "steps": [
            {
                "step_number": i,
                "title": f"工程 {i}: 基盤構築",
                "description": f"マイクロサービス第{i}群の自動デプロイ",
            }
            for i in range(1, 5)  # Max 4 steps in split
        ]
    },
    "CORE-70": {
        "left_column": {
            "header": "現状課題 (As-Is)",
            "sub_header": "サイロ化された個別最適",
            "points": [f"部門間データ連携が手動バッチに依存 (課題{i})" for i in range(1, 5)],  # Max 4 pts in split
            "is_highlighted": False,
        },
        "right_column": {
            "header": "理想姿 (To-Be)",
            "sub_header": "リアルタイム統合基盤",
            "points": [f"全社リアルタイムAPI連携の実現 (施策{i})" for i in range(1, 5)],  # Max 4 pts in split
            "is_highlighted": True,
        },
    },
}


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

    em_width = sum(1.0 if ord(ch) > 127 else 0.58 for ch in p_text)
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


def _assert_shape_physical_invariants(shape: Any, slide_w: int, slide_h: int, label: str) -> int:
    """Verify a shape has positive non-collapsed dimensions, stays in bounds, and has zero text overflow."""
    left = int(shape.left or 0)
    top = int(shape.top or 0)
    width = int(shape.width or 0)
    height = int(shape.height or 0)
    tolerance = int(Inches(0.02))

    assert width >= int(Inches(0.35)), f"{label} width collapsed: {width / 914400:.2f}in"
    assert height >= int(Inches(0.20)), f"{label} height collapsed: {height / 914400:.2f}in"
    assert left >= 0 and top >= 0, f"{label} negative coordinates: ({left}, {top})"
    assert left + width <= slide_w + tolerance, f"{label} exceeds slide width"
    assert top + height <= slide_h + tolerance, f"{label} exceeds slide height"

    checked = 0
    if shape.has_text_frame:
        checked += _assert_text_frame_fits_physically(
            shape.text_frame,
            width / 12700.0,
            height / 12700.0,
            label,
        )

    if shape.has_table:
        table = shape.table
        col_w_pts = [col.width / 12700.0 for col in table.columns]
        row_h_pts = [row.height / 12700.0 for row in table.rows]
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


class TestAdaptiveLayoutAndPhysicalInvariants:
    """Verifies physical layout integrity, card non-collapse, zero collision, and text fit under stress."""

    def test_sample_deck_comprehensive_physical_integrity(self) -> None:
        """All 10 slides of sample_deck.json must pass zero-fallback, boundary, text-fit, and overlap checks."""
        sample_path = Path(__file__).parent.parent / "examples" / "sample_deck.json"
        deck = PresentationDeck.model_validate_json(sample_path.read_text(encoding="utf-8"))
        prs, summary = generate_presentation(deck)
        assert summary["total_slides"] == 10
        assert summary["rendered_slides"] == 10
        assert summary["fallback_slides"] == 0
        assert len(prs.slides) == 10
        assert prs.slide_width is not None and prs.slide_height is not None

        slide_w = int(prs.slide_width)
        slide_h = int(prs.slide_height)

        total_checked = 0
        for s_idx, slide in enumerate(prs.slides):
            _assert_no_peer_text_overlap(slide, f"SampleSlide[{s_idx}]")
            for sh_idx, shape in enumerate(slide.shapes):
                total_checked += _assert_shape_physical_invariants(
                    shape,
                    slide_w,
                    slide_h,
                    f"SampleSlide[{s_idx}].Shape[{sh_idx}]",
                )
        assert total_checked >= 65

        # Specifically verify the 10th slide (the heterogeneous split slide)
        slide_10 = prs.slides[9]
        if slide_10.has_notes_slide:
            assert "Slide Rendering Error" not in slide_10.notes_slide.notes_text_frame.text
        assert any(s.has_table for s in slide_10.shapes)

    @pytest.mark.parametrize(
        "template_id",
        ["CORE-01", "CORE-02", "CORE-10", "CORE-20", "CORE-30", "CORE-40", "CORE-50", "CORE-60", "CORE-70"],
    )
    def test_single_pane_max_capacity_stress(self, template_id: str) -> None:
        """Every template at maximum schema capacity must fit physically without overflow or collapse."""
        slide_def = SlideDefinition.model_validate({
            "template_id": template_id,
            "title": f"最大容量ストレステスト: {template_id}",
            "subtitle": "MAX CAPACITY VERIFICATION",
            "lead_text": None if template_id == "CORE-01" else "スキーマ上限の要素数および日本語長文テキスト投入検証。",
            "payload": MAX_CAPACITY_PAYLOADS[template_id],
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
        "left_id,right_id,split_ratio",
        [
            ("CORE-20", "CORE-60", 0.3),   # Compact process in 30% + 4 KPIs in 70%
            ("CORE-60", "CORE-20", 0.7),   # 4 KPIs in 70% + Compact process in 30%
            ("CORE-70", "CORE-10", 0.35),  # Vertical stacked comparison in 35% + 8-row table in 65%
            ("CORE-10", "CORE-40", 0.5),   # 8x4 Table + 2x2 Matrix at 50/50 split
            ("CORE-40", "CORE-70", 0.45),  # 2x2 Matrix in 45% + Comparison in 55%
            ("CORE-30", "CORE-50", 0.5),   # Native Chart + Visual Guide Frame
        ],
    )
    def test_extreme_split_max_capacity_invariants(
        self,
        left_id: str,
        right_id: str,
        split_ratio: float,
    ) -> None:
        """Max-capacity payloads in extreme split ratios must adapt layout and pass all physical checks."""
        slide_def = SlideDefinition.model_validate({
            "template_id": left_id,
            "title": f"極限スプリット検証: {left_id} + {right_id} ({split_ratio})",
            "lead_text": "狭小ペインにおけるグリッド・スタック自動切替およびフォント縮小の物理境界検証。",
            "payload": MAX_SPLIT_CAPACITY_PAYLOADS[left_id],
            "right_template_id": right_id,
            "right_payload": MAX_SPLIT_CAPACITY_PAYLOADS[right_id],
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
