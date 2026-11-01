"""Data contract layer defining Pydantic V2 schemas and strict boundary validation.

This module is completely side-effect free (zero I/O, zero direct logging).
It enforces text truncation limits, resource ceilings, and template registries.
"""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import BaseModel, ConfigDict, Field, StringConstraints, model_validator

# Maximum limits to prevent denial of service and memory exhaustion
MAX_SLIDES_PER_DECK = 100
MAX_SHORT_TEXT_LEN = 120
MAX_SAFE_TEXT_LEN = 800
MAX_LONG_TEXT_LEN = 2500

# String type constraints
ShortText = Annotated[str, StringConstraints(strip_whitespace=True, max_length=MAX_SHORT_TEXT_LEN)]
SafeText = Annotated[str, StringConstraints(strip_whitespace=True, max_length=MAX_SAFE_TEXT_LEN)]


class BasePayload(BaseModel):
    """Base model enforcing strict configuration for all template payloads."""

    model_config = ConfigDict(extra="forbid", frozen=True, validate_assignment=True)

    def validate_split_density(self, pane_name: str, template_id: str) -> None:
        """Validate payload density when rendered in a split pane. Default: no restriction."""
        return


# ----------------------------------------------------------------------
# Template Payload Definitions (CORE-00 to CORE-70)
# ----------------------------------------------------------------------


class Core00Payload(BasePayload):
    """Payload for system fallback slide (CORE-00)."""

    error_title: ShortText = Field(default="Slide Rendering Error")
    error_message: SafeText = Field(default="An unhandled exception occurred during rendering.")
    raw_payload_dump: str = Field(default="")


class Core01Payload(BasePayload):
    """Payload for Cover slide (CORE-01)."""

    title: ShortText
    subtitle: SafeText | None = None
    presenter: ShortText | None = None
    date_str: ShortText | None = None


class AgendaItem(BasePayload):
    """Item for Agenda slide."""

    number: ShortText
    title: ShortText
    description: SafeText | None = None
    is_active: bool = False


class Core02Payload(BasePayload):
    """Payload for Agenda / Section slide (CORE-02)."""

    items: list[AgendaItem] = Field(min_length=1, max_length=8)


class Core10Payload(BasePayload):
    """Payload for Data Table slide (CORE-10)."""

    headers: list[ShortText] = Field(min_length=1, max_length=8)
    rows: list[list[SafeText]] = Field(min_length=1, max_length=15)
    highlight_row_index: int | None = Field(default=None, ge=0)

    def validate_split_density(self, pane_name: str, template_id: str) -> None:
        """Enforce row and column limits for split layout."""
        if len(self.rows) > 8:
            raise ValueError(
                f"{pane_name} table '{template_id}' in split layout exceeds legibility limit: "
                f"{len(self.rows)} rows > 8 rows maximum."
            )
        if len(self.headers) > 4:
            raise ValueError(
                f"{pane_name} table '{template_id}' in split layout exceeds legibility limit: "
                f"{len(self.headers)} columns > 4 columns maximum."
            )


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


class ChartSeries(BasePayload):
    """Data series for native charts."""

    name: ShortText
    values: list[float] = Field(min_length=1, max_length=20)


class Core30Payload(BasePayload):
    """Payload for Native OpenXML Chart slide (CORE-30)."""

    chart_type: Literal["column", "bar", "line", "pie"]
    categories: list[ShortText] = Field(min_length=1, max_length=20)
    series: list[ChartSeries] = Field(min_length=1, max_length=6)

    @model_validator(mode="after")
    def validate_series_length(self) -> Core30Payload:
        """Validate that all series have exactly the same length as categories."""
        expected_len = len(self.categories)
        for s in self.series:
            if len(s.values) != expected_len:
                msg = (
                    f"Series '{s.name}' length ({len(s.values)}) does not match "
                    f"categories length ({expected_len})."
                )
                raise ValueError(msg)
        return self


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


class Core50Payload(BasePayload):
    """Payload for Visual Guide Frame slide (CORE-50).

    Used for rendering a manual screenshot/diagram guide frame and caption.
    Completely eliminates external file I/O and Pillow dependency.
    """

    placeholder_label: ShortText
    caption: SafeText | None = None


class KpiMetric(BasePayload):
    """Metric item for KPI summary card."""

    value: ShortText
    unit: ShortText | None = None
    label: ShortText
    trend: ShortText | None = None
    is_positive: bool | None = None


class Core60Payload(BasePayload):
    """Payload for KPI Summary Cards slide (CORE-60)."""

    metrics: list[KpiMetric] = Field(min_length=2, max_length=4)


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


# ----------------------------------------------------------------------
# Template Registry Mapping
# ----------------------------------------------------------------------


class TemplateSpec(BasePayload):
    """Metadata specification for each template."""

    payload_type: type[BasePayload]
    supports_split: bool
    description: str
    sample_payload: dict[str, Any]
    split_sample_payload: dict[str, Any] | None = None

    @property
    def effective_split_payload(self) -> dict[str, Any]:
        """Return split-optimized payload, or default sample_payload if none specified."""
        return self.split_sample_payload if self.split_sample_payload is not None else self.sample_payload


_SAMPLE_CORE00: dict[str, Any] = {
    "error_title": "システムエラー",
    "error_message": "スライドレンダリング処理中に例外が発生しました。",
    "raw_payload_dump": "{}",
}

_SAMPLE_CORE01: dict[str, Any] = {
    "title": "次世代エンタープライズ統合データプラットフォーム刷新計画および移行戦略ロードマップ策定",
    "subtitle": "2026年度〜2028年度 中期経営計画における基幹系データベースとデータレイクハウスの完全統合アーキテクチャ設計書",
    "presenter": "デジタル変革推進本部 エンタープライズアーキテクチャグループ",
    "date_str": "2026年11月1日 (最終版)",
}

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

_SAMPLE_CORE10: dict[str, Any] = {
    "headers": [f"評価指標項目 {c}" for c in range(1, 6)],
    "rows": [[f"データ値 {r}-{c} (検証済)" for c in range(1, 6)] for r in range(1, 16)],
    "highlight_row_index": 0,
}

_SAMPLE_CORE10_SPLIT: dict[str, Any] = {
    "headers": [f"評価指標 {c}" for c in range(1, 5)],
    "rows": [[f"データ {r}-{c}" for c in range(1, 5)] for r in range(1, 9)],
    "highlight_row_index": 0,
}

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

_SAMPLE_CORE30: dict[str, Any] = {
    "chart_type": "column",
    "categories": [f"Q{q}" for q in range(1, 9)],
    "series": [
        {"name": f"事業部 {s}", "values": [float(s * 10 + q) for q in range(1, 9)]}
        for s in range(1, 5)
    ],
}

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

_SAMPLE_CORE50: dict[str, Any] = {
    "placeholder_label": "ハイブリッドクラウド冗長化構成および閉域網ネットワークトポロジー図",
    "caption": "図1: 東西リージョン間でのアクティブ・アクティブ構成と自動フェイルオーバー経路の定義",
}

_SAMPLE_CORE60: dict[str, Any] = {
    "metrics": [
        {
            "label": f"重要業績評価指標 {i}",
            "value": f"{i * 128.5}",
            "unit": "億円",
            "trend": f"前年同期比 +{i * 6.4}% 達成",
            "is_positive": True,
        }
        for i in range(1, 5)
    ]
}

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


TEMPLATE_REGISTRY: dict[str, TemplateSpec] = {
    "CORE-00": TemplateSpec(
        payload_type=Core00Payload,
        supports_split=False,
        description="System Fallback",
        sample_payload=_SAMPLE_CORE00,
    ),
    "CORE-01": TemplateSpec(
        payload_type=Core01Payload,
        supports_split=False,
        description="Cover Slide",
        sample_payload=_SAMPLE_CORE01,
    ),
    "CORE-02": TemplateSpec(
        payload_type=Core02Payload,
        supports_split=False,
        description="Agenda / Section",
        sample_payload=_SAMPLE_CORE02,
    ),
    "CORE-10": TemplateSpec(
        payload_type=Core10Payload,
        supports_split=True,
        description="Data Table",
        sample_payload=_SAMPLE_CORE10,
        split_sample_payload=_SAMPLE_CORE10_SPLIT,
    ),
    "CORE-20": TemplateSpec(
        payload_type=Core20Payload,
        supports_split=True,
        description="Process Flow",
        sample_payload=_SAMPLE_CORE20,
        split_sample_payload=_SAMPLE_CORE20_SPLIT,
    ),
    "CORE-30": TemplateSpec(
        payload_type=Core30Payload,
        supports_split=True,
        description="Native Chart",
        sample_payload=_SAMPLE_CORE30,
    ),
    "CORE-40": TemplateSpec(
        payload_type=Core40Payload,
        supports_split=True,
        description="2x2 Matrix",
        sample_payload=_SAMPLE_CORE40,
    ),
    "CORE-50": TemplateSpec(
        payload_type=Core50Payload,
        supports_split=True,
        description="Visual Guide Frame",
        sample_payload=_SAMPLE_CORE50,
    ),
    "CORE-60": TemplateSpec(
        payload_type=Core60Payload,
        supports_split=True,
        description="KPI Summary Cards",
        sample_payload=_SAMPLE_CORE60,
    ),
    "CORE-70": TemplateSpec(
        payload_type=Core70Payload,
        supports_split=True,
        description="Comparison Box",
        sample_payload=_SAMPLE_CORE70,
        split_sample_payload=_SAMPLE_CORE70_SPLIT,
    ),
}


# ----------------------------------------------------------------------
# Slide and Deck Definitions
# ----------------------------------------------------------------------


def _validate_single_pane(template_id: str, payload_data: Any, is_right: bool = False) -> BasePayload:
    """Validate a single template payload against the registry.

    Args:
        template_id: Template identifier (e.g., 'CORE-10').
        payload_data: Raw payload data (dict or model).
        is_right: Whether this is validating the right-hand split pane.

    Returns:
        Validated BasePayload instance.
    """
    spec = TEMPLATE_REGISTRY.get(template_id)
    if not spec:
        raise ValueError(f"Unknown template_id: '{template_id}'.")

    if is_right and not spec.supports_split:
        raise ValueError(f"Template '{template_id}' does not support horizontal split.")

    if isinstance(payload_data, spec.payload_type):
        return payload_data
    return spec.payload_type.model_validate(payload_data)


class SlideDefinition(BaseModel):
    """Complete specification for a single presentation slide."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    template_id: str
    title: ShortText = Field(default="")
    subtitle: ShortText | None = None
    lead_text: SafeText | None = None
    payload: Any

    # Horizontal split layout options
    right_template_id: str | None = None
    right_payload: Any | None = None
    split_ratio: float = Field(default=0.5, ge=0.25, le=0.75)

    @model_validator(mode="after")
    def validate_slide_integrity(self) -> SlideDefinition:
        """Validate left and optional right split payloads with zero deep nesting."""
        # 1. Validate primary (left) pane
        left_obj = _validate_single_pane(self.template_id, self.payload, is_right=False)
        object.__setattr__(self, "payload", left_obj)

        # 2. Return early if no split is requested
        if self.right_payload is None and self.right_template_id is None:
            return self

        # 3. Guard against split on primary template that does not support split
        left_spec = TEMPLATE_REGISTRY[self.template_id]
        if not left_spec.supports_split:
            raise ValueError(f"Primary template '{self.template_id}' does not support horizontal split.")

        # 4. Resolve and validate right pane
        effective_right_id = self.right_template_id or self.template_id
        if self.right_payload is None:
            raise ValueError("right_payload must be provided when right_template_id is specified.")

        right_obj = _validate_single_pane(effective_right_id, self.right_payload, is_right=True)
        object.__setattr__(self, "right_template_id", effective_right_id)
        object.__setattr__(self, "right_payload", right_obj)

        # 5. Enforce split pane physical density constraints via polymorphic payload methods
        left_obj.validate_split_density("Left", self.template_id)
        right_obj.validate_split_density("Right", effective_right_id)
        return self


class PresentationDeck(BaseModel):
    """Root model for an entire presentation deck."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    theme_name: ShortText = Field(default="corporate_blue")
    slides: list[SlideDefinition] = Field(min_length=1, max_length=MAX_SLIDES_PER_DECK)
