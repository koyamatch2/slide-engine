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


class ProcessStep(BasePayload):
    """Step definition for Process Flow."""

    step_number: int = Field(ge=1, le=10)
    title: ShortText
    description: SafeText


class Core20Payload(BasePayload):
    """Payload for Process Flow slide (CORE-20)."""

    steps: list[ProcessStep] = Field(min_length=2, max_length=6)


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


# ----------------------------------------------------------------------
# Template Registry Mapping
# ----------------------------------------------------------------------


class TemplateSpec(BasePayload):
    """Metadata specification for each template."""

    payload_type: type[BasePayload]
    supports_split: bool
    description: str


TEMPLATE_REGISTRY: dict[str, TemplateSpec] = {
    "CORE-00": TemplateSpec(payload_type=Core00Payload, supports_split=False, description="System Fallback"),
    "CORE-01": TemplateSpec(payload_type=Core01Payload, supports_split=False, description="Cover Slide"),
    "CORE-02": TemplateSpec(payload_type=Core02Payload, supports_split=False, description="Agenda / Section"),
    "CORE-10": TemplateSpec(payload_type=Core10Payload, supports_split=True, description="Data Table"),
    "CORE-20": TemplateSpec(payload_type=Core20Payload, supports_split=True, description="Process Flow"),
    "CORE-30": TemplateSpec(payload_type=Core30Payload, supports_split=True, description="Native Chart"),
    "CORE-40": TemplateSpec(payload_type=Core40Payload, supports_split=True, description="2x2 Matrix"),
    "CORE-50": TemplateSpec(payload_type=Core50Payload, supports_split=True, description="Visual Guide Frame"),
    "CORE-60": TemplateSpec(payload_type=Core60Payload, supports_split=True, description="KPI Summary Cards"),
    "CORE-70": TemplateSpec(payload_type=Core70Payload, supports_split=True, description="Comparison Box"),
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
    split_ratio: float = Field(default=0.5, ge=0.2, le=0.8)

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
        return self


class PresentationDeck(BaseModel):
    """Root model for an entire presentation deck."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    theme_name: ShortText = Field(default="corporate_blue")
    slides: list[SlideDefinition] = Field(min_length=1, max_length=MAX_SLIDES_PER_DECK)