"""Data contract layer defining Pydantic V2 schemas and strict boundary validation.

Enforces text truncation limits, resource ceilings, and dynamically discovers
template schemas and specifications from plugin render modules.
"""

from __future__ import annotations

import glob
import importlib
import os
from typing import Annotated, Any, Literal

from pydantic import BaseModel, ConfigDict, Field, StringConstraints, model_validator

# Maximum limits to prevent denial of service and memory exhaustion
MAX_SLIDES_PER_DECK = 100
MAX_SHORT_TEXT_LEN = 80
MAX_SAFE_TEXT_LEN = 150

# String type constraints
ShortText = Annotated[str, StringConstraints(strip_whitespace=True, max_length=MAX_SHORT_TEXT_LEN)]
SafeText = Annotated[str, StringConstraints(strip_whitespace=True, max_length=MAX_SAFE_TEXT_LEN)]
ThemeName = Annotated[
    str,
    StringConstraints(
        strip_whitespace=True,
        min_length=1,
        max_length=MAX_SHORT_TEXT_LEN,
        pattern=r"^[a-zA-Z0-9_-]+$",
    ),
]


class BasePayload(BaseModel):
    """Base model enforcing strict configuration for all template payloads."""

    model_config = ConfigDict(extra="forbid", frozen=True, validate_assignment=True)

    def validate_split_density(self, pane_name: str, template_id: str) -> None:
        """Validate payload density when rendered in a split pane. Default: no restriction."""
        return


# ----------------------------------------------------------------------
# Template Specification Structure
# ----------------------------------------------------------------------


class TemplateSpec(BasePayload):
    """Metadata specification for each template plugin."""

    payload_type: type[BasePayload]
    supports_split: bool
    description: str
    sample_payload: dict[str, Any]
    split_sample_payload: dict[str, Any] | None = None
    invalid_sample_payloads: list[dict[str, Any]] = Field(default_factory=list)
    renders_header: bool = True
    layout_type: Literal["cover", "agenda", "body", "blank"] = "body"
    is_fallback: bool = False

    @property
    def effective_split_payload(self) -> dict[str, Any]:
        """Return split-optimized payload, or default sample_payload if none specified."""
        return self.split_sample_payload if self.split_sample_payload is not None else self.sample_payload

    def create_fallback_payload(
        self,
        error_title: str,
        error_message: str,
        raw_dump: str,
    ) -> BasePayload:
        """Create a fallback payload instance if this template is configured as a fallback."""
        if not self.is_fallback:
            raise TypeError(f"Template '{self.description}' is not configured as a fallback template.")
        return self.payload_type.model_validate({
            "error_title": error_title,
            "error_message": error_message,
            "raw_payload_dump": raw_dump,
        })


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
    payload: BasePayload

    # Horizontal split layout options
    right_template_id: str | None = None
    right_payload: BasePayload | None = None
    split_ratio: float = Field(default=0.5, ge=0.25, le=0.75)

    @model_validator(mode="before")
    @classmethod
    def prepare_panes(cls, data: Any) -> Any:
        """Validate and construct pane payloads before model instantiation."""
        if not isinstance(data, dict):
            return data

        data_dict = dict(data)
        tid = data_dict.get("template_id")
        if not isinstance(tid, str):
            return data_dict

        # 1. Validate primary payload
        raw_payload = data_dict.get("payload")
        data_dict["payload"] = _validate_single_pane(tid, raw_payload, is_right=False)

        # 2. Handle split pane
        raw_right_payload = data_dict.get("right_payload")
        right_tid = data_dict.get("right_template_id")

        if raw_right_payload is None and right_tid is None:
            return data_dict
        if raw_right_payload is None or right_tid is None:
            raise ValueError(
                "Both 'right_template_id' and 'right_payload' must be provided together for split layout."
            )

        data_dict["right_template_id"] = right_tid
        data_dict["right_payload"] = _validate_single_pane(right_tid, raw_right_payload, is_right=True)
        return data_dict

    @model_validator(mode="after")
    def validate_split_constraints(self) -> SlideDefinition:
        """Validate split layout rules and density constraints with zero __setattr__ bypass."""
        if self.right_payload is None:
            return self

        left_spec = TEMPLATE_REGISTRY.get(self.template_id)
        if left_spec and not left_spec.supports_split:
            raise ValueError(f"Primary template '{self.template_id}' does not support horizontal split.")

        if not self.right_template_id:
            raise ValueError("right_template_id must be specified when right_payload is present.")

        self.payload.validate_split_density("Left", self.template_id)
        self.right_payload.validate_split_density("Right", self.right_template_id)
        return self


class PresentationDeck(BaseModel):
    """Root model for an entire presentation deck."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    theme_name: ThemeName = Field(default="corporate_blue")
    slides: list[SlideDefinition] = Field(min_length=1, max_length=MAX_SLIDES_PER_DECK)


# ----------------------------------------------------------------------
# Dynamic Discovery and Module Initialization (At Bottom to Avoid Cycles)
# ----------------------------------------------------------------------

_LOADED_TEMPLATE_MODULES: dict[str, Any] = {}


def get_loaded_template_modules() -> dict[str, Any]:
    """Return cached dictionary of loaded template modules."""
    return _LOADED_TEMPLATE_MODULES


def discover_template_specs() -> dict[str, TemplateSpec]:
    """Dynamically discover and register TemplateSpec instances from render_core_*.py modules.

    Scans the package directory for 'render_core_*.py' files and retrieves their
    top-level 'SPEC' attribute conforming to TemplateSpec.

    Returns:
        Mapping of template_id (e.g. 'CORE-10') to TemplateSpec.

    Raises:
        RuntimeError: If any template module fails to import.
        ValueError: If any template module does not export a valid TemplateSpec.
    """
    registry: dict[str, TemplateSpec] = {}
    current_dir = os.path.dirname(os.path.abspath(__file__))
    pattern = os.path.join(current_dir, "render_core_*.py")

    for file_path in sorted(glob.glob(pattern)):
        module_name = os.path.splitext(os.path.basename(file_path))[0]
        suffix = module_name.replace("render_core_", "")
        template_id = f"CORE-{suffix.upper()}"

        try:
            mod = importlib.import_module(f"slide_engine.{module_name}")
        except Exception as exc:
            raise RuntimeError(f"Failed to import template module '{module_name}': {exc}") from exc

        spec = getattr(mod, "SPEC", None)
        if not isinstance(spec, TemplateSpec):
            raise ValueError(f"Module '{module_name}' must export a valid 'SPEC: TemplateSpec' attribute.")

        _LOADED_TEMPLATE_MODULES[template_id] = mod
        registry[template_id] = spec

    return registry


# Initial discovery of all available template specifications
TEMPLATE_REGISTRY: dict[str, TemplateSpec] = discover_template_specs()
