"""Integration and orchestration layer.

Manages dynamic template dispatching, per-slide transactional rendering,
and fail-soft fallback handling.
"""

from __future__ import annotations

import os
import time
import traceback
from collections.abc import Callable
from pathlib import Path
from typing import TYPE_CHECKING, Any

from pptx import Presentation
from pptx.slide import Slide

from slide_engine.core_base import BoundingBox, ShapeBuilder, SlideMaster, render_standard_header
from slide_engine.models import (
    TEMPLATE_REGISTRY,
    PresentationDeck,
    SlideDefinition,
    get_loaded_template_modules,
)

if TYPE_CHECKING:
    from pptx.presentation import Presentation as PresentationType

RenderFunction = Callable[[Slide, Any, BoundingBox], None]


def build_dynamic_dispatch_table() -> dict[str, RenderFunction]:
    """Dynamically bind render functions from discovered template modules.

    Returns:
        Mapping of template_id to rendering callable.
    """
    dispatch_table: dict[str, RenderFunction] = {}
    modules = get_loaded_template_modules()

    for template_id, mod in modules.items():
        suffix = template_id.replace("CORE-", "").lower()
        render_fn = getattr(mod, f"render_core_{suffix}", None)
        if not callable(render_fn):
            raise ValueError(f"Module '{mod.__name__}' must export callable 'render_core_{suffix}'.")

        dispatch_table[template_id] = render_fn

    return dispatch_table


def _get_asset_search_dirs() -> list[Path]:
    """Return ordered list of canonical directory paths to search for assets."""
    search_dirs: list[Path] = []
    if "SLIDE_ENGINE_ASSETS_DIR" in os.environ:
        search_dirs.append(Path(os.environ["SLIDE_ENGINE_ASSETS_DIR"]).resolve())
    search_dirs.append(Path("/app/assets").resolve())
    search_dirs.append(Path("assets").resolve())
    search_dirs.append((Path(__file__).resolve().parent.parent.parent / "assets").resolve())
    return search_dirs


def _resolve_safe_asset_path(theme_name: str) -> Path | None:
    """Resolve and verify theme asset path ensuring strict containment within asset directories."""
    filename = f"{theme_name}.pptx"
    for base_dir in _get_asset_search_dirs():
        if not base_dir.is_dir():
            continue
        try:
            candidate = (base_dir / filename).resolve()
            if candidate.is_relative_to(base_dir) and candidate.is_file():
                return candidate
        except (ValueError, OSError):
            continue
    return None


def _initialize_presentation(theme_name: str) -> PresentationType:
    """Initialize Presentation using master assets or clean fallback.

    Args:
        theme_name: Name of the master template asset.

    Returns:
        Initialized Presentation object.
    """
    safe_path = _resolve_safe_asset_path(theme_name)
    if safe_path is not None:
        try:
            prs = Presentation(str(safe_path))
            ShapeBuilder.purge_existing_slides(prs)
            return prs
        except Exception:
            # Fall through to clean blank master on corrupted or invalid asset
            pass

    prs = Presentation()
    prs.slide_width = SlideMaster.SLIDE_WIDTH
    prs.slide_height = SlideMaster.SLIDE_HEIGHT
    return prs


def _execute_slide_render(
    slide: Slide,
    slide_def: SlideDefinition,
    dispatch_table: dict[str, RenderFunction],
) -> None:
    """Execute standard layout resolution and rendering with early returns."""
    # 1. Header & Lead Text Render (Driven by TemplateSpec metadata)
    spec = TEMPLATE_REGISTRY.get(slide_def.template_id)
    if spec is None or spec.renders_header:
        render_standard_header(
            slide=slide,
            title=slide_def.title,
            subtitle=slide_def.subtitle,
            lead_text=slide_def.lead_text,
        )

    # 2. Determine target content BoundingBox
    has_lead = bool(slide_def.lead_text)
    _, body_box = SlideMaster.get_body_box_with_lead(has_lead)

    # 3. Dispatch rendering
    primary_fn = dispatch_table.get(slide_def.template_id)
    if not primary_fn:
        raise ValueError(f"No renderer registered for '{slide_def.template_id}'.")

    # Single-pane layout
    if slide_def.right_payload is None:
        primary_fn(slide, slide_def.payload, body_box)
        return

    # Split-pane layout
    right_id = slide_def.right_template_id or slide_def.template_id
    right_fn = dispatch_table.get(right_id)
    if not right_fn:
        raise ValueError(f"No renderer registered for right pane '{right_id}'.")

    left_box, right_box = SlideMaster.split_horizontal(body_box, slide_def.split_ratio)
    primary_fn(slide, slide_def.payload, left_box)
    right_fn(slide, slide_def.right_payload, right_box)


def _handle_slide_fallback(
    slide: Slide,
    slide_def: SlideDefinition,
    exc: Exception,
    dispatch_table: dict[str, RenderFunction],
) -> None:
    """Record exception in speaker notes and render system fallback slide."""
    # 1. Record trace in speaker notes for troubleshooting
    try:
        notes_slide = slide.notes_slide
        notes_tf = notes_slide.notes_text_frame
        if notes_tf is not None:
            notes_tf.text = f"Slide Rendering Error:\n{traceback.format_exc()}"
    except Exception:
        pass

    # 2. Reset canvas
    ShapeBuilder.clean_slide_canvas(slide)

    # 3. Resolve fallback template and render
    fallback_tid, fallback_spec = next(
        ((tid, spec) for tid, spec in TEMPLATE_REGISTRY.items() if spec.is_fallback),
        (None, None),
    )
    if not (fallback_tid and fallback_spec):
        return

    fallback_fn = dispatch_table.get(fallback_tid)
    if not fallback_fn:
        return

    fallback_payload = fallback_spec.create_fallback_payload(
        error_title=f"Error in {slide_def.template_id}",
        error_message=str(exc),
        raw_dump=str(slide_def.model_dump_json()),
    )
    fallback_fn(slide, fallback_payload, SlideMaster.CONTENT_BOX)


def _render_slide_with_failsoft(
    prs: PresentationType,
    slide_def: SlideDefinition,
    dispatch_table: dict[str, RenderFunction],
) -> bool:
    """Render a slide inside a safe transactional boundary.

    Returns:
        True if rendered normally, False if fallback was triggered.
    """
    spec = TEMPLATE_REGISTRY.get(slide_def.template_id)
    slide_type = spec.layout_type if spec else "body"
    layout = SlideMaster.resolve_slide_layout(prs, slide_type)
    slide = prs.slides.add_slide(layout)

    # Purge master default placeholder shapes immediately
    ShapeBuilder.clean_slide_canvas(slide)

    try:
        _execute_slide_render(slide, slide_def, dispatch_table)
        return True
    except Exception as exc:
        _handle_slide_fallback(slide, slide_def, exc, dispatch_table)
        return False


def generate_presentation(
    deck: PresentationDeck,
    dispatch_table: dict[str, RenderFunction] | None = None,
) -> tuple[PresentationType, dict[str, Any]]:
    """Generate a PowerPoint presentation deck from validated specification.

    Args:
        deck: Validated PresentationDeck instance.
        dispatch_table: Optional custom dispatch table (built automatically if None).

    Returns:
        Tuple of (Presentation instance, telemetry_summary_dict).
    """
    start_time = time.perf_counter()
    table = dispatch_table or build_dynamic_dispatch_table()
    prs = _initialize_presentation(deck.theme_name)

    rendered_count = 0
    fallback_count = 0

    for slide_def in deck.slides:
        is_success = _render_slide_with_failsoft(prs, slide_def, table)
        if is_success:
            rendered_count += 1
        else:
            fallback_count += 1

    elapsed_ms = (time.perf_counter() - start_time) * 1000.0
    summary = {
        "event": "deck_generation_summary",
        "total_slides": len(deck.slides),
        "rendered_slides": rendered_count,
        "fallback_slides": fallback_count,
        "duration_ms": round(elapsed_ms, 2),
    }

    return prs, summary
