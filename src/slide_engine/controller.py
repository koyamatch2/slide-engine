"""Integration and orchestration layer.

Manages dynamic template dispatching, per-slide transactional rendering,
fail-soft fallback handling, and structured telemetry emission to STDERR.
"""

from __future__ import annotations

import glob
import importlib
import json
import os
import sys
import time
import traceback
from collections.abc import Callable
from typing import TYPE_CHECKING, Any

from pptx import Presentation

from slide_engine.core_base import BoundingBox, ShapeBuilder, SlideMaster, render_standard_header
from slide_engine.models import Core00Payload, PresentationDeck, SlideDefinition

if TYPE_CHECKING:
    from pptx.presentation import Presentation as PresentationType
    from pptx.slide import Slide

RenderFunction = Callable[[Any, Any, BoundingBox], None]


def build_dynamic_dispatch_table() -> dict[str, RenderFunction]:
    """Dynamically discover and bind all render_core_XX modules.

    Scans the local directory for 'render_core_*.py' files and imports their
    render function conforming to the naming convention 'render_core_XX'.

    Returns:
        Mapping of template_id to rendering callable.
    """
    dispatch_table: dict[str, RenderFunction] = {}
    current_dir = os.path.dirname(os.path.abspath(__file__))
    pattern = os.path.join(current_dir, "render_core_*.py")

    for file_path in glob.glob(pattern):
        module_name = os.path.splitext(os.path.basename(file_path))[0]
        suffix = module_name.replace("render_core_", "")
        template_id = f"CORE-{suffix.upper()}"

        try:
            mod = importlib.import_module(f"slide_engine.{module_name}")
            render_fn = getattr(mod, f"render_core_{suffix.lower()}", None)
            if callable(render_fn):
                dispatch_table[template_id] = render_fn
        except Exception as exc:
            sys.stderr.write(f"Warning: Failed to load renderer module '{module_name}': {exc}\n")

    return dispatch_table


def _initialize_presentation(theme_name: str) -> PresentationType:
    """Initialize Presentation using master assets or clean fallback.

    Args:
        theme_name: Name of the master template asset.

    Returns:
        Initialized Presentation object.
    """
    asset_path = os.path.join("/app/assets", f"{theme_name}.pptx")
    if not os.path.exists(asset_path):
        asset_path = os.path.join("assets", f"{theme_name}.pptx")

    if os.path.exists(asset_path):
        prs = Presentation(asset_path)
        # Purge any pre-existing dummy slides from the master file
        for i in range(len(prs.slides) - 1, -1, -1):
            r_id = prs.slides._sldIdLst[i].rId  # noqa: SLF001
            prs.part.drop_rel(r_id)
            del prs.slides._sldIdLst[i]  # noqa: SLF001
        return prs

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
    # 1. Header & Lead Text Render (Excluded on cover slide)
    if slide_def.template_id != "CORE-01":
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


def _render_slide_with_failsoft(
    prs: PresentationType,
    slide_def: SlideDefinition,
    dispatch_table: dict[str, RenderFunction],
) -> bool:
    """Render a slide inside a safe transactional boundary.

    Returns:
        True if rendered normally, False if fallback was triggered.
    """
    slide_type = "cover" if slide_def.template_id == "CORE-01" else "body"
    layout = SlideMaster.resolve_slide_layout(prs, slide_type)
    slide = prs.slides.add_slide(layout)

    # Purge master default placeholder shapes immediately
    ShapeBuilder.clean_slide_canvas(slide)

    try:
        _execute_slide_render(slide, slide_def, dispatch_table)
        return True
    except Exception as exc:
        # Record trace in speaker notes for troubleshooting
        tb_str = traceback.format_exc()
        try:
            notes_slide = slide.notes_slide
            notes_tf = notes_slide.notes_text_frame
            notes_tf.text = f"Slide Rendering Error:\n{tb_str}"
        except Exception:
            pass

        # Clean canvas and render fallback CORE-00 slide
        ShapeBuilder.clean_slide_canvas(slide)
        fallback_fn = dispatch_table.get("CORE-00")
        if fallback_fn:
            fallback_payload = Core00Payload(
                error_title=f"Error in {slide_def.template_id}",
                error_message=str(exc),
                raw_payload_dump=str(slide_def.model_dump_json()),
            )
            fallback_fn(slide, fallback_payload, SlideMaster.CONTENT_BOX)
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

    # Emit telemetry to STDERR conforming to Observability rules
    sys.stderr.write(json.dumps(summary) + "\n")
    sys.stderr.flush()

    return prs, summary
