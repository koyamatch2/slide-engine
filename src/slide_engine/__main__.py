"""Command Line Interface entry point for Slide Engine.

Handles standard input/output streaming, arguments parsing,
and proper operating system exit codes.
"""

from __future__ import annotations

import argparse
import io
import json
import sys
from typing import TYPE_CHECKING

from slide_engine.controller import generate_presentation
from slide_engine.models import PresentationDeck

if TYPE_CHECKING:
    from pptx.presentation import Presentation as PresentationType


def _load_deck(input_file: str | None) -> PresentationDeck:
    """Read and validate the PresentationDeck from file or stdin."""
    try:
        if input_file:
            with open(input_file, encoding="utf-8") as f:
                json_str = f.read()
        else:
            json_str = sys.stdin.read()
        return PresentationDeck.model_validate_json(json_str)
    except Exception as exc:
        sys.stderr.write(f"Schema Validation Error: {exc}\n")
        sys.exit(1)


def _save_presentation(prs: PresentationType, output_file: str | None) -> None:
    """Save the Presentation to a file or stream to stdout buffer."""
    try:
        if output_file:
            prs.save(output_file)
        else:
            buffer = io.BytesIO()
            prs.save(buffer)
            buffer.seek(0)
            sys.stdout.buffer.write(buffer.read())
            sys.stdout.buffer.flush()
    except Exception as exc:
        sys.stderr.write(f"Output I/O Error: {exc}\n")
        sys.exit(3)


def main() -> None:
    """Execute slide-engine CLI pipeline."""
    parser = argparse.ArgumentParser(
        description="Slide Engine - Production automated PowerPoint generation CLI",
    )
    parser.add_argument(
        "-i", "--input",
        dest="input_file",
        help="Path to input JSON specification file (defaults to STDIN)",
    )
    parser.add_argument(
        "-o", "--output",
        dest="output_file",
        help="Path to output .pptx file (defaults to STDOUT)",
    )
    parser.add_argument(
        "--fail-on-fallback",
        action="store_true",
        help="Exit with non-zero exit code if any slide triggers fallback (CORE-00)",
    )

    args = parser.parse_args()

    # 1. Read input JSON payload
    deck = _load_deck(args.input_file)

    # 2. Generate presentation
    try:
        prs, summary = generate_presentation(deck)
    except Exception as exc:
        sys.stderr.write(f"Fatal Engine Failure: {exc}\n")
        sys.exit(2)

    # Emit telemetry to STDERR conforming to Observability rules
    sys.stderr.write(json.dumps(summary) + "\n")
    sys.stderr.flush()

    # Validate strict generation constraint
    if args.fail_on_fallback and summary.get("fallback_slides", 0) > 0:
        fallback_count = summary["fallback_slides"]
        sys.stderr.write(
            f"Strict Generation Failure: {fallback_count} slide(s) fell back to error placeholder.\n"
        )
        sys.exit(4)

    # 3. Output PowerPoint binary
    _save_presentation(prs, args.output_file)
    sys.exit(0)



if __name__ == "__main__":
    main()
