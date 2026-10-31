"""Command Line Interface entry point for Slide Engine.

Handles standard input/output streaming, arguments parsing,
and proper operating system exit codes.
"""

from __future__ import annotations

import argparse
import io
import sys

from slide_engine.controller import generate_presentation
from slide_engine.models import PresentationDeck


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

    args = parser.parse_args()

    # 1. Read input JSON payload
    try:
        if args.input_file:
            with open(args.input_file, encoding="utf-8") as f:
                json_str = f.read()
        else:
            json_str = sys.stdin.read()

        deck = PresentationDeck.model_validate_json(json_str)
    except Exception as exc:
        sys.stderr.write(f"Schema Validation Error: {exc}\n")
        sys.exit(1)

    # 2. Generate presentation
    try:
        prs, _ = generate_presentation(deck)
    except Exception as exc:
        sys.stderr.write(f"Fatal Engine Failure: {exc}\n")
        sys.exit(2)

    # 3. Output PowerPoint binary
    try:
        if args.output_file:
            prs.save(args.output_file)
        else:
            # Memory stream to STDOUT
            buffer = io.BytesIO()
            prs.save(buffer)
            buffer.seek(0)
            sys.stdout.buffer.write(buffer.read())
            sys.stdout.buffer.flush()
    except Exception as exc:
        sys.stderr.write(f"Output I/O Error: {exc}\n")
        sys.exit(3)

    # Clean exit
    sys.exit(0)


if __name__ == "__main__":
    main()