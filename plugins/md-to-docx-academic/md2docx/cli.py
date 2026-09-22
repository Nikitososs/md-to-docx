"""Command-line interface for md2docx."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .api import convert_file


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="md2docx", description="Convert structured Markdown to DOCX using a JSON style profile")
    parser.add_argument("input", type=Path, help="Markdown source")
    parser.add_argument("-o", "--output", type=Path, help="Output DOCX path (default: input stem.docx)")
    parser.add_argument("-c", "--config", type=Path, required=True, help="JSON style profile")
    parser.add_argument("--vars", type=Path, help="Optional JSON object with cover fields and substitutions")
    args = parser.parse_args(argv)
    output = args.output or args.input.with_suffix(".docx")
    try:
        convert_file(args.input, output, args.config, args.vars)
        print(f"Created {output.resolve()}")
        return 0
    except (OSError, ValueError, KeyError) as exc:
        print(f"md2docx: error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
