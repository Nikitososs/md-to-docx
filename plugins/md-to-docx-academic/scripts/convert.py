"""Run the bundled md2docx converter without requiring an editable install."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

PLUGIN_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PLUGIN_ROOT))

from md2docx.api import convert_file  # noqa: E402

PROFILES = {
    "fit-academic": "fit-academic.json",
    "fit-gost-bibliography": "fit-gost-bibliography-7.0.100-2018.json",
    "gost-report": "gost-report-7.32-2017.json",
}


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert Markdown to DOCX using the bundled academic profiles.")
    parser.add_argument("input", type=Path, help="Markdown source")
    parser.add_argument("-o", "--output", type=Path, help="Output DOCX (defaults to input stem.docx)")
    choice = parser.add_mutually_exclusive_group()
    choice.add_argument("--profile", choices=sorted(PROFILES), default="fit-academic", help="Bundled JSON profile")
    choice.add_argument("--config", type=Path, help="Path to a custom JSON profile")
    parser.add_argument("--vars", type=Path, help="Optional JSON cover variables")
    args = parser.parse_args()
    config = args.config or (PLUGIN_ROOT / "configs" / PROFILES[args.profile])
    output = args.output or args.input.with_suffix(".docx")
    try:
        convert_file(args.input, output, config, args.vars)
    except (OSError, ValueError, KeyError) as exc:
        print(f"md2docx: error: {exc}", file=sys.stderr)
        return 2
    print(f"Created {output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
