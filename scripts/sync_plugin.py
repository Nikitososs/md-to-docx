"""Sync only the thin plugin manifest and skill; library code stays in the repository root."""
from __future__ import annotations

import argparse
import shutil
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PLUGIN_SOURCE = ROOT / "plugins" / "md-to-docx-academic"
MIRRORED_FILES = {
    Path(".codex-plugin/plugin.json"): PLUGIN_SOURCE / ".codex-plugin" / "plugin.json",
    Path("skills/md-to-docx-academic/SKILL.md"): PLUGIN_SOURCE / "skills" / "md-to-docx-academic" / "SKILL.md",
}
FORBIDDEN_COPIES = (
    Path("configs"),
    Path("examples"),
    Path("md2docx"),
    Path("scripts"),
    Path("README.md"),
    Path("pyproject.toml"),
    Path("md2docx_academic.egg-info"),
)


def source_map() -> dict[Path, Path]:
    return MIRRORED_FILES.copy()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--destination",
        type=Path,
        default=PLUGIN_SOURCE,
        help="Plugin directory to update (defaults to the repository bundle).",
    )
    parser.add_argument("--check", action="store_true", help="Report differences without copying files.")
    args = parser.parse_args()
    destination = args.destination.expanduser().resolve()
    files = source_map()
    differences: list[Path] = []

    for relative in FORBIDDEN_COPIES:
        target = destination / relative
        if target.exists():
            differences.append(relative)
            if not args.check:
                if target.is_dir():
                    shutil.rmtree(target)
                else:
                    target.unlink()

    for relative, source in sorted(files.items(), key=lambda item: str(item[0]).casefold()):
        target = destination / relative
        if not target.is_file() or source.read_bytes() != target.read_bytes():
            differences.append(relative)
            if not args.check:
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source, target)

    if args.check:
        if differences:
            print("Plugin bundle needs sync (stale duplicate files or out-of-date skill/manifest):")
            for relative in differences:
                print(f"  {relative}")
            return 1
        print(f"Plugin bundle is in sync: {destination}")
        return 0

    print(f"Synchronized the skill and manifest; removed {len([p for p in differences if p in FORBIDDEN_COPIES])} duplicate paths: {destination}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
