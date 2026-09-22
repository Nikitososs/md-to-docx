"""Mirror the canonical converter and docs into the standalone plugin bundle."""
from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PLUGIN_SOURCE = ROOT / "plugins" / "md-to-docx-academic"
MIRRORED_DIRS = ("md2docx", "configs", "examples")
MIRRORED_FILES = {
    Path("README.md"): ROOT / "README.md",
    Path("pyproject.toml"): PLUGIN_SOURCE / "pyproject.toml",
    Path("scripts/convert.py"): PLUGIN_SOURCE / "scripts" / "convert.py",
    Path("skills/md-to-docx-academic/SKILL.md"): PLUGIN_SOURCE / "skills" / "md-to-docx-academic" / "SKILL.md",
    Path(".codex-plugin/plugin.json"): PLUGIN_SOURCE / ".codex-plugin" / "plugin.json",
}
IGNORED_PARTS = {"__pycache__", ".pytest_cache", ".mypy_cache"}


def source_map() -> dict[Path, Path]:
    files: dict[Path, Path] = {}
    for directory in MIRRORED_DIRS:
        for source in (ROOT / directory).rglob("*"):
            if not source.is_file() or IGNORED_PARTS.intersection(source.parts) or ".egg-info" in source.parts:
                continue
            if source.suffix == ".pyc":
                continue
            files[source.relative_to(ROOT)] = source
    files.update({relative: source for relative, source in MIRRORED_FILES.items()})
    return files


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

    for relative, source in sorted(files.items(), key=lambda item: str(item[0]).casefold()):
        target = destination / relative
        if not target.is_file() or source.read_bytes() != target.read_bytes():
            differences.append(relative)
            if not args.check:
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source, target)

    if args.check:
        if differences:
            print("Plugin bundle is out of sync:")
            for relative in differences:
                print(f"  {relative}")
            return 1
        print(f"Plugin bundle is in sync: {destination}")
        return 0

    print(f"Synchronized {len(files)} files to {destination}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
