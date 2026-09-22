"""Loading and merging JSON style profiles."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

def _merge_config(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    merged = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key] = _merge_config(merged[key], value)
        else:
            merged[key] = value
    return merged


def load_json(path: Path, _seen: set[Path] | None = None) -> dict[str, Any]:
    path = path.resolve()
    seen = _seen or set()
    if path in seen:
        raise ValueError(f"Circular config inheritance: {path}")
    seen.add(path)
    try:
        value = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Cannot read JSON file {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise ValueError(f"Expected a JSON object in {path}")
    parent = value.pop("extends", None)
    if parent:
        base = load_json((path.parent / parent), seen)
        return _merge_config(base, value)
    return value
