"""Public file-to-file conversion API."""
from __future__ import annotations

from pathlib import Path

from .config import load_json
from .converter import Converter


def convert_file(input_path: str | Path, output_path: str | Path, config_path: str | Path,
                 variables_path: str | Path | None = None) -> Path:
    source_path = Path(input_path)
    output = Path(output_path)
    config = load_json(Path(config_path))
    variables = load_json(Path(variables_path)) if variables_path else {}
    source = source_path.read_text(encoding="utf-8-sig")
    document = Converter(config, variables).convert(source, source_path.resolve().parent)
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)
    return output
