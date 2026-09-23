"""Public file-to-file conversion API."""
from __future__ import annotations

import sys
import warnings
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
    toc = config.get("table_of_contents", {})
    if toc.get("enabled", False) and toc.get("update_on_windows", False) and sys.platform == "win32":
        _update_word_fields(output)
    return output


def _update_word_fields(output: Path) -> None:
    """Use installed Microsoft Word to populate cached TOC and page-number fields."""
    try:
        import win32com.client
    except ImportError:
        warnings.warn(
            "Microsoft Word field update skipped: install pywin32 or update fields in Word with Ctrl+A, F9.",
            RuntimeWarning,
        )
        return

    word = document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(
            str(output.resolve()),
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
        )
        document.Fields.Update()
        for index in range(1, document.TablesOfContents.Count + 1):
            document.TablesOfContents.Item(index).Update()
        document.Save()
    except Exception as exc:
        warnings.warn(
            f"Microsoft Word could not update document fields ({exc}); update them in Word with Ctrl+A, F9.",
            RuntimeWarning,
        )
    finally:
        if document is not None:
            document.Close(SaveChanges=False)
        if word is not None:
            word.Quit()
