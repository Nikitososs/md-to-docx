"""Public file-to-file conversion API."""
from __future__ import annotations

import sys
import warnings
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

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
        _normalize_list_indents(output, config)
    return output


def _normalize_list_indents(output: Path, config: dict) -> None:
    """Restore profile list indents after Word updates TOC fields.

    Word may replace direct paragraph indents with a hanging indent inherited
    from the built-in list style while saving updated fields. Reapply the
    profile values after that save so every list item's text starts at the
    configured column in the delivered DOCX.
    """
    options = config.get("lists", {})
    document = Document(output)
    base_left = float(options.get("left_indent_cm", 0.75))
    nested_step = float(options.get("nested_indent_cm", 0.5))
    hanging = float(options.get("hanging_indent_cm", 0.5))
    first_line = options.get("first_line_indent_cm")

    for paragraph in document.paragraphs:
        ppr = paragraph._p.pPr
        num_pr = ppr.numPr if ppr is not None else None
        if num_pr is None:
            continue
        level = int(num_pr.ilvl.val) if num_pr.ilvl is not None else 0
        left = int(Cm(base_left + level * nested_step).twips)
        indent = ppr.find(qn("w:ind"))
        if indent is None:
            indent = OxmlElement("w:ind")
            ppr.append(indent)
        if first_line is not None:
            indent.set(qn("w:left"), "0")
            indent.set(qn("w:firstLine"), str(int(Cm(float(first_line)).twips)))
            indent.attrib.pop(qn("w:hanging"), None)
        else:
            indent.set(qn("w:left"), str(left))
        if first_line is None and hanging:
            indent.set(qn("w:hanging"), str(int(Cm(hanging).twips)))
            indent.attrib.pop(qn("w:firstLine"), None)
        elif first_line is None:
            indent.attrib.pop(qn("w:hanging"), None)
            indent.set(qn("w:firstLine"), "0")
            indent.set(qn("w:firstLine"), "0")

    document.save(output)


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
