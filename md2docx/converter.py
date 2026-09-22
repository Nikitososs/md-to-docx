"""Markdown token rendering and Word document construction."""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from markdown_it import MarkdownIt
from markdown_it.token import Token

from .metadata import _SafeFormat, _frontmatter, _normalize_metadata

ALIGN = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
         "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}

def get_inline(tokens: list[Token] | None) -> list[Token]:
    return tokens or []


def token_text(tokens: list[Token] | None) -> str:
    out = []
    for token in get_inline(tokens):
        if token.type in ("text", "code_inline", "html_inline"):
            out.append(token.content)
        elif token.type == "image":
            out.append(token.content or "")
        elif token.type == "softbreak":
            out.append(" ")
    return "".join(out)


class Converter:
    def __init__(self, config: dict[str, Any], variables: dict[str, Any] | None = None):
        self.config = config
        self.variables = variables or {}
        self.doc = Document()
        self.counts = {"figure": 0, "table": 0, "listing": 0}
        self._in_bibliography = False
        self._configure_page()
        self._configure_styles()
        self.md = MarkdownIt("commonmark").enable("table")

    def _configure_page(self) -> None:
        page = self.config.get("page", {})
        for section in self.doc.sections:
            self._apply_page_settings(section, page)

    @staticmethod
    def _apply_page_settings(section, page: dict[str, Any]) -> None:
        section.page_width = Cm(float(page.get("width_cm", 21.0)))
        section.page_height = Cm(float(page.get("height_cm", 29.7)))
        for prop, key, default in (("top_margin", "top_cm", 2.0), ("bottom_margin", "bottom_cm", 2.0),
                                   ("left_margin", "left_cm", 3.0), ("right_margin", "right_cm", 1.5)):
            setattr(section, prop, Cm(float(page.get(key, default))))
        if "header_distance_cm" in page:
            section.header_distance = Cm(float(page["header_distance_cm"]))
        if "footer_distance_cm" in page:
            section.footer_distance = Cm(float(page["footer_distance_cm"]))

    def _configure_styles(self) -> None:
        body = self.config.get("body", {})
        style = self.doc.styles["Normal"]
        style.font.name = body.get("font", "Times New Roman")
        style.font.size = Pt(float(body.get("size_pt", 14)))
        style.paragraph_format.space_before = Pt(float(body.get("space_before_pt", 0)))
        style.paragraph_format.space_after = Pt(float(body.get("space_after_pt", 0)))
        style.paragraph_format.line_spacing = float(body.get("line_spacing", 1.2))
        style.paragraph_format.first_line_indent = Cm(float(body.get("first_line_indent_cm", 1.25)))
        self._set_east_asia(style._element.get_or_add_rPr(), body.get("font", "Times New Roman"))
        self._configure_heading_styles()
        self._configure_toc_styles()
        self._configure_page_numbers()

    def _configure_heading_styles(self) -> None:
        headings = self.config.get("headings", {})
        body = self.config.get("body", {})
        for level in range(1, 4):
            options = headings.get(str(level), {})
            try:
                style = self.doc.styles[f"Heading {level}"]
            except KeyError:
                continue
            style.font.name = body.get("font", "Times New Roman")
            style.font.size = Pt(float(options.get("size_pt", body.get("size_pt", 14))))
            style.font.bold = bool(options.get("bold", level <= 2))
            style.font.color.rgb = RGBColor.from_string(options.get("color_hex", "000000"))
            self._set_east_asia(style._element.get_or_add_rPr(), body.get("font", "Times New Roman"))
            style.paragraph_format.alignment = ALIGN.get(options.get("align", "left"), WD_ALIGN_PARAGRAPH.LEFT)
            style.paragraph_format.first_line_indent = Cm(float(options.get("first_line_indent_cm", 0)))
            style.paragraph_format.left_indent = Cm(float(options.get("left_indent_cm", 0)))
            style.paragraph_format.space_before = Pt(float(options.get("space_before_pt", 0)))
            style.paragraph_format.space_after = Pt(float(options.get("space_after_pt", 0)))
            style.paragraph_format.line_spacing = float(options.get("line_spacing", 1.5))
            style.paragraph_format.keep_with_next = bool(options.get("keep_with_next", True))

    def _configure_toc_styles(self) -> None:
        toc = self.config.get("table_of_contents", {})
        if not toc.get("enabled", False):
            return
        font = toc.get("font", self.config.get("body", {}).get("font", "Times New Roman"))
        size = float(toc.get("size_pt", self.config.get("body", {}).get("size_pt", 14)))
        indent_step = float(toc.get("indent_step_cm", 0.5))
        for level in range(1, int(toc.get("levels", 3)) + 1):
            try:
                style = self.doc.styles[f"TOC {level}"]
            except KeyError:
                style = self.doc.styles.add_style(f"TOC {level}", WD_STYLE_TYPE.PARAGRAPH)
            style.font.name = font
            style.font.size = Pt(size)
            style.font.color.rgb = RGBColor.from_string(toc.get("color_hex", "000000"))
            self._set_east_asia(style._element.get_or_add_rPr(), font)
            style.paragraph_format.left_indent = Cm(indent_step * (level - 1))
            style.paragraph_format.first_line_indent = Cm(0)
            style.paragraph_format.space_before = Pt(float(toc.get("space_before_pt", 0)))
            style.paragraph_format.space_after = Pt(float(toc.get("space_after_pt", 0)))
            style.paragraph_format.line_spacing = float(toc.get("line_spacing", 1.5))
        for name in ("Hyperlink", "FollowedHyperlink"):
            try:
                style = self.doc.styles[name]
            except KeyError:
                continue
            style.font.color.rgb = RGBColor.from_string(toc.get("color_hex", "000000"))
            style.font.underline = False

    def _configure_page_numbers(self) -> None:
        numbers = self.config.get("page_numbers", {})
        if not numbers.get("enabled", False):
            return
        for section in self.doc.sections:
            section.different_first_page_header_footer = bool(numbers.get("omit_first_page", True))
            paragraph = section.footer.paragraphs[0]
            paragraph.alignment = ALIGN.get(numbers.get("align", "center"), WD_ALIGN_PARAGRAPH.CENTER)
            paragraph.paragraph_format.first_line_indent = Cm(0)
            run = paragraph.add_run()
            run.font.name = self.config.get("body", {}).get("font", "Times New Roman")
            run.font.size = Pt(float(numbers.get("size_pt", 12)))
            begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
            instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = " PAGE "
            separate = OxmlElement("w:fldChar"); separate.set(qn("w:fldCharType"), "separate")
            cached = OxmlElement("w:t"); cached.text = "1"
            end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
            for node in (begin, instr, separate, cached, end):
                run._r.append(node)
        settings = self.doc.settings._element
        update = settings.find(qn("w:updateFields"))
        if update is None:
            update = OxmlElement("w:updateFields")
            settings.append(update)
        update.set(qn("w:val"), "true")

    @staticmethod
    def _set_east_asia(rpr, font: str) -> None:
        rfonts = rpr.rFonts
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
            rfonts.set(qn(f"w:{attr}"), font)

    def _format_paragraph(self, p, style: dict[str, Any] | None = None) -> None:
        style = style or {}
        fmt = p.paragraph_format
        fmt.space_before = Pt(float(style.get("space_before_pt", 0)))
        fmt.space_after = Pt(float(style.get("space_after_pt", 0)))
        fmt.line_spacing = float(style.get("line_spacing", self.config.get("body", {}).get("line_spacing", 1.2)))
        fmt.first_line_indent = Cm(float(style.get("first_line_indent_cm", self.config.get("body", {}).get("first_line_indent_cm", 1.25))))
        if "left_indent_cm" in style:
            fmt.left_indent = Cm(float(style["left_indent_cm"]))
        p.alignment = ALIGN.get(style.get("align", "justify"), WD_ALIGN_PARAGRAPH.JUSTIFY)
        if style.get("keep_with_next"):
            fmt.keep_with_next = True
        if style.get("keep_together"):
            fmt.keep_together = True
        if style.get("page_break_before"):
            fmt.page_break_before = True

    def _text(self, text: str, p, *, bold=False, italic=False, code=False, size=None, font=None,
              link_url=None, preserve_nbsp=False) -> None:
        if not code:
            formatting = self.config.get("formatting", {})
            if formatting.get("convert_em_dash_to_en_dash", True) and not self._in_bibliography:
                text = text.replace("—", "–")
            if formatting.get("normalize_quotes_to_russian", True) and not self._in_bibliography:
                text = re.sub(r'"([^"\n]*)"', r"«\1»", text)
                text = text.replace("“", "«").replace("”", "»")
            if not preserve_nbsp:
                text = text.replace("\u00a0", " ")
        # The department sample requests italics for Latin text; keep code exempt.
        use_latin_italic = bool(self.config.get("formatting", {}).get("italicize_latin", True)) and not code and not self._in_bibliography
        parts = re.split(r"([A-Za-z]+(?:[A-Za-z0-9_./:+#-]*[A-Za-z0-9_+#])?)", text) if use_latin_italic else [text]
        for part in filter(None, parts):
            r = p.add_run(part)
            is_latin = use_latin_italic and bool(re.fullmatch(r"[A-Za-z][A-Za-z0-9_./:+#-]*", part))
            r.bold = bold
            r.italic = italic or is_latin
            r.font.name = font or (self.config.get("code", {}).get("font", "DejaVu Sans") if code else self.config.get("body", {}).get("font", "Times New Roman"))
            if size is not None:
                r.font.size = Pt(float(size))
            self._set_east_asia(r._element.get_or_add_rPr(), r.font.name)
            if link_url:
                r.font.underline = True
                r.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
                rel_id = p.part.relate_to(link_url, RT.HYPERLINK, is_external=True)
                hyperlink = OxmlElement("w:hyperlink")
                hyperlink.set(qn("r:id"), rel_id)
                r._element.getparent().remove(r._element)
                hyperlink.append(r._element)
                p._p.append(hyperlink)

    def _inline(self, tokens: list[Token] | None, p, marks=()) -> None:
        active = list(marks)
        link_url = None
        for t in get_inline(tokens):
            if t.type in ("strong_open", "em_open", "s_open"):
                active.append(t.tag)
            elif t.type in ("strong_close", "em_close", "s_close"):
                if active and active[-1] == t.tag:
                    active.pop()
                elif t.tag in active:
                    active.remove(t.tag)
            elif t.type == "link_open":
                link_url = t.attrGet("href")
            elif t.type == "link_close":
                link_url = None
            elif t.type in ("text", "code_inline", "html_inline"):
                self._text(t.content, p, bold="strong" in active, italic="em" in active,
                           code=t.type == "code_inline", size=self.config.get("code", {}).get("size_pt") if t.type == "code_inline" else None,
                           link_url=link_url)
            elif t.type in ("softbreak", "hardbreak"):
                p.add_run().add_break()
            elif t.type == "image":
                self._text(t.content, p, link_url=link_url)

    def _new_text_paragraph(self, text_tokens: list[Token] | None, style=None, heading=False):
        p = self.doc.add_paragraph()
        self._format_paragraph(p, style)
        self._inline(text_tokens, p)
        return p

    def _caption(self, kind: str, title: str, align: str) -> None:
        self.counts[kind] += 1
        labels = self.config.get("captions", {}).get(kind, {})
        number = labels.get("start_number", 1) + self.counts[kind] - 1
        pattern = labels.get("pattern", {"figure": "Рисунок {n} – {title}", "table": "Таблица {n} – {title}", "listing": "Листинг {n} – {title}"}[kind])
        p = self.doc.add_paragraph()
        self._format_paragraph(p, {"align": align, "first_line_indent_cm": 0, "space_before_pt": 0,
                                   "space_after_pt": 0, "keep_with_next": kind != "figure"})
        self._text(pattern.format(n=number, title=title), p, bold=bool(labels.get("bold", False)), size=labels.get("size_pt"))

    def _image(self, token: Token, source_dir: Path, title: str | None = None) -> None:
        src = token.attrGet("src") or ""
        path = Path(src)
        if not path.is_absolute():
            path = source_dir / path
        if not path.is_file():
            raise ValueError(f"Image not found: {path}")
        self.doc.add_paragraph()  # required separation before figure
        p = self.doc.add_paragraph()
        self._format_paragraph(p, self.config.get("figure", {}))
        width = self.config.get("figure", {}).get("max_width_cm", 15)
        p.add_run().add_picture(str(path), width=Cm(float(width)))
        p.paragraph_format.keep_with_next = True
        self._caption("figure", title or token.content or path.stem, "center")
        self.doc.add_paragraph()

    def _make_table(self, rows: list[list[list[Token]]], title: str | None) -> None:
        if title:
            self.doc.add_paragraph()
            self._caption("table", title, self.config.get("table", {}).get("caption_align", "left"))
        if not rows:
            return
        width = max(map(len, rows))
        table = self.doc.add_table(rows=0, cols=width)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True
        style_name = self.config.get("table", {}).get("style", "Table Grid")
        try:
            table.style = style_name
        except KeyError:
            pass
        for ri, cells in enumerate(rows):
            out = table.add_row().cells
            for ci in range(width):
                cell = out[ci]
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                cell.text = ""
                p = cell.paragraphs[0]
                self._format_paragraph(p, {"align": "center" if ri == 0 else self.config.get("table", {}).get("cell_align", "left"), "first_line_indent_cm": 0})
                if ci < len(cells): self._inline(cells[ci], p)
            if ri == 0 and self.config.get("table", {}).get("repeat_header", True):
                trpr = table.rows[0]._tr.get_or_add_trPr()
                header = OxmlElement("w:tblHeader"); header.set(qn("w:val"), "true"); trpr.append(header)
        self.doc.add_paragraph()

    def _cover(self) -> None:
        cover = self.config.get("cover")
        if not cover:
            return
        for part in cover.get("parts", cover.get("paragraphs", [])):
            if part.get("type", "paragraph") == "table":
                rows = part.get("rows", [])
                if not rows:
                    continue
                ncols = max(len(row) for row in rows)
                table = self.doc.add_table(rows=0, cols=ncols)
                table.autofit = bool(part.get("autofit", False))
                widths = part.get("widths_cm", [])
                for col, width in enumerate(widths[:ncols]):
                    table.columns[col].width = Cm(float(width))
                try:
                    table.style = part.get("style", "Normal Table")
                except KeyError:
                    pass
                for _ in rows:
                    table.add_row()
                for row_index, start_col, end_col in part.get("merges", []):
                    table.cell(int(row_index), int(start_col)).merge(table.cell(int(row_index), int(end_col)))
                for row_index, rowdata in enumerate(rows):
                    populated_cells = set()
                    for index in range(ncols):
                        cell = table.cell(row_index, index)
                        cell_key = id(cell._tc)
                        if cell_key in populated_cells:
                            continue
                        populated_cells.add(cell_key)
                        cell.text = ""
                        p = cell.paragraphs[0]
                        cell_cfg = part.get("cell", {})
                        self._format_paragraph(p, {"align": cell_cfg.get("align", "left"), "first_line_indent_cm": 0,
                                                   "line_spacing": cell_cfg.get("line_spacing", 1.0)})
                        if index < len(rowdata):
                            self._text(str(rowdata[index]).format_map(_SafeFormat(self.variables)), p,
                                       bold=cell_cfg.get("bold", False), size=cell_cfg.get("size_pt", 12))
                continue
            value = part.get("text", "").format_map(_SafeFormat(self.variables))
            p = self.doc.add_paragraph()
            self._format_paragraph(p, {"align": part.get("align", "center"),
                                       "first_line_indent_cm": part.get("first_line_indent_cm", 0),
                                       "left_indent_cm": part.get("left_indent_cm", 0),
                                       "line_spacing": part.get("line_spacing", 1.0),
                                       "space_before_pt": part.get("space_before_pt", 0),
                                       "space_after_pt": part.get("space_after_pt", 0)})
            if "runs" in part:
                for item in part["runs"]:
                    if "field" in item:
                        field = item["field"]
                        run_text = str(self.variables.get(field, "{" + field + "}"))
                    else:
                        run_text = str(item.get("text", "")).format_map(_SafeFormat(self.variables))
                    start_run = len(p.runs)
                    self._text(run_text, p, bold=item.get("bold", part.get("bold", False)),
                               size=item.get("size_pt", part.get("size_pt", 12)),
                               preserve_nbsp=item.get("preserve_nbsp", False))
                    for run in p.runs[start_run:]:
                        run.underline = item.get("underline", False)
            else:
                self._text(value, p, bold=part.get("bold", False), size=part.get("size_pt", 12))
        if cover.get("start_body_section", False):
            if cover.get("page"):
                self._apply_page_settings(self.doc.sections[0], cover["page"])
            first_section = self.doc.sections[0]
            first_section.different_first_page_header_footer = True
            self._cover_header_footer(first_section.first_page_header, cover.get("header", []))
            self._cover_header_footer(first_section.first_page_footer, cover.get("footer", []))
            body_section = self.doc.add_section(WD_SECTION.NEW_PAGE)
            self._apply_page_settings(body_section, self.config.get("page", {}))
            # The cover header/footer belong only to the first page of section 1.
            # The report section uses its regular (normally blank/page-number) parts.
            body_section.different_first_page_header_footer = False
        elif cover.get("page_break_after", True):
            self.doc.add_page_break()

    def _cover_header_footer(self, container, parts: list[dict[str, Any]]) -> None:
        if not parts:
            return
        container.is_linked_to_previous = False
        paragraphs = container.paragraphs
        first = paragraphs[0]
        first.clear()
        for index, part in enumerate(parts):
            p = first if index == 0 else container.add_paragraph()
            self._format_paragraph(p, {"align": part.get("align", "center"),
                                       "first_line_indent_cm": 0,
                                       "left_indent_cm": part.get("left_indent_cm", 0),
                                       "line_spacing": part.get("line_spacing", 1.0),
                                       "space_before_pt": part.get("space_before_pt", 0),
                                       "space_after_pt": part.get("space_after_pt", 0)})
            value = part.get("text", "").format_map(_SafeFormat(self.variables))
            self._text(value, p, bold=part.get("bold", False), size=part.get("size_pt", 12))

    def convert(self, markdown: str, source_dir: Path) -> Document:
        data, markdown = _frontmatter(markdown)
        data = _normalize_metadata(data, self.config)
        merged = {**self.config.get("metadata_defaults", {}), **data, **self.variables}
        self.variables = merged
        self._cover()
        tokens = self.md.parse(markdown)
        i = 0
        pending_caption: tuple[str, str] | None = None
        in_bibliography = False
        in_toc = False
        toc_page_break_pending = False
        body_content_seen = False
        list_stack: list[dict[str, Any]] = []
        while i < len(tokens):
            t = tokens[i]
            if t.type == "heading_open":
                inline = tokens[i + 1]
                title = token_text(inline.children).strip()
                level = int(t.tag[1:])
                if in_toc:
                    in_toc = False
                    toc_page_break_pending = True
                norms = [s.casefold() for s in self.config.get("structure_headings", [])]
                if title.casefold() in norms:
                    p = self.doc.add_paragraph()
                    if title.casefold() != self.config.get("table_of_contents", {}).get("heading", "СОДЕРЖАНИЕ").casefold():
                        p.style = self.doc.styles["Heading 1"]
                    structural_style = dict(self.config.get("structural_heading", {"align":"center", "first_line_indent_cm":0}))
                    structural_page_break = structural_style.pop("page_break_before", False)
                    self._format_paragraph(p, structural_style)
                    p.paragraph_format.page_break_before = bool(structural_page_break and body_content_seen)
                    if toc_page_break_pending and self.config.get("table_of_contents", {}).get("page_break_before_next_heading", True):
                        p.paragraph_format.page_break_before = True
                    if toc_page_break_pending:
                        toc_page_break_pending = False
                    self._text(title.upper(), p, bold=self.config.get("structural_heading", {}).get("bold", True))
                    in_bibliography = title.casefold() in {x.casefold() for x in self.config.get("bibliography_headings", [self.config.get("bibliography_heading", "Список использованных источников")])}
                    body_content_seen = True
                else:
                    hconf = self.config.get("headings", {}).get(str(level), {})
                    p = self.doc.add_paragraph()
                    p.style = self.doc.styles[f"Heading {min(level, 3)}"]
                    self._format_paragraph(p, hconf)
                    if toc_page_break_pending and self.config.get("table_of_contents", {}).get("page_break_before_next_heading", True):
                        p.paragraph_format.page_break_before = True
                    if toc_page_break_pending:
                        toc_page_break_pending = False
                    self._inline(inline.children, p)
                    for r in p.runs:
                        if hconf.get("bold", level <= 2): r.bold = True
                    in_bibliography = False
                    body_content_seen = True
                if title.casefold() == self.config.get("table_of_contents", {}).get("heading", "СОДЕРЖАНИЕ").casefold():
                    toc = self.config.get("table_of_contents", {})
                    if toc.get("enabled", False):
                        self._add_table_of_contents(toc)
                        in_toc = True
                self._in_bibliography = in_bibliography
                i += 3; continue
            if in_toc:
                i += 1
                continue
            if t.type == "paragraph_open":
                inline = tokens[i + 1]
                imgs = [x for x in get_inline(inline.children) if x.type == "image"]
                raw = token_text(inline.children).strip()
                if raw.lower().startswith("таблица:"):
                    pending_caption = ("table", raw.split(":", 1)[1].strip()); i += 3; continue
                if raw.lower().startswith("листинг:"):
                    pending_caption = ("listing", raw.split(":", 1)[1].strip()); i += 3; continue
                if imgs:
                    before = []
                    for part in get_inline(inline.children):
                        if part.type == "image": self._image(part, source_dir, part.content or None)
                        elif part.type == "text" and part.content.strip(): before.append(part)
                    body_content_seen = True
                    i += 3; continue
                if raw.startswith("<pagebreak>"):
                    self.doc.add_page_break(); body_content_seen = True; i += 3; continue
                p = self.doc.add_paragraph()
                if in_bibliography:
                    bib = self.config.get("bibliography", {})
                    body = self.config.get("body", {})
                    self._format_paragraph(p, {
                        "left_indent_cm": float(bib.get("left_indent_cm", 0)),
                        "first_line_indent_cm": float(bib.get(
                            "first_line_indent_cm", body.get("first_line_indent_cm", 0)
                        )),
                    })
                else:
                    self._format_paragraph(p)
                self._in_bibliography = in_bibliography
                self._inline(inline.children, p)
                body_content_seen = True
                i += 3; continue
            if t.type in ("bullet_list_open", "ordered_list_open"):
                list_stack.append({"kind": "ordered" if t.type == "ordered_list_open" else "bullet", "counter": int(t.attrGet("start") or 1) - 1}); i += 1; continue
            if t.type in ("bullet_list_close", "ordered_list_close"):
                if list_stack: list_stack.pop()
                i += 1; continue
            if t.type == "list_item_open":
                j = i + 1
                while j < len(tokens) and tokens[j].type not in ("list_item_close", "list_item_open"):
                    if tokens[j].type == "inline":
                        frame = list_stack[-1] if list_stack else {"kind": "bullet", "counter": 0}
                        frame["counter"] += 1
                        p = self.doc.add_paragraph()
                        if in_bibliography:
                            bib = self.config.get("bibliography", {})
                            body = self.config.get("body", {})
                            self._format_paragraph(p, {
                                "first_line_indent_cm": float(bib.get(
                                    "first_line_indent_cm", body.get("first_line_indent_cm", 0)
                                )),
                                "left_indent_cm": float(bib.get("left_indent_cm", 0)),
                            })
                            marker = f"{frame['counter']}. " if bib.get("numbered", True) else ""
                        else:
                            self._format_paragraph(p, {"first_line_indent_cm": 0, "left_indent_cm": 0.75 * max(1, len(list_stack))})
                            marker = f"{frame['counter']}. " if frame["kind"] == "ordered" else "• "
                        self._in_bibliography = in_bibliography
                        self._text(marker, p)
                        self._inline(tokens[j].children, p)
                        body_content_seen = True
                        i = j + 1
                        break
                    j += 1
                else:
                    i = j
                continue
            if t.type == "fence":
                if pending_caption and pending_caption[0] == "listing": self._caption("listing", pending_caption[1], "left"); pending_caption = None
                p = self.doc.add_paragraph()
                self._format_paragraph(p, {"align":"left", "first_line_indent_cm":0, "left_indent_cm":float(self.config.get("code", {}).get("left_indent_cm", 0)),
                                           "line_spacing":self.config.get("code", {}).get("line_spacing", 1.0),
                                           "space_before_pt":self.config.get("code", {}).get("space_before_pt", 0), "space_after_pt":self.config.get("code", {}).get("space_after_pt", 0)})
                self._text(t.content.rstrip("\n"), p, code=True, size=self.config.get("code", {}).get("size_pt", 11), font=self.config.get("code", {}).get("font", "DejaVu Sans"))
                p.paragraph_format.keep_together = True
                body_content_seen = True
                i += 1; continue
            if t.type == "table_open":
                rows = []; i += 1; current = None
                while i < len(tokens) and tokens[i].type != "table_close":
                    if tokens[i].type in ("tr_open",): current = []; rows.append(current)
                    elif tokens[i].type == "inline" and current is not None: current.append(tokens[i].children or [])
                    i += 1
                title = pending_caption[1] if pending_caption and pending_caption[0] == "table" else None
                pending_caption = None
                self._make_table(rows, title); body_content_seen = True; i += 1; continue
            if t.type == "hr":
                self.doc.add_page_break(); body_content_seen = True; i += 1; continue
            i += 1
        return self.doc

    def _add_table_of_contents(self, toc: dict[str, Any]) -> None:
        """Insert a native Word TOC field; Word fills its entries and page numbers."""
        paragraph = self.doc.add_paragraph()
        self._format_paragraph(paragraph, {
            "align": "left",
            "first_line_indent_cm": 0,
            "line_spacing": toc.get("line_spacing", 1.5),
            "space_before_pt": toc.get("space_before_pt", 0),
            "space_after_pt": toc.get("space_after_pt", 0),
        })
        run = paragraph.add_run()
        run.font.name = toc.get("font", self.config.get("body", {}).get("font", "Times New Roman"))
        run.font.size = Pt(float(toc.get("size_pt", self.config.get("body", {}).get("size_pt", 14))))
        self._set_east_asia(run._element.get_or_add_rPr(), toc.get("font", "Times New Roman"))
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        switches = ["TOC", f'\\o "1-{int(toc.get("levels", 3))}"']
        if toc.get("hyperlinks", True):
            switches.append("\\h")
        if toc.get("hide_page_numbers_in_web_layout", True):
            switches.append("\\z")
        if toc.get("use_outline_levels", True):
            switches.append("\\u")
        instr.text = " " + " ".join(switches) + " "
        separate = OxmlElement("w:fldChar")
        separate.set(qn("w:fldCharType"), "separate")
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        for element in (begin, instr, separate, end):
            run._r.append(element)
        settings = self.doc.settings._element
        update = settings.find(qn("w:updateFields"))
        if update is None:
            update = OxmlElement("w:updateFields")
            settings.append(update)
        update.set(qn("w:val"), "true")
