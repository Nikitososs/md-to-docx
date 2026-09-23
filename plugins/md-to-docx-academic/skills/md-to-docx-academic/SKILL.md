---
name: md-to-docx-academic
description: Convert academic Markdown reports into DOCX using this local plugin's bundled converter and JSON style profiles.
---

# Convert an academic Markdown report to DOCX

Use the converter bundled with this plugin at `<USER_HOME>\plugins\md-to-docx-academic\scripts\convert.py`. Python 3.10+ and the declared dependencies (`markdown-it-py`, `python-docx`, `Pillow`) must be available. Follow these Markdown source conventions. Keep the author's meaning, numbers, formulas, and reference details unchanged. Do not invent missing student, reviewer, course, or publication data.

## 1. Inspect the inputs

- Read the full Markdown file and any JSON front matter at the top.
- Resolve each image path relative to the Markdown file; confirm the file exists.
- Check that each table and code block has the expected caption line immediately before it.
- Identify the bibliography heading and list items.
- Read the assignment/template instructions to choose the right profile. Do not infer a GOST standard just because the task is academic.

## 2. Choose one profile

- `configs/fit-academic.json`: supplied ФИТ requirements and title-page structure.
- `configs/fit-gost-bibliography-7.0.100-2018.json`: ФИТ body/title layout and a separate bibliography section profile for ГОСТ Р 7.0.100-2018.
- `configs/gost-report-7.32-2017.json`: research-report layout based on ГОСТ 7.32-2017. Use only when that standard applies to the assignment.
- A user-provided profile: use when the user supplies a department, university, or course configuration.

When the appropriate standard or local requirement is unclear, use the supplied department profile and state the unresolved choice. The bibliography profile does not automatically rewrite raw references into complete GOST descriptions.

## 3. Prepare the Markdown source

At the very top of the Markdown source, use a fixed JSON metadata block so the cover fills automatically:

```markdown
{{json
{
  "discipline": "Системы управления базами данных (СУБД PostgreSQL 15)",
  "title": "Полнотекстовый поиск и ранжирование релевантности",
  "student": {"name": "Иванов Иван Иванович", "course": "3", "group": "Группа-001", "program": "00.00.00 Название направления"},
  "teacher": {"title": "доц., к.н.", "name": "Петрова Анна Сергеевна"},
  "institution": {"short": "ОмГТУ", "department": "ПМиФИ", "year": "2026"},
  "work": {"type": "Домашнее задание", "variant": "1"}
}
}}
```

The JSON block must be the first content in the file and must contain valid JSON between the `{{json` and `}}` lines. The parser reads explicit fields; it does not infer course, group, names, or titles from prose. The FIT title page follows the IDZ template, including its first-page header with institution and department, its work/student/reviewer fields, underlined signature slots, and city/year footer. YAML and free-text metadata are not supported. Values in the optional `--vars` file override Markdown metadata. Never fill unknown placeholders with guesses.

Use these source rules:

- Separate paragraphs and block elements with a blank line.
- In profiles with `table_of_contents.enabled: true`, add `# СОДЕРЖАНИЕ` where the clickable table should appear. The converter inserts a native Word TOC from subsequent Markdown headings; do not write manual entries beneath that heading. On Windows with Microsoft Word installed, the FIT profile populates the TOC during conversion. If Word is unavailable or the TOC is blank, open the document in Word, select the document (`Ctrl+A`), and update fields (`F9`).
- Structural section headings use level 1 (`#`): the FIT profile renders them uppercase and starts each section on a new page. Level 2 (`##`) uses bold, does not start a new page, and the FIT profile numbers headings from `1.` automatically. Existing heading numbers are retained; headings written entirely in uppercase are converted to sentence case.
- Use `#`, `##`, and `###` for heading levels. In the FIT profile, level 2 headings (`##`) are numbered automatically from `1.`; write `## Section` without a number, or provide a number to retain. Level 3 headings (`###`) still need an explicit number such as `1.1`.
- Write a figure as a standalone paragraph: `![Caption text](relative/image.png)`. The alt text becomes the numbered figure caption below the image.
- Immediately before a pipe table, write `Таблица: Caption text` in a paragraph of its own. The converter creates a real Word table and places a numbered caption above it.
- Immediately before a fenced code block, write `Листинг: Caption text` in a paragraph of its own.
- Put the bibliography beneath a configured structural heading as a Markdown list, one complete source description per item. The FIT GOST profile prints entries without list markers and applies the profile's paragraph indent. Keep the source text in the required citation standard yourself; the converter does not infer missing metadata or rewrite it.
- Use regular Markdown for emphasis, inline code, links, and ordered/unordered lists. Image paths are relative to the source `.md` file.

Use the bundled `examples/report-template.md` as the short starter or `examples/full-report-template.md` as the complete syntax reference. Remove all instructional placeholders before delivery. A standalone `---` is treated as a page break by this converter.

## 4. Convert

Run the bundled wrapper from PowerShell. `fit-academic` is the default profile; use `fit-gost-bibliography` or `gost-report` when appropriate. For a user-provided profile, pass its path with `--config`.

```powershell
python "$env:USERPROFILE\plugins\md-to-docx-academic\scripts\convert.py" "C:\path\to\report.md" --profile fit-academic --vars "C:\path\to\report-vars.json" -o "C:\path\to\report.docx"
```

Omit `--vars` when the Markdown metadata contains the required values. The wrapper accepts `--config <path>` for a custom JSON profile. A missing image is a conversion error; fix the source path or asset and rerun. Do not silently remove the image or its caption.

## 5. Review the result

Open the DOCX and verify title-page fields, headings, tables, captions, images, code blocks, bibliography entries, and page numbering. If a document renderer is available, render and inspect every page. Pagination can change with Word/LibreOffice versions and installed fonts; do not promise identical page breaks across systems.

## Supported source blocks and known limits

The converter handles paragraphs; headings up to level 3; automatic clickable Word tables of contents when enabled by the profile; automatic level 2 numbering and uppercase-to-sentence-case conversion when enabled by the profile; inline bold/italic/code and links; ordered and unordered lists; Markdown pipe tables; standalone images; fenced code; and the configured bibliography section. It does not process YAML metadata, arbitrary raw HTML, footnotes, automatic level 3 numbering, citations from reference keys, OCR, or automatic GOST metadata generation.
