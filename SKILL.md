---
name: md-to-docx-academic
description: Convert an academic Markdown report into DOCX with this repository's JSON style profiles. Use when preparing a Word document from structured Markdown and local images.
---

# Convert an academic Markdown report to DOCX

Use the repository converter and the Markdown source conventions below. Keep the author's meaning, numbers, formulas, and reference details unchanged. Do not invent missing student, reviewer, course, or publication data.

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
  "discipline": "Название дисциплины",
  "title": "Тема работы",
  "student": {"name": "Иванов Иван Иванович", "course": "3", "group": "Группа-001", "program": "00.00.00 Название направления"},
  "teacher": {"title": "доцент, кандидат наук", "name": "Петров Пётр Петрович"},
  "institution": {"short": "ОмГТУ", "department": "ПМиФИ", "year": "2026"},
  "work": {"type": "Домашнее задание"}
}
}}
```

The JSON block must be the first content in the file and must contain valid JSON between the `{{json` and `}}` lines. The parser reads explicit fields; it does not infer course, group, names, or titles from prose. `title` is the work's topic; `work.type` is its kind. Add `work.variant` only when the assignment has a variant number; otherwise the separate variant line is omitted. The FIT title page follows the IDZ template, including its first-page header with institution and department, its work/student/reviewer fields, underlined signature slots, and city/year footer. YAML and free-text metadata are not supported. Values in the optional `--vars` file override Markdown metadata. Never fill unknown placeholders with guesses.

Use these source rules:

- Separate paragraphs and block elements with a blank line.
- In profiles with `table_of_contents.enabled: true`, add `# СОДЕРЖАНИЕ` where the clickable table should appear. The converter inserts a native Word TOC from subsequent Markdown headings; do not write manual entries beneath that heading. On Windows with Microsoft Word installed, the FIT profile populates the TOC during conversion. If Word is unavailable or the TOC is blank, open the document in Word, select the document (`Ctrl+A`), and update fields (`F9`).
- Structural section names configured in the profile (such as `ВВЕДЕНИЕ`, `ЗАКЛЮЧЕНИЕ`, and `СПИСОК ЛИТЕРАТУРЫ`) are promoted to level 1 even if Markdown uses `##` or `###`; they render uppercase, without bold, and without numbering. In the FIT profile, if the first level 2 or 3 heading after the introduction appears before `ОСНОВНАЯ ЧАСТЬ`, insert that level 1 heading automatically. An explicit `# ОСНОВНАЯ ЧАСТЬ` prevents a duplicate.
- Use `#`, `##`, and `###` for heading levels. In the FIT profile, regular level 2 headings (`##`) are numbered automatically from `1.`; write `## Section` without a number, or provide a number to retain. Level 3 headings (`###`) still need an explicit number such as `1.1`.
- Heading 2 and 3 styles use their configured spacing; the converter does not insert an extra empty paragraph after them. Do not add blank lines to the DOCX expectation based on Markdown's block separation.
- Write a figure as a standalone paragraph: `![Caption text](relative/image.png)`. The alt text becomes the numbered figure caption below the image. The converter adds an empty paragraph before the image/caption block and after its caption.
- Immediately before a pipe table, write `Таблица: Caption text` in a paragraph of its own. The converter creates a real Word table and places a numbered caption above it, with an empty paragraph before the caption and after the table.
- Immediately before a fenced code block, write `Листинг: Caption text` in a paragraph of its own.
- Each source line in a fenced listing becomes its own Word paragraph. Long listings may continue across pages; the caption stays with the first code line. The converter does not keep the whole listing on one page.
- Write inline TeX math as `$...$` and display math as `$$...$$` on separate lines. These are the supported delimiters; `\(...\)` and `\[...\]` are not parsed. Both forms become native editable Word equations (OMML), not images. Display equations are centered by default; formulas inside backticks remain code. Equation numbering is not automatic. Invalid or unsupported TeX raises a conversion error so source is not silently printed as raw TeX.
- Put the bibliography beneath a configured structural heading as a Markdown list, one complete source description per item. The FIT GOST profile prints entries without list markers and applies the profile's paragraph indent. It preserves entry text and punctuation; prepare complete descriptions in the required citation standard yourself. The profile does not infer missing metadata or rewrite references into GOST form.
- Use regular Markdown for emphasis, inline code, links, and ordered/unordered lists. Body lists become native Word lists with editable numbering/bullets; ordered lists honor their starting number and nested lists are indented. Bibliography list behavior remains controlled by its profile. When `formatting.italicize_latin` is enabled, Latin-script words (including extended Unicode Latin letters) are italicized in prose and inline code; fenced code listings and bibliography entries are exempt. Image paths are relative to the source `.md` file.

Use `examples/report-template.md` as the short starter or `examples/full-report-template.md` as the complete syntax reference. Remove all instructional placeholders before delivery. A standalone `---` is treated as a page break by this converter.

## 4. Convert

Run from the repository root. For the department profile:

```powershell
python -m md2docx .\report.md --config .\configs\fit-academic.json --vars .\report-vars.json -o .\out\report.docx
```

Substitute the selected profile and the matching variables JSON as needed. A missing image is a conversion error; fix the source path or asset and rerun. Do not silently remove the image or its caption.

## 5. Review the result

Open the DOCX and verify title-page fields, headings, tables, captions, images, code blocks, bibliography entries, and page numbering. If a document renderer is available, render and inspect every page. Pagination can change with Word/LibreOffice versions and installed fonts; do not promise identical page breaks across systems.

## Supported source blocks and known limits

The converter handles paragraphs; headings up to level 3; configured structural headings; automatic insertion of `ОСНОВНАЯ ЧАСТЬ` after the introduction when needed; automatic clickable Word tables of contents when enabled by the profile; automatic level 2 numbering and uppercase-to-sentence-case conversion when enabled by the profile; inline bold/italic/code and links; native ordered and unordered Word lists with nested indentation; Markdown pipe tables; standalone images; fenced code; inline and display TeX math as native OMML equations; and the configured bibliography section. The Windows API attempts to refresh Word fields when enabled and Microsoft Word/`pywin32` are available; otherwise update fields in Word with `Ctrl+A`, `F9`. It does not process YAML metadata, arbitrary raw HTML, footnotes, automatic level 3 numbering, equation numbering, citations from reference keys, OCR, or automatic GOST metadata generation.
