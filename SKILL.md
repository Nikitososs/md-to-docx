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
  "discipline": "Системы управления базами данных (СУБД PostgreSQL 15)",
  "title": "Полнотекстовый поиск и ранжирование релевантности",
  "student": {"name": "Иванов Иван Иванович", "course": "3", "group": "Группа-001"},
  "teacher": {"title": "доц., к.н.", "name": "Петрова Анна Сергеевна"},
  "institution": {"short": "ОмГТУ", "department": "ПМиФИ", "year": "2026"},
  "work": {"type": "Лабораторная работа", "number": "1"}
}
}}
```

The JSON block must be the first content in the file and must contain valid JSON between the `{{json` and `}}` lines. The parser reads explicit fields; it does not infer course, group, names, or titles from prose. The configured ОмГТУ alias supplies the full university name and city. Other institutions need an alias in the JSON profile or explicit values in `--vars`. YAML and free-text metadata are not supported. Values in the optional `--vars` file override Markdown metadata. For sample title-page data, see `examples/report-vars.json` and `examples/gost-report-vars.json`. Never fill unknown placeholders with guesses.

Use these source rules:

- Separate paragraphs and block elements with a blank line.
- Use `#`, `##`, and `###` for heading levels. Write section numbers into heading text, for example `## 1 Section` and `### 1.1 Subsection`; the converter does not generate heading numbering.
- Write a figure as a standalone paragraph: `![Caption text](relative/image.png)`. The alt text becomes the numbered figure caption below the image.
- Immediately before a pipe table, write `Таблица: Caption text` in a paragraph of its own. The converter creates a real Word table and places a numbered caption above it.
- Immediately before a fenced code block, write `Листинг: Caption text` in a paragraph of its own.
- Put the bibliography beneath a configured structural heading as a numbered Markdown list, one complete source description per list item. Keep the source text in the required citation standard yourself; the converter does not infer missing metadata or rewrite it.
- Use regular Markdown for emphasis, inline code, links, and ordered/unordered lists. Image paths are relative to the source `.md` file.

Use `examples/report-template.md` as the short starter or `examples/full-report-template.md` as the complete syntax reference. Remove all instructional placeholders before delivery. A standalone `---` is treated as a page break by this converter.

## 4. Convert

Run from the repository root. For the department profile:

```powershell
md2docx .\report.md --config .\configs\fit-academic.json --vars .\report-vars.json -o .\out\report.docx
```

Substitute the selected profile and the matching variables JSON as needed. A missing image is a conversion error; fix the source path or asset and rerun. Do not silently remove the image or its caption.

## 5. Review the result

Open the DOCX and verify title-page fields, headings, tables, captions, images, code blocks, bibliography entries, and page numbering. If a document renderer is available, render and inspect every page. Pagination can change with Word/LibreOffice versions and installed fonts; do not promise identical page breaks across systems.

## Supported source blocks and known limits

The converter handles paragraphs; headings up to level 3; inline bold/italic/code and links; ordered and unordered lists; Markdown pipe tables; standalone images; fenced code; and the configured bibliography section. It does not process YAML metadata, arbitrary raw HTML, footnotes, automatic section numbering, citations from reference keys, OCR, or automatic GOST metadata generation.
