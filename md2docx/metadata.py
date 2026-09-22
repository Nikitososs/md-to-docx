"""Parsing fixed JSON document metadata and mapping it to cover fields."""
from __future__ import annotations

import json
from typing import Any

class _SafeFormat(dict):
    def __missing__(self, key): return "{" + key + "}"


def _frontmatter(source: str) -> tuple[dict[str, Any], str]:
    """Read the optional fixed JSON header; ordinary prose is never guessed as metadata."""
    lines = source.splitlines(keepends=True)
    if not lines or lines[0].strip() != "{{json":
        return {}, source
    close = next((i for i in range(1, len(lines)) if lines[i].strip() == "}}"), None)
    if close is None:
        raise ValueError("JSON metadata block starts with {{json but has no closing }} line")
    payload = "".join(lines[1:close])
    try:
        metadata = json.loads(payload)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid JSON metadata block: {exc}") from exc
    if not isinstance(metadata, dict):
        raise ValueError("JSON metadata block must contain an object")
    return metadata, "".join(lines[close + 1:])


def _normalize_metadata(metadata: dict[str, Any], source: dict[str, Any]) -> dict[str, Any]:
    """Map documented nested JSON fields to the flat cover placeholders."""
    data = {key: value for key, value in metadata.items()
            if key not in {"student", "teacher", "instructor", "supervisor", "institution", "work"}}
    student = metadata.get("student", {})
    if isinstance(student, dict):
        if "name" in student: data["student_name"] = student["name"]
        if "course" in student: data["course"] = student["course"]
        if "group" in student: data["group"] = student["group"]
    teacher = metadata.get("teacher", metadata.get("instructor", metadata.get("supervisor", {})))
    if isinstance(teacher, dict):
        if "title" in teacher: data["supervisor_title"] = teacher["title"]
        if "name" in teacher: data["supervisor_name"] = teacher["name"]
    institution = metadata.get("institution", {})
    if isinstance(institution, dict):
        short = institution.get("short")
        aliases = source.get("institution_aliases", {})
        if short:
            data["university_short"] = short
            data.update(aliases.get(str(short).casefold(), {}))
            data.setdefault("organization_unit", data.get("university_full", short))
        for field, target in (("full_name", "university_full"), ("city", "city"),
                              ("department", "department"), ("year", "year"),
                              ("organization", "organization"), ("organization_unit", "organization_unit")):
            if field in institution:
                data[target] = institution[field]
    work = metadata.get("work", {})
    if isinstance(work, dict):
        kind = str(work.get("type", "")).strip()
        number = str(work.get("number", "")).strip()
        if kind:
            data["work_type"] = kind if not number else f"{kind} № {number}"
    return data
