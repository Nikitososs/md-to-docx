"""Backward-compatible imports for the original single-module API.

New code should import from :mod:`md2docx.api`, :mod:`md2docx.converter`,
:mod:`md2docx.config`, or :mod:`md2docx.metadata`.
"""
from .api import convert_file
from .cli import main
from .config import load_json
from .converter import Converter
from .metadata import _SafeFormat, _frontmatter, _normalize_metadata

__all__ = ["Converter", "convert_file", "load_json", "main"]
