"""Configurable academic Markdown to DOCX converter."""

from .api import convert_file
from .converter import Converter

__all__ = ["Converter", "convert_file"]
