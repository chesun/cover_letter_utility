"""Batch-customize .docx cover letters from a CSV of applications.

For each row in the CSV, placeholder fields in a Word template (written as
``${field_name}``) are replaced with that row's values, and the result is
saved as its own letter in a per-application folder.
"""
from cover_letter_utility.core import (
    customize_cover_letter,
    ensure_unique_slug,
    process_csv,
    read_csv_to_dicts,
)

__version__ = "0.1.0"

__all__ = [
    'customize_cover_letter',
    'read_csv_to_dicts',
    'ensure_unique_slug',
    'process_csv',
    '__version__',
]
