"""Tests for cover_letter_utility.core."""
from __future__ import annotations

import os

import pytest
from docx import Document

from cover_letter_utility.core import (
    ensure_unique_slug,
    process_csv,
    read_csv_to_dicts,
)

EXAMPLES = os.path.join(os.path.dirname(__file__), "..", "examples")
EXAMPLE_TEMPLATE = os.path.join(EXAMPLES, "example_template.docx")


def write_csv(path, rows, header="slug,institution_name"):
    path.write_text(header + "\n" + "\n".join(rows) + "\n", encoding="utf-8")
    return str(path)


# ---- read_csv_to_dicts ----------------------------------------------------

def test_reads_rows_and_strips_whitespace(tmp_path):
    csv_path = write_csv(tmp_path / "apps.csv", ["  a  ,  Alpha College  "])
    rows = read_csv_to_dicts(csv_path)
    assert rows == [{"slug": "a", "institution_name": "Alpha College"}]


def test_duplicate_slugs_get_suffix_when_unique(tmp_path):
    csv_path = write_csv(tmp_path / "apps.csv", ["mit,MIT", "mit,MIT Again"])
    rows = read_csv_to_dicts(csv_path)
    assert [r["slug"] for r in rows] == ["mit", "mit_2"]


def test_duplicate_slugs_raise_in_strict_mode(tmp_path):
    csv_path = write_csv(tmp_path / "apps.csv", ["mit,MIT", "mit,MIT Again"])
    with pytest.raises(ValueError, match="Duplicate slug"):
        read_csv_to_dicts(csv_path, make_unique=False)


def test_blank_slug_raises(tmp_path):
    csv_path = write_csv(tmp_path / "apps.csv", [",Nameless College"])
    with pytest.raises(ValueError, match="Missing slug in row 2"):
        read_csv_to_dicts(csv_path)


# ---- ensure_unique_slug ---------------------------------------------------

def test_unique_slug_returns_base_when_free(tmp_path):
    assert ensure_unique_slug(str(tmp_path), "mit") == "mit"


def test_unique_slug_increments_past_existing_folders(tmp_path):
    (tmp_path / "mit").mkdir()
    (tmp_path / "mit_2").mkdir()
    assert ensure_unique_slug(str(tmp_path), "mit") == "mit_3"


# ---- process_csv (full pipeline) ------------------------------------------

def test_process_csv_fills_every_placeholder(tmp_path):
    out = tmp_path / "out"
    csv_path = os.path.join(EXAMPLES, "example_apps.csv")
    written = process_csv(EXAMPLE_TEMPLATE, csv_path, str(out))

    assert len(written) == 3
    for slug in written:
        letter = out / slug / f"cover_letter_{slug}.docx"
        assert letter.is_file()
        text = "\n".join(p.text for p in Document(str(letter)).paragraphs)
        assert "${" not in text  # no placeholder left unreplaced


def test_process_csv_does_not_overwrite_on_rerun(tmp_path):
    out = tmp_path / "out"
    csv_path = os.path.join(EXAMPLES, "example_apps.csv")

    first = process_csv(EXAMPLE_TEMPLATE, csv_path, str(out))
    second = process_csv(EXAMPLE_TEMPLATE, csv_path, str(out))

    # second run must not reuse any first-run folder
    assert set(first).isdisjoint(second)
    assert all(s.endswith("_2") for s in second)
