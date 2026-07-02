"""Batch-customize .docx cover letters from a CSV of applications.

For each row in the CSV, the placeholder fields in a Word template
(written as ``${field_name}``) are replaced with that row's values, and
the result is saved as its own letter in a per-application folder.

Run as a CLI::

    python utility.py --template letter_template.docx --csv apps.csv --out ./applications

or import :func:`process_csv` / :func:`customize_cover_letter` as a library.
"""
from __future__ import annotations

import argparse
import csv
import os

from python_docx_replace import docx_replace
from docx import Document

__all__ = [
    'customize_cover_letter',
    'read_csv_to_dicts',
    'ensure_unique_slug',
    'process_csv',
]

DEFAULT_SLUG_FIELD = 'slug'


def customize_cover_letter(
        template_name: str,
        template_path: str,
        app_path: str,
        slug: str,
        replace_dict: dict,
) -> str:
    """Replace ``${key}`` fields in a template and save one letter.

    Creates a folder named ``slug`` inside ``app_path`` (uniquified if it
    already exists) and writes ``cover_letter_<slug>.docx`` into it.

    Args:
        template_name: File name of the .docx template.
        template_path: Directory containing the template.
        app_path: Directory in which to create the application folder.
        slug: Short unique identifier for this application / folder.
        replace_dict: Maps template field names to replacement values.
            A field ``${institution_name}`` is filled from key
            ``institution_name``.

    Returns:
        The slug actually used on disk (may differ from ``slug`` if a
        folder of that name already existed).
    """
    # ensure slug is unique on disk so a previous run is never overwritten
    slug = ensure_unique_slug(app_path, slug)
    app_dir = os.path.join(app_path, slug)

    os.makedirs(app_dir, exist_ok=True)

    template = Document(os.path.join(template_path, template_name))
    docx_replace(template, **replace_dict)
    template.save(os.path.join(app_dir, f'cover_letter_{slug}.docx'))
    return slug


def read_csv_to_dicts(
        filepath: str,
        slug_field: str = DEFAULT_SLUG_FIELD,
        make_unique: bool = True,
) -> list[dict]:
    """Read a CSV into a list of row dicts, keyed by column header.

    The column named ``slug_field`` identifies each application. When
    ``make_unique`` is True (the default), repeated slugs within the CSV
    get a numeric suffix (``mit`` -> ``mit_2``); when False, a duplicate
    slug raises ``ValueError``. All string values are stripped of
    surrounding whitespace.
    """
    data = []
    seen_slugs = {}  # base_slug -> count
    slug_line = {}   # slug -> first line number (for error messages)

    with open(filepath, 'r', newline='', encoding='utf-8-sig') as csvfile:
        reader = csv.DictReader(csvfile, skipinitialspace=True)
        for line_num, row in enumerate(reader, start=2):  # header is line 1
            slug = row.get(slug_field, "").strip()

            if not slug:
                raise ValueError(f"Missing slug in row {line_num}")

            base_slug = slug

            if make_unique:
                # If we've seen this base slug before, append suffix
                count = seen_slugs.get(base_slug, 0) + 1
                seen_slugs[base_slug] = count
                if count > 1:
                    slug = f"{base_slug}_{count}"
                row[slug_field] = slug
            else:
                # Strict mode: error out on duplicates
                if slug in slug_line:
                    first_line = slug_line[slug]
                    raise ValueError(
                        f"Duplicate slug '{slug}' in CSV: "
                        f"first at line {first_line}, again at line {line_num}."
                    )
                slug_line[slug] = line_num

            # Strip spaces from each string value
            clean_row = {
                key: (value.strip() if isinstance(value, str) else value)
                for key, value in row.items()
            }
            data.append(clean_row)

    return data


def ensure_unique_slug(app_path: str, slug: str) -> str:
    """Return ``slug``, or ``slug_2``/``slug_3``/... if a folder of that
    name already exists in ``app_path``, so no prior run is overwritten.
    """
    base = slug
    i = 1
    full_path = os.path.join(app_path, slug)

    while os.path.exists(full_path):
        i += 1
        slug = f"{base}_{i}"
        full_path = os.path.join(app_path, slug)

    return slug


def process_csv(
        template: str,
        csv_path: str,
        out_path: str,
        slug_field: str = DEFAULT_SLUG_FIELD,
) -> list[str]:
    """Generate one cover letter per row of ``csv_path``.

    Args:
        template: Path to the .docx template (directory + file name).
        csv_path: Path to the applications CSV. Its header row names the
            template fields; one column must be ``slug_field``.
        out_path: Directory in which per-application folders are created.
        slug_field: Name of the slug column (default ``"slug"``).

    Returns:
        The list of slugs actually written on disk.
    """
    apps = read_csv_to_dicts(csv_path, slug_field=slug_field)
    template_dir, template_name = os.path.split(template)

    written = []
    for app in apps:
        replace_dict = {k: v for k, v in app.items() if k != slug_field}
        used_slug = customize_cover_letter(
            template_name=template_name,
            template_path=template_dir,
            app_path=out_path,
            slug=app[slug_field],
            replace_dict=replace_dict,
        )
        written.append(used_slug)
    return written


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        description="Batch-customize .docx cover letters from a CSV. "
                    "Template fields use ${field_name}; CSV headers name the fields.",
    )
    parser.add_argument(
        '--template', required=True,
        help="Path to the .docx template containing ${field} placeholders.",
    )
    parser.add_argument(
        '--csv', dest='csv_path', required=True,
        help="Path to the applications CSV (headers name the template fields).",
    )
    parser.add_argument(
        '--out', dest='out_path', required=True,
        help="Output directory; one folder per application is created here.",
    )
    parser.add_argument(
        '--slug-field', default=DEFAULT_SLUG_FIELD,
        help=f"CSV column identifying each application (default: {DEFAULT_SLUG_FIELD!r}).",
    )
    args = parser.parse_args(argv)

    for label, path in (('template', args.template), ('CSV', args.csv_path)):
        if not os.path.isfile(path):
            parser.error(f"{label} not found: {path}")

    written = process_csv(
        template=args.template,
        csv_path=args.csv_path,
        out_path=args.out_path,
        slug_field=args.slug_field,
    )
    print(f"Generated {len(written)} cover letter(s) in {args.out_path}")


if __name__ == '__main__':
    main()
