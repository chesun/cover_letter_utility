"""Command-line interface: ``cover-letter --template ... --csv ... --out ...``."""
from __future__ import annotations

import argparse
import os

from cover_letter_utility import __version__
from cover_letter_utility.core import DEFAULT_SLUG_FIELD, process_csv


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog='cover-letter',
        description="Batch-customize .docx cover letters from a CSV. "
                    "Template fields use ${field_name}; CSV headers name the fields.",
    )
    parser.add_argument(
        '--version', action='version', version=f'%(prog)s {__version__}',
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
    return parser


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
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
