"""Utility functions to replace fields in cover letter (.docx) and save as .pdf"""
import os
import socket
import csv
from dataclasses import dataclass, field
from python_docx_replace import docx_replace
from docx import Document

__all__ = ['PathData', 'FileData',
           'customize_cover_letter', 'read_csv_to_dicts']


class InvalidHostError(Exception):
    """Custom exception for unknown host machine"""
    pass


@dataclass
class PathData:
    """File path strings"""
    hostname: str = socket.gethostname()
    template_path: str = field(init=False)
    app_path: str = field(init=False)

    def __post_init__(self):
        if self.hostname == 'DH444T2TQ9':
            self.template_path = '/Users/chesun1/Dropbox/Davis/job_market/cover_letter'
            self.app_path = ';/Users/chesun1/Dropbox/Davis/job_market/applications/US'
        else:
            self.template_path = '/Users/christinasun/Library/CloudStorage/Dropbox/Davis/job_market/cover_letter'
            self.app_path = '/Users/christinasun/Library/CloudStorage/Dropbox/Davis/job_market/applications/US'


@dataclass
class FileData:
    """Cover letter template file names"""
    applied_template_name: str = 'applied_micro_academic_cover_letter_template.docx'
    behavioral_template_name: str = 'behavioral_academic_cover_letter_template.docx'
    teaching_template_name: str = 'teaching_cover_letter_template.docx'


# a function to replace fields in cover letter template and save to application directory
def customize_cover_letter(
        template_name: str,
        template_path: str,
        app_path: str,
        slug: str,
        replace_dict: dict
) -> None:
    """Function to replace key fields using replace_dict in cover letter template, 
    create a directory with name of slug, and save new letter in directory 
    Args:
        template_name (str): full path name for cover letter template
        template_path (str): path for cover letter template
        app_path (str): path for applications
        slug (str): unique slug for application folder
        replace_dict (dict): dictionary with where keys are field names in template, 
        and values are replacement values
    Returns:
        None
    """
    # ensure slug is unique on disk
    slug = ensure_unique_slug(app_path, slug)
    app_dir = os.path.join(app_path, slug)

    # create application directory
    os.makedirs(app_dir,
                exist_ok=True)
    # get template
    template = Document(os.path.join(template_path, template_name))
    # replace fields using dict
    docx_replace(template, **replace_dict)
    # save in new application directory
    template.save(os.path.join(app_dir, f'cover_letter_{slug}.docx'))


def read_csv_to_dicts(
        filepath: str,
        slug_field: str = "slug",
        make_unique: bool = True
) -> list[dict]:
    """
    Reads a CSV file and returns a list of dictionaries,
    where each dictionary represents a row.
    Default to making the values in `slug_field` unique.
    """
    data = []
    seen_slugs = {}  # base_slug -> count
    slug_line = {}   # slug -> first line number (for error messages)

    with open(filepath, 'r', newline='', encoding='utf-8-sig') as csvfile:
        reader = csv.DictReader(csvfile, skipinitialspace=True)
        for line_num, row in enumerate(reader, start=2):  # header is line 1
            slug = row.get(slug_field, "").strip()

            if not slug:
                # you can decide whether to allow blank slugs or raise
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
    """
    If a folder with `slug` already exists in app_path,
    append _2, _3, ... until we find a free one.
    """
    base = slug
    i = 1
    full_path = os.path.join(app_path, slug)

    while os.path.exists(full_path):
        i += 1
        slug = f"{base}_{i}"
        full_path = os.path.join(app_path, slug)

    return slug
