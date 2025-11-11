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
        elif self.hostname == 'Christinas-MacBook-Air.local':
            self.template_path = '/Users/christinasun/Library/CloudStorage/Dropbox/Davis/job_market/cover_letter'
            self.app_path = '/Users/christinasun/Library/CloudStorage/Dropbox/Davis/job_market/applications/US'
        else:
            raise InvalidHostError("Unknown host name")


@dataclass
class FileData:
    """Cover letter template file names"""
    applied_template_name: str = 'applied_micro_academic_cover_letter_template.docx'
    behavioral_template_name: str = 'behavioral_academic_cover_letter_template.docx'


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
    # create application directory
    os.makedirs(os.path.join(app_path, slug),
                exist_ok=True)
    # get template
    template = Document(os.path.join(template_path, template_name))
    # replace fields using dict
    docx_replace(template, **replace_dict)
    # save in new application directory
    template.save(os.path.join(app_path, slug,
                  'cover_letter_' + slug + '.docx'))


def read_csv_to_dicts(filepath: str) -> list[dict]:
    """
    Reads a CSV file and returns a list of dictionaries,
    where each dictionary represents a row.
    """
    data = []
    with open(filepath, 'r', newline='', encoding='utf-8-sig') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            data.append(row)
    return data
