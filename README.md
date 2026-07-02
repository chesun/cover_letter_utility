# cover_letter_utility

Batch-customize Word (`.docx`) cover letters from a spreadsheet. Write one
template with `${placeholder}` fields, list your applications in a CSV, and
generate a personalized letter for each — one folder per application.

Built for the academic job market, where a candidate may send near-identical
letters to dozens of institutions that differ only in a few fields (name,
department, title, date). Instead of editing dozens of Word files by hand, you
edit one template and one CSV.

## How it works

1. In your Word template, write `${field_name}` anywhere a value should change
   per application — for example `${institution_name}`, `${dept_name}`,
   `${job_title}`, `${letter_date}`. The `${...}` syntax is required (it comes
   from [python-docx-replace](https://pypi.org/project/python-docx-replace/)).
2. In a CSV, make one column named `slug` (a short id used for the output
   folder) plus one column per placeholder. Each row is one application:

   ```csv
   slug,institution_name,dept_name,job_title,letter_date
   example_college_ap,Example College,Department of Economics,Assistant Professor,"November 16, 2025"
   sample_university_pd,Sample University,School of Public Policy,Postdoctoral Fellow,"November 16, 2025"
   ```

3. Run the tool. For each row it fills the template and writes
   `cover_letter_<slug>.docx` into a folder named `<slug>` under your output
   directory.

Column headers in the CSV must match the placeholder names in the template
exactly (minus the `${}`).

## Install

Requires Python 3.9+.

```bash
pip install -r requirements.txt
```

Runtime dependencies are just `python-docx` and `python-docx-replace`. For the
notebook, install the dev extras instead: `pip install -r requirements-dev.txt`.

## Usage

### Command line

```bash
cd src
python utility.py \
    --template example_template.docx \
    --csv example_apps.csv \
    --out ./demo_output
```

Options:

- `--template` — path to the `.docx` template containing `${field}` placeholders.
- `--csv` — path to the applications CSV (its headers name the template fields).
- `--out` — output directory; one folder per application is created here.
- `--slug-field` — name of the id column (default: `slug`).

The repo ships `src/example_template.docx` and `src/example_apps.csv` so the
command above runs out of the box.

### From Python

```python
from utility import process_csv

process_csv(
    template="example_template.docx",
    csv_path="example_apps.csv",
    out_path="./demo_output",
)
```

`src/customize_letter.ipynb` walks through the same workflow in a notebook.

## Notes

- **Output is `.docx` only.** Export to PDF from Word or LibreOffice when you're
  ready to submit. (Automated `.docx`→PDF conversion is unreliable across
  platforms, so it's intentionally left out.)
- **Re-runs never overwrite.** If a `slug` folder already exists, the next run
  writes to `slug_2`, `slug_3`, and so on — so an accidental re-run can't clobber
  letters you've already edited by hand.
- **Duplicate slugs within a CSV** are also disambiguated with a numeric suffix.

## License

MIT — see [LICENSE](LICENSE).
