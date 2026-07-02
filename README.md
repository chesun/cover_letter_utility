# cover_letter_utility

**Write your cover letter once, then generate a personalized copy for every job
you apply to — automatically.**

If you're on the academic job market, you might send the same letter to dozens of
places that differ in only a few words: the institution's name, the department,
the job title, the date. Editing each one by hand is slow and easy to get wrong
(there's a special kind of dread in realizing you sent School B a letter
addressed to School A). This tool does the swapping for you: you keep one
template and one spreadsheet, and it produces a tidy, individually-addressed
letter for each application.

You don't need to be a programmer to use it — if you can edit a Word document and
a spreadsheet and copy-paste a couple of commands, you're set.

## Quick start

You'll need [Python 3.9 or newer](https://www.python.org/downloads/). Then, in a
terminal:

```bash
git clone https://github.com/chesun/cover_letter_utility.git
cd cover_letter_utility
pip install -e .
cover-letter --template examples/example_template.docx --csv examples/example_apps.csv --out ./demo_output
```

That last command uses the bundled example files and drops three finished letters
into a new `demo_output/` folder. Open them to see what happened, then swap in
your own template and spreadsheet (next section).

## How it works

There are two things you provide: a **template** (your letter) and a **CSV** (your
list of applications).

1. In your Word template, type `${field_name}` anywhere a value should change from
   one application to the next — for example `${institution_name}`, `${dept_name}`,
   `${job_title}`, `${letter_date}`. Everything else stays exactly as you wrote it.
   The `${...}` marker is what tells the tool "replace this."

2. In a spreadsheet saved as CSV, make one column called `slug` (a short nickname
   for each application — it becomes the folder name) plus one column for each
   `${field_name}` in your letter. Each row is one job:

   ```csv
   slug,institution_name,dept_name,job_title,letter_date
   example_college_ap,Example College,Department of Economics,Assistant Professor,"November 16, 2025"
   sample_university_pd,Sample University,School of Public Policy,Postdoctoral Fellow,"November 16, 2025"
   ```

3. Run the tool. For each row it fills in your template and saves
   `cover_letter_<slug>.docx` inside a folder named after that application.

The one rule to remember: the column names in your spreadsheet must match the
`${field_name}` markers in your letter exactly (just without the `${}`).

## Install

Requires Python 3.9+.

```bash
pip install -e .
```

This installs the tool (the `cover-letter` command) along with everything it needs
to run. Two optional add-ons: `pip install -e ".[dev]"` also installs the test
tools, and `pip install -e ".[notebook]"` adds what you need to run the example
notebook.

## Usage

### Command line

The `cover-letter` command comes with the package. The repo ships an example
template and CSV, so this works immediately:

```bash
cover-letter \
    --template examples/example_template.docx \
    --csv examples/example_apps.csv \
    --out ./demo_output
```

Options:

- `--template` — path to the `.docx` template containing `${field}` placeholders.
- `--csv` — path to the applications CSV (its headers name the template fields).
- `--out` — output directory; one folder per application is created here.
- `--slug-field` — name of the id column (default: `slug`).
- `--version` — print the version and exit.

### From Python

If you'd rather call it from your own script or a notebook:

```python
from cover_letter_utility import process_csv

process_csv(
    template="examples/example_template.docx",
    csv_path="examples/example_apps.csv",
    out_path="./demo_output",
)
```

`examples/customize_letter.ipynb` walks through the same workflow step by step in a
notebook.

## Good to know

- **You get `.docx` files, not PDFs.** Open a finished letter in Word or
  LibreOffice and export to PDF there when you're ready to submit. (Automatic
  PDF conversion is unreliable across computers, so it's left out on purpose.)
- **Running it again won't overwrite your work.** If a `slug` folder already
  exists, the next run writes to `slug_2`, `slug_3`, and so on — so a rerun can
  never clobber a letter you've already tweaked by hand.
- **Repeated slugs in one spreadsheet** are handled the same way, with a numeric
  suffix, so nothing silently collides.

## Development

```bash
pip install -e ".[dev]"
pytest
```

## License

MIT — see [LICENSE](LICENSE).
