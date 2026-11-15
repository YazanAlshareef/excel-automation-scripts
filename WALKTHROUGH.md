# Walkthrough: every file in plain English

## Top level

- **README.md**: what the project is, the table of scripts, real output, limits.
- **WALKTHROUGH.md**: this file.
- **LICENSE**: MIT. Anyone may reuse the code; no warranty.
- **.github/workflows/ci.yml**: GitHub Actions. On each push it installs the requirements and runs `python -m pytest -q` on Ubuntu and Windows.
- **.gitignore**: tells git to skip Python cache folders and Excel temp files (`~$...`).
- **requirements.txt**: the three packages needed: pandas (tables), openpyxl (read/write `.xlsx`, cell styling), pytest (tests).

## scripts/

- **_common.py**: one small helper, `existing_file`, used as an argparse `type=` so a wrong input path gives "file not found" and exit 2 instead of a traceback.

All scripts follow one pattern: `argparse` reads the command line (and gives `--help` for free), a small function does the real work, `main()` wires them together and returns an exit code (0 = fine, 2 = bad input, 1 = validation failed).

- **make_samples.py**: builds the two synthetic input workbooks with a fixed random seed (so they are identical every run). Plants three problems on purpose: an outlier (`units` 900), a duplicated row, and a blank `product`.
- **merge_workbooks.py**: refuses an output path equal to an input, reads the first sheet of each input with `pd.read_excel`, adds a `source_file` column, joins them with `pd.concat`, writes one file.
- **dedupe_rows.py**: `drop_duplicates(subset=keys, keep="first")`. With no `--keys` it compares whole rows. Rejects key names that are not columns, and refuses an output path equal to the input.
- **validate_columns.py**: for each required column, reports "missing" if absent, or lists the Excel row numbers of blank cells (`isna()` or whitespace-only text; index + 2 converts to Excel rows). Exit code 1 means problems, so it can gate other steps in a shell script.
- **pivot_summary.py**: `pd.pivot_table` with the chosen row column, optional column-spread, value column and aggregate; `margins=True` adds the "All" totals. Prints the table and saves it. A non-numeric `--values` column with sum/mean gives a clear error (exit 2), not a traceback.
- **highlight_anomalies.py**: pandas finds outliers with the IQR rule (below Q1 − 1.5×IQR or above Q3 + 1.5×IQR; `--k` changes 1.5). Then openpyxl reopens the original workbook and sets a yellow `PatternFill("solid")` on those cells, so the original formatting is kept. It always edits the first sheet, the same one pandas read.

## sample_input/ (synthetic)

- **sales_north.xlsx, sales_south.xlsx**: made-up orders (id, date, region, product, units, price). Produced by `make_samples.py`.

## sample_output/ (produced by really running the scripts)

- **merged.xlsx, deduped.xlsx, pivot.xlsx, highlighted.xlsx**: results of the commands in the README, in order.
- **real_run.txt**: the exact terminal output of that run.

## tests/

- **test_scripts.py**: pytest. Builds a small workbook in a temp folder, then checks: every script's `--help` works; merge row count and `source_file`; dedupe with and without `--keys` and with a bad key; validate on valid, blank and missing cases; pivot totals; only the true outlier cell is filled (even when another sheet is active); output=input is refused; whitespace counts as blank; mean on a text column exits 2; a missing input file exits 2; highlighting a text column exits 2. Run with `python -m pytest -q`.
