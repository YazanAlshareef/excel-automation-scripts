# excel-automation-scripts

Five small Python scripts that automate common Excel chores with `pandas` and `openpyxl`. Each has `--help`, a synthetic sample input, a sample output and pytest tests.

**All sample data in this repo is synthetic** (made up by `scripts/make_samples.py` with a fixed random seed). It is not real sales data.

## Scripts

| Script | What it does | Example command |
|---|---|---|
| `merge_workbooks.py` | Stack the first sheet of several workbooks into one, adding a `source_file` column | `python scripts/merge_workbooks.py sample_input/sales_north.xlsx sample_input/sales_south.xlsx -o sample_output/merged.xlsx` |
| `dedupe_rows.py` | Drop duplicate rows (all columns, or only `--keys`) | `python scripts/dedupe_rows.py sample_output/merged.xlsx -o sample_output/deduped.xlsx` |
| `validate_columns.py` | Check required columns exist and have no blanks; exit code 1 if not | `python scripts/validate_columns.py sample_output/merged.xlsx --required order_id region product units price` |
| `pivot_summary.py` | Pivot table (group + sum/mean/count/min/max) with totals | `python scripts/pivot_summary.py sample_output/deduped.xlsx -o sample_output/pivot.xlsx --index product --columns region --values units --agg sum` |
| `highlight_anomalies.py` | Fill outliers in a numeric column yellow (IQR rule) | `python scripts/highlight_anomalies.py sample_output/deduped.xlsx -o sample_output/highlighted.xlsx --column units` |

### Real output

Condensed from [sample_output/real_run.txt](sample_output/real_run.txt) (the full, unedited terminal output of running the commands above in order; I removed the command lines and blank lines and added the note on the validate line):

```
merged 2 files -> 25 rows, wrote sample_output/merged.xlsx

INVALID
 - blank in 'product' at Excel row(s): 19            (validate; exit code 1)

25 rows in, 1 duplicates removed, 24 rows out -> sample_output/deduped.xlsx

region    North  South   All
product                     
Folder       12     20    32
Notebook    911     31   942
Pen          64     29    93
Stapler      30     37    67
All        1017    117  1134

1 anomalies in 'units' at Excel row(s): 5
```

The sample data has planted problems so each script has something to find: one duplicated row, one blank `product`, and one outlier (`units` = 900).

## How to run

```bash
pip install -r requirements.txt
python scripts/make_samples.py      # optional: regenerate sample_input/
python scripts/dedupe_rows.py --help
python -m pytest -q                 # 14 tests, all passed on my machine
```

Tested with Python 3.12, pandas 3.0.5, openpyxl 3.1.5.

## How an AI coding assistant helped

An AI coding assistant wrote the first draft of the scripts and tests, and I checked the pandas and openpyxl calls against the current docs (`pivot_table`, `drop_duplicates`, `PatternFill`). I read and ran every file; see [WALKTHROUGH.md](WALKTHROUGH.md) for a plain-English explanation of each one.

## What broke and how it was fixed

Honest account: no test failed on the final run, but these were pitfalls I hit or designed around:

- **Anomaly rule.** I chose IQR over a z-score on purpose: an outlier inflates the mean and standard deviation, which hides it on small tables (and on tiny tables a z-score of 3 cannot be reached at all). IQR uses quartiles, which one extreme value barely moves.
- **Row numbers.** pandas row 0 is Excel row 2 (header + 1-based). Reports add 2; tests check the exact row (`19`, `5`).
- **Whitespace.** A cell containing only spaces looks empty in Excel, so `validate_columns.py` counts it as blank.
- **Blank cells in pivots.** The row with a blank `product` is silently left out of the pivot (pandas drops missing group keys). That is why `validate_columns.py` exists: run it first.

## Limits

- Only the first sheet of each workbook is read (the highlighter also edits that first sheet).
- `merge_workbooks.py` writes cells that start with `=` as Excel formulas, so only merge files you trust.
- Scripts refuse an output path equal to an input path (exit 2) so you cannot overwrite your source.
- Everything is loaded into memory; fine for thousands of rows, not millions.
- `highlight_anomalies.py` keeps cell styles of the input, but charts, images and some features may be lost when openpyxl re-saves.
- IQR is a simple rule: it flags unusual values, not necessarily wrong ones.
- Dedupe compares exact values; `"Pen"` and `"pen "` count as different.
- Not tested on very large files, `.xls`, or password-protected workbooks. Not opened in desktop Excel; I checked the highlight by reading the fill back with openpyxl in a test.

## License

MIT, see [LICENSE](LICENSE).
