# Improvement log

## 2026-11-15, cycle 1

Baseline: 20 tests pass (`python -m pytest -q`). No linter or type checker is configured.

### Done this cycle (branches awaiting review, not merged)

| Item | Branch | Commit | Result |
|---|---|---|---|
| pivot_summary could overwrite its own input file (scripts/pivot_summary.py:21-33). Bug fix, behaviour change: output equal to input now exits 2, as the other scripts already do. | fix/pivot-overwrite-guard | 5ddefa0 | 21 tests pass |
| A corrupt or non-Excel input gave a raw traceback (scripts/_common.py:6-10). Bug fix, behaviour change: all scripts now exit 2 with a short message. validate_columns no longer exits 1 on a crash. Legacy .xls files are also rejected by this check. | fix/bad-excel-input | 0c62543 | 25 tests pass |
| merge_workbooks silently replaced an existing `source_file` column (scripts/merge_workbooks.py:20). Bug fix, behaviour change: an input with that column now exits 2. Any pandas ValueError during the merge also becomes exit 2. | fix/merge-source-file-clash | e3c4994 | 21 tests pass |

All three change behaviour. Each branch holds one commit on top of main and was reviewed on its own.

### Merge note

fix/bad-excel-input and fix/merge-source-file-clash both edit tests/test_scripts.py, so they may conflict when merged together. Merge one, then rerun the tests before the next.

### Not verified

- The three branches were not tested together. Each count above is for that branch alone.
- Password-protected files and symlinked paths were not tried.

### Backlog, re-ranked

1. The output-equals-input check is copy-pasted in three scripts, which is how pivot_summary missed it (scripts/dedupe_rows.py:20-22, scripts/highlight_anomalies.py:28-30, scripts/merge_workbooks.py:30-32). Do after the pivot branch lands, since it touches the same lines. Effort S, risk low.
2. Missing tests for an empty workbook, `--columns`, `--k`, other `--agg` values, a multi-sheet merge and NaN gaps (tests/test_scripts.py, README.md:96). Effort M, risk low.
3. The output folder is not checked before writing (scripts/merge_workbooks.py:34, scripts/dedupe_rows.py:30, scripts/pivot_summary.py:33, scripts/highlight_anomalies.py:45). Probable from the code, not run. Effort S, risk low.
4. highlight_anomalies maps rows as `i + 2`, which assumes the header is on row 1 (scripts/highlight_anomalies.py:44 and :46). Not reproduced. Documenting it is enough. Effort S, risk medium if changed in code.
5. requirements.txt has only lower bounds, and the README reports the tested versions (requirements.txt:1-3, README.md:103). The README also says CI has not been seen (README.md:103). Effort S, risk low.

### Lessons

- The README promised that scripts refuse to overwrite their input (README.md:49). One script did not, and no test covered it (tests/test_scripts.py:231-237).
- Putting the unreadable-file check in the shared helper fixed all five scripts with one change, but it also changed the .xls behaviour.
- Two branches that add tests to the same file will conflict even when the code changes are separate.

### Diagrams

Architecture and sequence diagrams are in `docs/diagrams/` (components.md, highlight-anomalies.md, sequences.md). They live on the `docs/diagrams` branch, not on main.

### Status

All branches are local. Nothing is pushed or merged. They wait for the owner's approval.
