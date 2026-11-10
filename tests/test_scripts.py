import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

import dedupe_rows, highlight_anomalies, merge_workbooks, pivot_summary, validate_columns  # noqa: E402

SCRIPTS = ["merge_workbooks", "dedupe_rows", "validate_columns", "pivot_summary", "highlight_anomalies"]


@pytest.fixture
def book(tmp_path):
    df = pd.DataFrame({
        "id": [1, 2, 2, 3, 4, 5, 6, 7],
        "cat": ["a", "b", "b", "a", None, "b", "a", "b"],
        "n": [10, 11, 11, 12, 10, 11, 12, 500],
    })
    p = tmp_path / "in.xlsx"
    df.to_excel(p, index=False)
    return p


@pytest.mark.parametrize("name", SCRIPTS)
def test_help_works(name):
    r = subprocess.run([sys.executable, str(ROOT / "scripts" / f"{name}.py"), "--help"], capture_output=True, text=True)
    assert r.returncode == 0 and "usage" in r.stdout.lower()


def test_merge_adds_source_file(book, tmp_path):
    out = tmp_path / "m.xlsx"
    assert merge_workbooks.main([str(book), str(book), "-o", str(out)]) == 0
    df = pd.read_excel(out)
    assert len(df) == 16 and set(df["source_file"]) == {"in.xlsx"}


def test_dedupe_all_columns_and_keys(book, tmp_path):
    out = tmp_path / "d.xlsx"
    assert dedupe_rows.main([str(book), "-o", str(out)]) == 0
    assert len(pd.read_excel(out)) == 7
    assert dedupe_rows.main([str(book), "-o", str(out), "--keys", "cat"]) == 0
    assert len(pd.read_excel(out)) == 3  # a, b, blank
    assert dedupe_rows.main([str(book), "-o", str(out), "--keys", "nope"]) == 2


def test_validate_reports_missing_and_blank(book):
    assert validate_columns.main([str(book), "--required", "id", "n"]) == 0
    assert validate_columns.main([str(book), "--required", "cat"]) == 1  # blank at Excel row 6
    assert validate_columns.main([str(book), "--required", "zzz"]) == 1
    df = pd.read_excel(book)
    assert validate_columns.validate(df, ["cat"]) == ["blank in 'cat' at Excel row(s): 6"]


def test_pivot_values(book, tmp_path):
    out = tmp_path / "p.xlsx"
    assert pivot_summary.main([str(book), "-o", str(out), "--index", "cat", "--values", "n"]) == 0
    t = pd.read_excel(out, index_col=0)
    assert t.loc["a", "n"] == 34 and t.loc["All", "n"] == 10 + 11 + 11 + 12 + 11 + 12 + 500


def test_highlight_marks_only_outlier(book, tmp_path):
    out = tmp_path / "h.xlsx"
    assert highlight_anomalies.main([str(book), "-o", str(out), "--column", "n"]) == 0
    ws = load_workbook(out).active
    filled = [c.row for row in ws.iter_rows() for c in row if c.fill.fgColor.rgb == "00FFFF00"]
    assert filled == [9]  # the 500, Excel row 9
    assert highlight_anomalies.main([str(book), "-o", str(out), "--column", "nope"]) == 2
