"""Check that a workbook has the required columns and no blanks in them.

Exit code 0 = valid, 1 = problems found, 2 = bad usage.
"""
import argparse
import sys

import pandas as pd


def validate(df, required):
    """Return a list of problem strings (empty list = valid)."""
    problems = [f"missing column: {c}" for c in required if c not in df.columns]
    for c in required:
        if c in df.columns:
            empty = df[c].map(lambda v: pd.isna(v) or (isinstance(v, str) and not v.strip()))  # whitespace-only counts as blank
            blanks = df.index[empty] + 2  # +2: header row + 1-based Excel rows
            if len(blanks):
                problems.append(f"blank in '{c}' at Excel row(s): {', '.join(map(str, blanks))}")
    return problems


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.strip().splitlines()[0])
    ap.add_argument("input", help="input .xlsx file")
    ap.add_argument("--required", nargs="+", required=True, help="column names that must exist and be filled")
    args = ap.parse_args(argv)
    problems = validate(pd.read_excel(args.input), args.required)
    if problems:
        print("INVALID")
        for p in problems:
            print(" -", p)
        return 1
    print("OK: all required columns present and filled")
    return 0


if __name__ == "__main__":
    sys.exit(main())
