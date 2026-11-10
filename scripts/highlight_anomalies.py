"""Highlight outliers in a numeric column (IQR rule) with a yellow fill."""
import argparse
import sys

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

YELLOW = PatternFill("solid", fgColor="FFFF00")


def find_outliers(series, k=1.5):
    """Index labels whose value is outside [Q1 - k*IQR, Q3 + k*IQR]."""
    q1, q3 = series.quantile(0.25), series.quantile(0.75)
    iqr = q3 - q1
    return series.index[(series < q1 - k * iqr) | (series > q3 + k * iqr)]


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.strip().splitlines()[0])
    ap.add_argument("input", help="input .xlsx file")
    ap.add_argument("-o", "--output", required=True, help="output .xlsx file")
    ap.add_argument("--column", required=True, help="numeric column to check")
    ap.add_argument("--k", type=float, default=1.5, help="IQR multiplier (default 1.5)")
    args = ap.parse_args(argv)
    df = pd.read_excel(args.input)
    if args.column not in df.columns:
        print(f"error: unknown column: {args.column}", file=sys.stderr)
        return 2
    bad = find_outliers(pd.to_numeric(df[args.column], errors="coerce"), args.k)
    wb = load_workbook(args.input)
    ws = wb.active
    col = list(df.columns).index(args.column) + 1
    for i in bad:
        ws.cell(row=i + 2, column=col).fill = YELLOW  # +2: header row + 1-based rows
    wb.save(args.output)
    print(f"{len(bad)} anomalies in '{args.column}' at Excel row(s): {', '.join(str(i + 2) for i in bad) or 'none'}")
    print(f"wrote {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
