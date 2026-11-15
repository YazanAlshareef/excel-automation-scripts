"""Build a pivot summary (group + aggregate) and save it as a workbook."""
import argparse
import sys

import pandas as pd
from _common import existing_file


def pivot(df, index, values, columns=None, agg="sum"):
    return pd.pivot_table(df, index=index, columns=columns, values=values, aggfunc=agg, margins=True)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.strip().splitlines()[0])
    ap.add_argument("input", type=existing_file, help="input .xlsx file")
    ap.add_argument("-o", "--output", required=True, help="output .xlsx file")
    ap.add_argument("--index", required=True, help="column for the rows")
    ap.add_argument("--values", required=True, help="numeric column to aggregate")
    ap.add_argument("--columns", help="optional column to spread across the top")
    ap.add_argument("--agg", default="sum", choices=["sum", "mean", "count", "min", "max"])
    args = ap.parse_args(argv)
    df = pd.read_excel(args.input)
    need = [c for c in (args.index, args.values, args.columns) if c]
    missing = [c for c in need if c not in df.columns]
    if missing:
        print(f"error: unknown column(s): {', '.join(missing)}", file=sys.stderr)
        return 2
    try:
        table = pivot(df, args.index, args.values, args.columns, args.agg).round(2)
    except (TypeError, ValueError, pd.errors.DataError) as e:
        print(f"error: cannot apply '{args.agg}' to column '{args.values}' (is it numeric?): {e}", file=sys.stderr)
        return 2
    table.to_excel(args.output)
    print(table.to_string())
    print(f"wrote {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
