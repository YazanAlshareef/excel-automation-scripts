"""Remove duplicate rows from a workbook, keeping the first occurrence."""
import argparse
import sys
from pathlib import Path

import pandas as pd
from _common import existing_file


def dedupe(df, keys=None):
    return df.drop_duplicates(subset=keys, keep="first")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.strip().splitlines()[0])
    ap.add_argument("input", type=existing_file, help="input .xlsx file")
    ap.add_argument("-o", "--output", required=True, help="output .xlsx file")
    ap.add_argument("--keys", nargs="+", help="columns that define a duplicate (default: all columns)")
    args = ap.parse_args(argv)
    if Path(args.output).resolve() in {Path(p).resolve() for p in [args.input]}:
        print("error: output path is the same as an input file; pick a different -o", file=sys.stderr)
        return 2
    df = pd.read_excel(args.input)
    if args.keys:
        missing = [k for k in args.keys if k not in df.columns]
        if missing:
            print(f"error: unknown column(s): {', '.join(missing)}", file=sys.stderr)
            return 2
    out = dedupe(df, args.keys)
    out.to_excel(args.output, index=False)
    print(f"{len(df)} rows in, {len(df) - len(out)} duplicates removed, {len(out)} rows out -> {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
