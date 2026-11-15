"""Merge several Excel workbooks (first sheet of each) into one.

A `source_file` column records which file each row came from.
An input that already has a `source_file` column is refused (exit 2).

WARNING: text cells starting with "=" are written back as Excel formulas.
Only merge workbooks you trust; a malicious file could plant formulas.
"""
import argparse
import sys
from pathlib import Path

import pandas as pd
from _common import existing_file


def merge(paths):
    frames = []
    for p in paths:
        df = pd.read_excel(p)
        if "source_file" in df.columns:
            raise ValueError(f"{Path(p).name} already has a 'source_file' column; rename it first")
        df["source_file"] = Path(p).name
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.strip().splitlines()[0])
    ap.add_argument("inputs", nargs="+", type=existing_file, help="input .xlsx files")
    ap.add_argument("-o", "--output", required=True, help="output .xlsx file")
    args = ap.parse_args(argv)
    if Path(args.output).resolve() in {Path(p).resolve() for p in args.inputs}:
        print("error: output path is the same as an input file; pick a different -o", file=sys.stderr)
        return 2
    try:
        out = merge(args.inputs)
    except ValueError as e:
        print(f"error: {e}", file=sys.stderr)
        return 2
    out.to_excel(args.output, index=False)
    print(f"merged {len(args.inputs)} files -> {len(out)} rows, wrote {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
