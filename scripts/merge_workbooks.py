"""Merge several Excel workbooks (first sheet of each) into one.

A `source_file` column records which file each row came from.
"""
import argparse
import sys
from pathlib import Path

import pandas as pd


def merge(paths):
    frames = []
    for p in paths:
        df = pd.read_excel(p)
        df["source_file"] = Path(p).name
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.strip().splitlines()[0])
    ap.add_argument("inputs", nargs="+", help="input .xlsx files")
    ap.add_argument("-o", "--output", required=True, help="output .xlsx file")
    args = ap.parse_args(argv)
    out = merge(args.inputs)
    out.to_excel(args.output, index=False)
    print(f"merged {len(args.inputs)} files -> {len(out)} rows, wrote {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
