"""Create the SYNTHETIC sample workbooks in sample_input/. All data is made up.

Usage: python scripts/make_samples.py
"""
import random
from pathlib import Path

import pandas as pd

OUT = Path(__file__).resolve().parent.parent / "sample_input"


def make(region, start_id, n, rng):
    rows = []
    for i in range(n):
        rows.append({
            "order_id": start_id + i,
            "date": f"2026-0{rng.randint(1, 6)}-{rng.randint(10, 28)}",
            "region": region,
            "product": rng.choice(["Pen", "Notebook", "Stapler", "Folder"]),
            "units": rng.randint(1, 20),
            "price": round(rng.uniform(2, 15), 2),
        })
    return pd.DataFrame(rows)


def main():
    rng = random.Random(42)  # fixed seed so output is reproducible
    OUT.mkdir(exist_ok=True)
    north = make("North", 1000, 12, rng)
    south = make("South", 2000, 12, rng)
    north.loc[3, "units"] = 900  # planted outlier for highlight_anomalies
    south = pd.concat([south, south.iloc[[2]]], ignore_index=True)  # planted duplicate row
    south.loc[5, "product"] = None  # planted blank for validate_columns
    north.to_excel(OUT / "sales_north.xlsx", index=False)
    south.to_excel(OUT / "sales_south.xlsx", index=False)
    print("wrote sample_input/sales_north.xlsx and sales_south.xlsx (synthetic)")


if __name__ == "__main__":
    main()
