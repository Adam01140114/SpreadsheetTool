import pandas as pd
from pathlib import Path
import sys

def find_conflicting_premises(
    input_path,
    output_path="Premise_Conflicts.xlsx"
):
    # Load data
    df = pd.read_excel(input_path)

    # Ensure required columns exist
    required_cols = ["AcctNum", "Premise", "MeterNum"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # 1. Count distinct AcctNum and MeterNum per Premise
    grp = df.groupby("Premise").agg(
        distinct_accts=("AcctNum", "nunique"),
        distinct_meters=("MeterNum", "nunique"),
        row_count=("Premise", "size")
    ).reset_index()

    # 2. Filter to premises that have >1 acct OR >1 meter
    conflict_premises = grp[
        (grp["distinct_accts"] > 1) | (grp["distinct_meters"] > 1)
    ]

    # 3. Pull full detail rows for those premises
    conflict_rows = df[df["Premise"].isin(conflict_premises["Premise"])].copy()

    # Sort for easier reading
    conflict_rows = conflict_rows.sort_values(
        ["Premise", "AcctNum", "MeterNum", "READ DATE HIST"],
        na_position="last"
    )

    # 4. Save two sheets:
    #    - Summary: one row per Premise with counts
    #    - Detail: all rows for those Premises
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        conflict_premises.to_excel(writer, sheet_name="Summary", index=False)
        conflict_rows.to_excel(writer, sheet_name="Detail", index=False)

    print(f"Found {len(conflict_premises)} premises with conflicts.")
    print(f"Total rows in Detail sheet: {len(conflict_rows)}")
    print(f"Saved results to: {Path(output_path).resolve()}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python find_conflicting_premises.py <excel_file>")
        raise SystemExit(1)

    input_file = sys.argv[1]
    find_conflicting_premises(input_file)
