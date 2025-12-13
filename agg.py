import pandas as pd
from pathlib import Path
import sys

def aggregate_usage_by_month(
    input_path,
    output_path="Usage_By_Month_Nov24_to_Oct25_ORDERED.xlsx"
):
    # 1. Load the raw export
    df = pd.read_excel(input_path)

    # Expect these columns to exist:
    # Premise, MeterNum, Status, Cyc, READ DATE HIST, USAGE HIST

    # 2. Convert read date to datetime
    df["READ DATE HIST"] = pd.to_datetime(df["READ DATE HIST"], errors="coerce")

    # 3. Create a month label like "Nov24", "Dec24", "Jan25", etc.
    df["MonthLabel"] = df["READ DATE HIST"].dt.strftime("%b%y")

    # 4. Pivot into one row per Premise / Meter / Status / Cyc
    #    and one column per month with summed usage
    pivot = df.pivot_table(
        index=["Premise", "MeterNum", "Status", "Cyc"],
        columns="MonthLabel",
        values="USAGE HIST",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    # 5. Order the month columns from Nov24 through Oct25
    base_cols = ["Premise", "MeterNum", "Status", "Cyc"]
    month_order = [
        "Nov24", "Dec24",
        "Jan25", "Feb25", "Mar25", "Apr25", "May25",
        "Jun25", "Jul25", "Aug25", "Sep25", "Oct25",
    ]

    # Keep only the month columns that actually exist in this dataset
    existing_months = [m for m in month_order if m in pivot.columns]

    ordered = pivot[base_cols + existing_months]

    # 6. Save to Excel
    ordered.to_excel(output_path, index=False)
    print(f"Saved aggregated file to: {Path(output_path).resolve()}")

if __name__ == "__main__":
    # Get Excel file name from command line
    if len(sys.argv) < 2:
        print("Usage: python agg.py <excel_file>")
        raise SystemExit(1)

    input_file = sys.argv[1]
    aggregate_usage_by_month(input_file)
