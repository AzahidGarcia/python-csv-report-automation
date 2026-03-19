"""
CSV Report Automation Script
Automates data processing and Excel report generation from CSV files.
Author: Azahid García | github.com/azahidgarcia
"""

import pandas as pd
import os
import sys
from datetime import datetime
from pathlib import Path


# ── Configuration ─────────────────────────────────────────────
INPUT_FOLDER  = "data/input"
OUTPUT_FOLDER = "data/output"
REPORT_NAME   = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"


# ── Data Loading ───────────────────────────────────────────────
def load_csv_files(folder: str) -> pd.DataFrame:
    """Load and combine all CSV files from a folder."""
    folder_path = Path(folder)
    csv_files = list(folder_path.glob("*.csv"))

    if not csv_files:
        print(f"❌ No CSV files found in '{folder}'")
        sys.exit(1)

    print(f"📂 Found {len(csv_files)} CSV file(s): {[f.name for f in csv_files]}")

    dataframes = []
    for file in csv_files:
        try:
            df = pd.read_csv(file)
            df["source_file"] = file.name
            dataframes.append(df)
            print(f"   ✓ Loaded {file.name} — {len(df)} rows")
        except Exception as e:
            print(f"   ⚠ Skipped {file.name}: {e}")

    return pd.concat(dataframes, ignore_index=True)


# ── Data Cleaning ──────────────────────────────────────────────
def clean_data(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Clean dataset and return cleaned df + quality report."""
    original_rows = len(df)
    report = {}

    # 1. Remove duplicates
    df = df.drop_duplicates()
    report["duplicates_removed"] = original_rows - len(df)

    # 2. Strip whitespace from string columns
    str_cols = df.select_dtypes(include="object").columns
    df[str_cols] = df[str_cols].apply(lambda x: x.str.strip())

    # 3. Track null counts
    report["null_counts"] = df.isnull().sum().to_dict()
    report["null_total"] = int(df.isnull().sum().sum())

    # 4. Drop fully empty rows
    before = len(df)
    df = df.dropna(how="all")
    report["empty_rows_removed"] = before - len(df)

    report["rows_original"] = original_rows
    report["rows_clean"]    = len(df)

    print(f"\n🧹 Cleaning complete:")
    print(f"   Duplicates removed : {report['duplicates_removed']}")
    print(f"   Empty rows removed : {report['empty_rows_removed']}")
    print(f"   Nulls remaining    : {report['null_total']}")
    print(f"   Final row count    : {report['rows_clean']}")

    return df, report


# ── Summary Statistics ─────────────────────────────────────────
def generate_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Generate summary statistics for numeric columns."""
    numeric_cols = df.select_dtypes(include="number").columns
    if len(numeric_cols) == 0:
        return pd.DataFrame({"info": ["No numeric columns found"]})
    return df[numeric_cols].describe().round(2)


# ── Excel Report ───────────────────────────────────────────────
def export_report(df: pd.DataFrame, summary: pd.DataFrame,
                  quality: dict, output_path: str) -> None:
    """Export cleaned data + summary + quality report to Excel."""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        # Sheet 1: Clean Data
        df.to_excel(writer, sheet_name="Clean Data", index=False)

        # Sheet 2: Summary Statistics
        summary.to_excel(writer, sheet_name="Summary")

        # Sheet 3: Quality Report
        quality_df = pd.DataFrame([
            {"Metric": "Original rows",      "Value": quality["rows_original"]},
            {"Metric": "Clean rows",         "Value": quality["rows_clean"]},
            {"Metric": "Duplicates removed", "Value": quality["duplicates_removed"]},
            {"Metric": "Empty rows removed", "Value": quality["empty_rows_removed"]},
            {"Metric": "Total nulls left",   "Value": quality["null_total"]},
            {"Metric": "Report generated",   "Value": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
        ])
        quality_df.to_excel(writer, sheet_name="Quality Report", index=False)

    print(f"\n✅ Report saved to: {output_path}")


# ── Main ───────────────────────────────────────────────────────
def main():
    print("=" * 50)
    print("  CSV Report Automation Script")
    print(f"  Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)

    df              = load_csv_files(INPUT_FOLDER)
    df_clean, quality = clean_data(df)
    summary         = generate_summary(df_clean)
    output_path     = os.path.join(OUTPUT_FOLDER, REPORT_NAME)

    export_report(df_clean, summary, quality, output_path)

    print("\n🎉 Done! Check the output folder for your report.")
    print("=" * 50)


if __name__ == "__main__":
    main()
