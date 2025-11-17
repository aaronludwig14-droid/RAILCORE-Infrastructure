import pandas as pd
import json
import os

INPUT_FILE = "data/public/RAILCORE_Infrastructure_GitHubMaster.xlsx"
OUTPUT_DIR = "data/json"

# These are the sheet names expected inside the master workbook.
SHEETS = {
    "Sidings": "sidings.json",
    "Crossings": "crossings.json",
    "Track": "track.json"
}

def clean_value(v):
    """Converts NaN/None to a clean empty value."""
    if pd.isna(v):
        return ""
    return v

def df_to_records(df):
    """Convert DataFrame rows to dictionaries with clean values."""
    return [ {col: clean_value(row[col]) for col in df.columns} for idx, row in df.iterrows() ]

def convert_excel_to_json():
    # Ensure output directory exists
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print(f"Loading workbook: {INPUT_FILE}")

    try:
        workbook = pd.ExcelFile(INPUT_FILE)
    except Exception as e:
        print(f"ERROR: Could not open Excel file: {e}")
        return

    for sheet_name, output_file in SHEETS.items():
        print(f"Processing sheet: {sheet_name}")

        if sheet_name not in workbook.sheet_names:
            print(f"WARNING: Sheet '{sheet_name}' NOT found — skipping.")
            continue

        try:
            df = workbook.parse(sheet_name)
            df = df.fillna("")  # ensure no NaNs in final JSON

            records = df_to_records(df)

            output_path = os.path.join(OUTPUT_DIR, output_file)
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(records, f, indent=2)

            print(f"✔ Generated {output_file}")

        except Exception as e:
            print(f"ERROR processing sheet '{sheet_name}': {e}")

    print("Conversion complete.")


if __name__ == "__main__":
    convert_excel_to_json()
