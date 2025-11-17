import pandas as pd
import json
import os

INPUT_FILE = "data/public/RAILCORE_Infrastructure_GitHubMaster.xlsx"
OUTPUT_DIR = "data/json"

SHEETS = {
    "Sidings": "sidings.json",
    "Crossings": "crossings.json",
    "Track": "track.json",
}

def clean_value(v):
    if pd.isna(v):
        return ""
    return v

def convert():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    wb = pd.ExcelFile(INPUT_FILE)

    for sheet, output_name in SHEETS.items():
        if sheet not in wb.sheet_names:
            print(f"Skipping missing sheet: {sheet}")
            continue

        df = wb.parse(sheet).fillna("")

        records = [
            {col: clean_value(row[col]) for col in df.columns}
            for _, row in df.iterrows()
        ]

        with open(f"{OUTPUT_DIR}/{output_name}", "w", encoding="utf-8") as f:
            json.dump(records, f, indent=2)

        print(f"Generated {output_name}")

if __name__ == "__main__":
    convert()
