import json
import datetime
from pathlib import Path

import openpyxl


def load_sidings(ws):
    sidings = []
    # Expect header row: Subdivision, Siding, Start MP, End MP, Length (ft)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        subdivision, name, start_mp, end_mp, length_ft = row
        sidings.append(
            {
                "subdivision": str(subdivision).strip() if subdivision else "",
                "name": str(name).strip() if name else "",
                "start_mp": float(start_mp) if start_mp is not None else None,
                "end_mp": float(end_mp) if end_mp is not None else None,
                "length_ft": int(length_ft) if length_ft is not None else None,
            }
        )
    return sidings


def main():
    xlsx_path = Path("RAILCORE_Master_Sidings.xlsx")
    if not xlsx_path.exists():
        raise SystemExit(f"Missing file: {xlsx_path}")

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    # For now we only have a "Sidings" sheet. Later we can add Crossings, Track, etc.
    if "Sidings" not in wb.sheetnames:
        raise SystemExit("Workbook must contain a 'Sidings' sheet.")

    sidings_ws = wb["Sidings"]
    sidings = load_sidings(sidings_ws)

    infra = {
        "version": datetime.datetime.utcnow().isoformat(timespec="seconds") + "Z",
        "source_file": str(xlsx_path),
        "data": {
            "sidings": sidings,
            # placeholders for the future:
            "crossings": [],
            "track": [],
            "detectors": [],
        },
    }

    out_path = Path("railcore_infra.json")
    out_path.write_text(json.dumps(infra, indent=2), encoding="utf-8")
    print(f"Wrote {out_path} with {len(sidings)} sidings.")


if __name__ == "__main__":
    main()
