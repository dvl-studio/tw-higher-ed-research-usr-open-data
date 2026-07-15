from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = BASE_DIR / "tcsaward_2021_2025.xlsx"
DEFAULT_OUTPUT = BASE_DIR / "tcsaward_2021_2025_大學.xlsx"
REQUIRED_COLUMNS = ["年度", "企業名稱", "獎項分類", "獎項名稱"]


def filter_universities(input_path: Path) -> pd.DataFrame:
    df = pd.read_excel(input_path, dtype=str)
    missing_columns = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing columns: {', '.join(missing_columns)}")

    cleaned = df[REQUIRED_COLUMNS].copy()
    company_names = cleaned["企業名稱"].fillna("").astype(str).str.strip()
    university_mask = company_names.str.endswith("大學")

    result = cleaned.loc[university_mask].copy()
    result["企業名稱"] = company_names.loc[university_mask]
    return result.reset_index(drop=True)


def write_excel(df: pd.DataFrame, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="大學")
        worksheet = writer.sheets["大學"]
        widths = {
            "A": 10,
            "B": 42,
            "C": 44,
            "D": 58,
        }
        for column, width in widths.items():
            worksheet.column_dimensions[column].width = width
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Keep only TCSA records whose company name ends with 大學.")
    parser.add_argument(
        "--input",
        type=Path,
        default=DEFAULT_INPUT,
        help=f"Input xlsx path. Default: {DEFAULT_INPUT}",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Output xlsx path. Default: {DEFAULT_OUTPUT}",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    result = filter_universities(args.input)
    write_excel(result, args.output)
    print(f"Saved {len(result)} rows to {args.output}")


if __name__ == "__main__":
    main()
