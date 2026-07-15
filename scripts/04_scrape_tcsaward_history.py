from __future__ import annotations

import argparse
import re
import time
from pathlib import Path

import pandas as pd
import requests
import urllib3
from bs4 import BeautifulSoup


BASE_URL = "https://tcsaward.org.tw/tw/award/award_history"
DEFAULT_YEARS = list(range(2021, 2026))
DEFAULT_OUTPUT = Path(__file__).with_name("tcsaward_2021_2025.xlsx")
COLUMNS = ["年度", "企業名稱", "獎項分類", "獎項名稱"]


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def fetch_year(session: requests.Session, year: int, verify_ssl: bool) -> list[dict[str, str]]:
    response = session.get(
        BASE_URL,
        params={"ac_year": str(year)},
        timeout=60,
        verify=verify_ssl,
    )
    response.raise_for_status()

    soup = BeautifulSoup(response.content, "html.parser")
    records: list[dict[str, str]] = []

    for item in soup.select("li.member-page-list"):
        values = [clean_text(span.get_text(" ", strip=True)) for span in item.select("span.info-value")]
        if len(values) < 4:
            continue

        records.append(
            {
                "年度": values[0].replace("年", ""),
                "企業名稱": values[1],
                "獎項分類": values[2],
                "獎項名稱": values[3],
            }
        )

    return records


def scrape(years: list[int], verify_ssl: bool, delay_seconds: float) -> pd.DataFrame:
    if not verify_ssl:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/126.0 Safari/537.36"
            )
        }
    )

    all_records: list[dict[str, str]] = []
    for index, year in enumerate(years):
        records = fetch_year(session, year, verify_ssl=verify_ssl)
        print(f"{year}: {len(records)} records")
        all_records.extend(records)

        if delay_seconds > 0 and index < len(years) - 1:
            time.sleep(delay_seconds)

    return pd.DataFrame(all_records, columns=COLUMNS)


def write_excel(df: pd.DataFrame, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="TCSA歷屆榜單")
        worksheet = writer.sheets["TCSA歷屆榜單"]
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
    parser = argparse.ArgumentParser(description="Scrape TCSA award history data to Excel.")
    parser.add_argument(
        "--years",
        nargs="+",
        type=int,
        default=DEFAULT_YEARS,
        help="Years to scrape. Default: 2021 2022 2023 2024 2025",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Output xlsx path. Default: {DEFAULT_OUTPUT}",
    )
    parser.add_argument(
        "--verify-ssl",
        action="store_true",
        help="Verify SSL certificate. Omit this if your local certificate store rejects the site certificate.",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.5,
        help="Delay seconds between yearly requests. Default: 0.5",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    years = sorted(set(args.years))
    df = scrape(years, verify_ssl=args.verify_ssl, delay_seconds=args.delay)
    write_excel(df, args.output)
    print(f"Saved {len(df)} rows to {args.output}")


if __name__ == "__main__":
    main()
