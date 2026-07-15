from __future__ import annotations

import re
import time
from pathlib import Path

import pandas as pd
import requests
import urllib3
from bs4 import BeautifulSoup
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


BASE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BASE_DIR.parent
START_YEAR = "110"
END_YEAR = "115"
OUTPUT_FILE = PROJECT_ROOT / "data" / "raw" / f"01教學實踐研究計畫核定結果_{START_YEAR}-{END_YEAR}.xlsx"
BASE_URL = "https://tpr.moe.edu.tw/plan/result"

QUERY_PARAMS = {
    "startYear": START_YEAR,
    "endYear": END_YEAR,
    "type": "",
    "period": "",
    "schoolType": "",
    "school": "",
    "keyword": "",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/126.0 Safari/537.36"
    ),
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
}


VERIFY_SSL = False

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


EXPECTED_COLUMNS = ["年度", "學門", "期程", "學校", "系所", "姓名", "職稱", "計畫名稱"]


def fetch_page(session: requests.Session, page: int) -> str:
    params = {"page": page, **QUERY_PARAMS}
    last_error: Exception | None = None

    for attempt in range(1, 4):
        try:
            response = session.get(BASE_URL, params=params, headers=HEADERS, timeout=30, verify=VERIFY_SSL,)
            response.raise_for_status()
            response.encoding = "utf-8"
            return response.text
        except requests.RequestException as exc:
            last_error = exc
            wait_seconds = attempt * 2
            print(f"[warn] page {page} failed on attempt {attempt}; retrying in {wait_seconds}s")
            time.sleep(wait_seconds)

    raise RuntimeError(f"Failed to fetch page {page}: {last_error}")


def parse_total_pages(html: str) -> int:
    match = re.search(r"var\s+totalPages\s*=\s*(\d+)", html)
    if not match:
        raise ValueError("Could not find totalPages in the first page.")
    return int(match.group(1))


def parse_result_count(html: str) -> int | None:
    text = BeautifulSoup(html, "html.parser").get_text(" ", strip=True)
    match = re.search(r"共\s*([\d,]+)\s*筆", text)
    return int(match.group(1).replace(",", "")) if match else None


def parse_table_rows(html: str) -> list[dict[str, str]]:
    soup = BeautifulSoup(html, "html.parser")
    rows: list[dict[str, str]] = []

    for tr in soup.select("tbody tr"):
        cells = tr.find_all("td")
        if not cells:
            continue

        row: dict[str, str] = {}
        for index, cell in enumerate(cells):
            column = cell.get("thead") or (EXPECTED_COLUMNS[index] if index < len(EXPECTED_COLUMNS) else f"欄位{index + 1}")
            value = cell.get_text(" ", strip=True)
            row[column] = re.sub(r"\s+", " ", value)

        if row:
            rows.append(row)

    return rows


def download_results() -> pd.DataFrame:
    session = requests.Session()

    first_html = fetch_page(session, 1)
    total_pages = parse_total_pages(first_html)
    expected_count = parse_result_count(first_html)

    all_rows = parse_table_rows(first_html)
    print(f"[info] page 1/{total_pages}: {len(all_rows)} rows")

    for page in range(2, total_pages + 1):
        html = fetch_page(session, page)
        page_rows = parse_table_rows(html)
        all_rows.extend(page_rows)
        print(f"[info] page {page}/{total_pages}: +{len(page_rows)} rows, total {len(all_rows)}")
        time.sleep(0.1)

    df = pd.DataFrame(all_rows)
    for column in EXPECTED_COLUMNS:
        if column not in df.columns:
            df[column] = ""
    df = df[EXPECTED_COLUMNS]

    if expected_count is not None and len(df) != expected_count:
        raise ValueError(f"Expected {expected_count} rows from the website, but parsed {len(df)} rows.")

    return df


def clean_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    return df.map(lambda value: ILLEGAL_CHARACTERS_RE.sub("", value) if isinstance(value, str) else value)


def main() -> None:
    df = download_results()
    df = clean_for_excel(df)
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    print(f"[success] saved {len(df)} rows to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
