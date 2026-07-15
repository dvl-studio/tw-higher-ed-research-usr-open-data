from __future__ import annotations

import argparse
import re
import subprocess
import time
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable
from urllib.parse import urlencode
import shutil
import pandas as pd
from lxml import html as lxml_html


URL = "https://wsts.nstc.gov.tw/STSWeb/Award/AwardMultiQuery.aspx"
OUTPUT_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT_XLSX = OUTPUT_DIR / "03國科會學術補助計劃_110-115.xlsx"
COOKIE_FILE = OUTPUT_DIR / "nstc_sts_award_cookie.txt"
POST_FILE = OUTPUT_DIR / "nstc_sts_award_post_body.txt"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36"
)
PAGE_SIZE = "200"
CATEGORY_EVENT_TARGET = "dtlItem$ctl00$btnItem"
COLUMNS = [
    "計畫年度",
    "主持人姓名",
    "執行機關",
    "計畫名稱",
    "執行起迄",
    "總核定金額",
]


@dataclass
class SelectField:
    name: str
    options: list[dict[str, str]]


class FormParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.inputs: list[dict[str, str]] = []
        self.selects: list[SelectField] = []
        self._select_name: str | None = None
        self._options: list[dict[str, str]] = []

    def handle_starttag(self, tag: str, attrs: Iterable[tuple[str, str | None]]) -> None:
        attr = {key: value or "" for key, value in attrs}
        if tag == "input":
            self.inputs.append(attr)
        elif tag == "select":
            self._select_name = attr.get("name")
            self._options = []
        elif tag == "option" and self._select_name:
            self._options.append(attr)

    def handle_endtag(self, tag: str) -> None:
        if tag == "select" and self._select_name:
            self.selects.append(SelectField(self._select_name, self._options))
            self._select_name = None
            self._options = []



def fetch_page(post_fields: dict[str, str] | None = None) -> str:
    curl_command = shutil.which("curl") or shutil.which("curl.exe")

    if curl_command is None:
        raise RuntimeError("找不到 curl。請先確認你的系統已安裝 curl。")

    cmd = [
        curl_command,
        "-sS",
        "-L",
        "--compressed",
        "--retry",
        "3",
        "--retry-delay",
        "2",
        "-A",
        USER_AGENT,
        "-c",
        str(COOKIE_FILE),
        "-b",
        str(COOKIE_FILE),
    ]
    if post_fields is not None:
        POST_FILE.write_text(urlencode(post_fields), encoding="utf-8")
        cmd.extend(
            [
                "-H",
                "Content-Type: application/x-www-form-urlencoded",
                "--data-binary",
                f"@{POST_FILE}",
            ]
        )
    cmd.append(URL)

    proc = subprocess.run(cmd, capture_output=True, check=True)
    return proc.stdout.decode("utf-8", errors="replace")


def parse_form_fields(page_html: str) -> tuple[dict[str, str], list[SelectField]]:
    parser = FormParser()
    parser.feed(page_html)

    fields: dict[str, str] = {}
    for item in parser.inputs:
        name = item.get("name")
        input_type = item.get("type", "").lower()
        if not name or input_type == "image":
            continue
        if input_type in {"radio", "checkbox"} and "checked" not in item:
            continue
        fields[name] = item.get("value", "")

    for select in parser.selects:
        if not select.name or select.name in fields:
            continue
        selected = next(
            (option.get("value", "") for option in select.options if "selected" in option),
            None,
        )
        if selected is None and select.options:
            selected = select.options[0].get("value", "")
        if selected is not None:
            fields[select.name] = selected

    return fields, parser.selects


def select_academic_subsidy_category() -> str:
    page_html = fetch_page()
    fields, _ = parse_form_fields(page_html)
    fields["__EVENTTARGET"] = CATEGORY_EVENT_TARGET
    fields["__EVENTARGUMENT"] = ""
    fields["__LASTFOCUS"] = ""
    return fetch_page(fields)


def query_year_range(page_html: str, start_year: str, end_year: str) -> str:
    fields, _ = parse_form_fields(page_html)
    fields.update(
        {
            "__EVENTTARGET": "",
            "__EVENTARGUMENT": "",
            "__LASTFOCUS": "",
            "wUctlAwardQueryPage$repQuery$ctl01$ddlYRst": start_year,
            "wUctlAwardQueryPage$repQuery$ctl01$ddlYRend": end_year,
            "wUctlAwardQueryPage$ddlPageSize": PAGE_SIZE,
            "wUctlAwardQueryPage$hidPage": "",
            "wUctlAwardQueryPage$btnQuery.x": "20",
            "wUctlAwardQueryPage$btnQuery.y": "10",
        }
    )
    return fetch_page(fields)


def get_page_select(selects: list[SelectField]) -> SelectField:
    for select in selects:
        if "grdResult" in select.name and "ddlPage" in select.name:
            return select
    raise RuntimeError("找不到結果分頁下拉選單。")


def go_to_result_page(page_html: str, page_index: int) -> str:
    fields, selects = parse_form_fields(page_html)
    page_select = get_page_select(selects)
    page_value = str(page_index)
    fields[page_select.name] = page_value
    fields["wUctlAwardQueryPage$hidPage"] = page_value
    fields["__EVENTTARGET"] = page_select.name
    fields["__EVENTARGUMENT"] = ""
    fields.pop("wUctlAwardQueryPage$btnQuery.x", None)
    fields.pop("wUctlAwardQueryPage$btnQuery.y", None)
    return fetch_page(fields)


def text_content(node) -> str:
    return " ".join(part.strip() for part in node.xpath(".//text()") if part.strip())


def content_value(content_cell, field_name: str) -> str:
    label_nodes = content_cell.xpath(f'.//span[normalize-space()="{field_name}："]')
    if not label_nodes:
        return ""
    label_id = label_nodes[0].get("id", "")
    value_id = label_id[:-1] + "c" if label_id.endswith("t") else ""
    if value_id:
        value_nodes = content_cell.xpath(f'.//*[@id="{value_id}"]')
        if value_nodes:
            return text_content(value_nodes[0])

    pieces = []
    after_label = False
    for child in content_cell:
        if child is label_nodes[0]:
            after_label = True
            continue
        if after_label:
            if child.tag == "br":
                break
            pieces.append(text_content(child))
    return " ".join(piece for piece in pieces if piece).strip()


def parse_rows(page_html: str, years: set[str]) -> list[dict[str, str]]:
    doc = lxml_html.fromstring(page_html)
    rows: list[dict[str, str]] = []
    for tr in doc.xpath('//table[@id="wUctlAwardQueryPage_grdResult"]/tr'):
        cells = tr.xpath("./td")
        if len(cells) != 4:
            continue

        year = text_content(cells[0])
        if year not in years:
            continue

        amount = re.sub(r"\D", "", content_value(cells[3], "總核定金額"))
        rows.append(
            {
                "計畫年度": year,
                "主持人姓名": text_content(cells[1]),
                "執行機關": text_content(cells[2]),
                "計畫名稱": content_value(cells[3], "計畫名稱"),
                "執行起迄": content_value(cells[3], "執行起迄"),
                "總核定金額": amount,
            }
        )
    return rows


def total_pages(page_html: str) -> int:
    _, selects = parse_form_fields(page_html)
    page_select = get_page_select(selects)
    return len(page_select.options)


def save_excel(rows: list[dict[str, str]], output_xlsx: Path) -> None:
    df = pd.DataFrame(rows, columns=COLUMNS)
    df["計畫年度"] = pd.to_numeric(df["計畫年度"], errors="coerce").astype("Int64")
    df["總核定金額"] = pd.to_numeric(df["總核定金額"], errors="coerce").astype("Int64")
    df = df.drop_duplicates(subset=COLUMNS).sort_values(["計畫年度", "執行機關", "主持人姓名"])
    df.to_excel(output_xlsx, index=False)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="下載 NSTC 國科會學術補助計劃查詢結果。")
    parser.add_argument("--start-year", default="110", help="起始年度，例如 110")
    parser.add_argument("--end-year", default="115", help="結束年度，例如 115")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT_XLSX), help="輸出的 xlsx 檔案路徑")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    start_year = int(args.start_year)
    end_year = int(args.end_year)
    if start_year > end_year:
        raise ValueError("--start-year 不可大於 --end-year")

    years = {str(year) for year in range(start_year, end_year + 1)}
    output_xlsx = Path(args.output)
    output_xlsx.parent.mkdir(parents=True, exist_ok=True)

    for temp_file in (COOKIE_FILE, POST_FILE):
        if temp_file.exists():
            temp_file.unlink()

    print("開啟國科會學術補助計劃查詢頁...")
    category_html = select_academic_subsidy_category()
    first_page_html = query_year_range(category_html, str(start_year), str(end_year))
    pages = total_pages(first_page_html)
    rows = parse_rows(first_page_html, years)
    print(f"第 1/{pages} 頁：累計 {len(rows)} 筆")

    current_html = first_page_html
    for page_index in range(1, pages):
        time.sleep(0.1)
        current_html = go_to_result_page(current_html, page_index)
        rows.extend(parse_rows(current_html, years))
        print(f"第 {page_index + 1}/{pages} 頁：累計 {len(rows)} 筆")

    save_excel(rows, output_xlsx)
    for temp_file in (COOKIE_FILE, POST_FILE):
        if temp_file.exists():
            temp_file.unlink()
    print(f"完成：{output_xlsx}")


if __name__ == "__main__":
    main()
