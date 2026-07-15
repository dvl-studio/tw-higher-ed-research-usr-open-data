from pathlib import Path
import pandas as pd

DATA_DIR = Path("data/sample")
DOCS_DIR = Path("docs")
DOCS_DIR.mkdir(exist_ok=True)

files = {
    "教學實踐研究計畫": "01教學實踐研究計畫核定結果_110-115.xlsx",
    "大專學生研究計畫": "02大專學生研究計畫_110-115.xlsx",
    "國科會學術補助": "03國科會學術補助計劃_110-115.xlsx",
    "USR 年報": "05_USR_plan.xlsx",
	"臺灣永續獎":"06_tcsaward_2021_2025_universities.xlsx"
}

report_lines = []
report_lines.append("# Data Quality Report\n")

for name, file_name in files.items():
    path = DATA_DIR / file_name
    df = pd.read_excel(path, engine="openpyxl")

    report_lines.append(f"## {name}\n")
    report_lines.append(f"- Rows: {len(df)}")
    report_lines.append(f"- Columns: {len(df.columns)}")
    report_lines.append(f"- Column names: {', '.join(df.columns.astype(str))}")
    report_lines.append("\n### Missing Values\n")

    missing = df.isna().sum()
    for col, count in missing.items():
        report_lines.append(f"- {col}: {count}")

    report_lines.append("\n")

output_path = DOCS_DIR / "data_quality_report.md"
output_path.write_text("\n".join(report_lines), encoding="utf-8")

print(f"已產生：{output_path}")