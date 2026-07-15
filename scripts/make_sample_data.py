from pathlib import Path
import pandas as pd

RAW_DIR = Path("data/raw")
SAMPLE_DIR = Path("data/sample")
SAMPLE_DIR.mkdir(parents=True, exist_ok=True)

files = [
    "01教學實踐研究計畫核定結果_110-115.xlsx",
    "02大專學生研究計畫_110-115.xlsx",
    "03國科會學術補助計劃_110-115.xlsx",
    "05_USR_plan.xlsx",
	"06_tcsaward_2021_2025_universities.xlsx"
]

for file_name in files:
    input_path = RAW_DIR / file_name
    if not input_path.exists():
        print(f"找不到檔案：{input_path}")
        continue

    df = pd.read_excel(input_path, engine="openpyxl")
    sample_df = df.head(100)

    output_path = SAMPLE_DIR / file_name
    sample_df.to_excel(output_path, index=False, engine="openpyxl")

    print(f"已建立 sample：{output_path}")