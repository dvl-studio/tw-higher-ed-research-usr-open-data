from pathlib import Path
import pandas as pd

DATA_DIR = Path("../data/sample")

files = {
    "教學實踐研究計畫": "01教學實踐研究計畫核定結果_110-115.xlsx",
    "大專學生研究計畫": "02大專學生研究計畫_110-115.xlsx",
    "國科會學術補助": "03國科會學術補助計劃_110-115.xlsx",
    "USR 年報": "05_USR_plan.xlsx",
	"臺灣永續獎":"06_tcsaward_2021_2025_universities.xlsx"
}

for dataset_name, file_name in files.items():
    path = DATA_DIR / file_name
    df = pd.read_excel(path, engine="openpyxl")

    print("=" * 50)
    print(dataset_name)
    print("資料筆數：", len(df))
    print("欄位：", list(df.columns))
    print(df.head())