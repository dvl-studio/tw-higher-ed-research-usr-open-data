from __future__ import annotations

import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
RAW_DIR = PROJECT_ROOT / "data" / "raw"

RAW_DIR.mkdir(parents=True, exist_ok=True)


def run_command(command: list[str]) -> None:
    print("=" * 80)
    print("執行：")
    print(" ".join(command))
    print("=" * 80)

    result = subprocess.run(command, cwd=PROJECT_ROOT)

    if result.returncode != 0:
        raise RuntimeError(f"執行失敗：{' '.join(command)}")


def main() -> None:
    python = sys.executable

    jobs = [
        {
            "name": "教學實踐研究計畫核定結果",
            "command": [
                python,
                str(SCRIPTS_DIR / "01_download_tpr_results.py"),
            ],
            "expected_output": RAW_DIR / "01教學實踐研究計畫核定結果_110-115.xlsx",
        },
        {
            "name": "大專學生研究計畫",
            "command": [
                python,
                str(SCRIPTS_DIR / "02_scrape_nstc_student_awards.py"),
                "--start-year",
                "110",
                "--end-year",
                "115",
                "--output",
                str(RAW_DIR / "02大專學生研究計畫_110-115.xlsx"),
            ],
            "expected_output": RAW_DIR / "02大專學生研究計畫_110-115.xlsx",
        },
        {
            "name": "國科會學術補助計畫",
            "command": [
                python,
                str(SCRIPTS_DIR / "03_scrape_nstc_academic_subsidy.py"),
                "--start-year",
                "110",
                "--end-year",
                "115",
                "--output",
                str(RAW_DIR / "03國科會學術補助計劃_110-115.xlsx"),
            ],
            "expected_output": RAW_DIR / "03國科會學術補助計劃_110-115.xlsx",
        },
        {
            "name": "TCSA 永續獎歷屆榜單",
            "command": [
                python,
                str(SCRIPTS_DIR / "04_scrape_tcsaward_history.py"),
                "--years",
                "2021",
                "2022",
                "2023",
                "2024",
                "2025",
                "--output",
                str(RAW_DIR / "04_tcsaward_2021_2025.xlsx"),
            ],
            "expected_output": RAW_DIR / "04_tcsaward_2021_2025.xlsx",
        },

	{
            "name": "TCSA 永續獎大學資料篩選",
            "command": [
                python,
                str(SCRIPTS_DIR / "06_filter_tcsaward_universities.py"),
                "--input",
                str(RAW_DIR / "04_tcsaward_2021_2025.xlsx"),
                "--output",
                str(RAW_DIR / "06_tcsaward_2021_2025_universities.xlsx"),
            ],
            "expected_output": RAW_DIR / "06_tcsaward_2021_2025_universities.xlsx",
        },

        {
            "name": "USR 年報資料",
            "command": [
                python,
                str(SCRIPTS_DIR / "05_scrape_usr_plan.py"),
            ],
            "expected_output": RAW_DIR / "05_USR_plan.xlsx",
        },
    ]

    for job in jobs:
        print(f"\n開始執行：{job['name']}")
        run_command(job["command"])

    print("\n全部爬蟲執行完成。")
    print(f"請檢查輸出資料夾：{RAW_DIR}")


if __name__ == "__main__":
    main()