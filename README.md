# Taiwan Higher Education Research & USR Open Data Pipeline

這是一個以 Python 建立的臺灣高等教育研究補助與大學社會責任資料整理專案。

本專案蒐集與整理多個公開資料來源，包括：

- 教學實踐研究計畫核定結果
- 國科會大專學生研究計畫
- 國科會學術補助計畫
- 教育部 USR 大學社會責任年報
- TCSA 永續獎歷屆榜單

## 專案目標

公開資料常分散在不同網站、不同格式與不同查詢系統中。本專案嘗試使用 Python 將這些資料整理成結構化資料，方便後續分析、比較與視覺化。

## 使用技術

- Python
- pandas
- requests
- BeautifulSoup
- lxml
- openpyxl
- Streamlit
- GitHub

## 專案結構

```text
.
├── data/
│   ├── raw/
│   ├── processed/
│   └── sample/
├── scripts/
├── notebooks/
├── dashboard/
└── docs/
```

## Live Dashboard
https://tw-higher-ed-research-usr-open-data-wpxzmsv8ocpnu3jcmmwldl.streamlit.app/
