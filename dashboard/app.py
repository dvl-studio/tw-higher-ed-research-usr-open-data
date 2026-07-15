from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(
    page_title="Taiwan Higher Education Research & USR Dashboard",
    layout="wide",
)

st.title("Taiwan Higher Education Research & USR Open Data Dashboard")
st.write("這是一個展示臺灣高教研究補助、教學實踐研究與 USR 年報資料的互動式儀表板。")

DATA_DIR = Path(__file__).resolve().parents[1] / "data" / "raw"

dataset_options = {
    "教學實踐研究計畫": "01教學實踐研究計畫核定結果_110-115.xlsx",
    "大專學生研究計畫": "02大專學生研究計畫_110-115.xlsx",
    "國科會學術補助計劃": "03國科會學術補助計劃_110-115.xlsx",
    "USR 計畫": "05_USR_plan.xlsx",
    "TCSA 台灣企業永續獎": "06_tcsaward_2021_2025_universities.xlsx"
}

selected_dataset = st.sidebar.selectbox(
    "選擇資料集",
    list(dataset_options.keys())
)

file_path = DATA_DIR / dataset_options[selected_dataset]
df = pd.read_excel(file_path, engine="openpyxl")

st.subheader(selected_dataset)
st.write(f"資料筆數：{len(df)}")
st.write("欄位：", list(df.columns))

st.dataframe(df, use_container_width=True)

st.subheader("年度分布")

year_column_candidates = ["年度", "計畫年度", "年份"]
year_column = None

for col in year_column_candidates:
    if col in df.columns:
        year_column = col
        break

if year_column:
    count_by_year = df.groupby(year_column).size().reset_index(name="筆數")
    fig = px.bar(
        count_by_year,
        x=year_column,
        y="筆數",
        title=f"{selected_dataset} 年度筆數分布",
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("這個資料集沒有找到年度欄位。")

st.subheader("關鍵字搜尋")

keyword = st.text_input("輸入關鍵字，例如 AI、USR、永續、地方創生")

if keyword:
    text_df = df.astype(str)
    mask = text_df.apply(
        lambda row: row.str.contains(keyword, case=False, na=False).any(),
        axis=1
    )
    result = df[mask]
    st.write(f"找到 {len(result)} 筆相關資料")
    st.dataframe(result, use_container_width=True)