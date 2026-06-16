import streamlit as st
import pandas as pd
import os
import plotly.express as px

st.set_page_config(page_title="政府電子採購網決標觀測站", layout="wide")

EXCEL_PATH = os.path.join('data', '採購網_決標彙整.xlsx')

@st.cache_data(ttl=3600)  # 快取 1 小時，避免頻繁讀取檔案
def load_data():
    if os.path.exists(EXCEL_PATH):
        # 讀取全部資料
        df = pd.read_excel(EXCEL_PATH, sheet_name='全部彙整')
        df['日期'] = df['日期'].astype(str)
        return df
    return None

st.title("🌐 空間資訊/測繪相關標案 決標觀測站")
st.caption("數據來源：政府電子採購網 (Openfun API 每日自動同步)")

df = load_data()

if df is not None:
    # --- 頂部指標看板 ---
    total_tenders = len(df)
    latest_date = df['日期'].max()
    
    col1, col2, col3 = st.columns(3)
    col1.metric("總觀測標案量", f"{total_tenders} 件")
    col2.metric("最新數據日期", f"{latest_date[:4]}-{latest_date[4:6]}-{latest_date[6:]}")
    col3.metric("監測關鍵字組", f"{len(df.columns) - 9} 組")

    st.markdown("---")

    # --- 側邊欄篩選器 ---
    st.sidebar.header("🔍 篩選條件")
    
    # 區域篩選
    all_regions = ["全部"] + list(df['區域'].unique())
    selected_region = st.sidebar.selectbox("選擇區域", all_regions)
    
    # 關鍵字篩選
    keyword_cols = df.columns[8:-1].tolist() # 自動取得關鍵字清單
    selected_keyword = st.sidebar.selectbox("主要關鍵字篩選", ["全部"] + keyword_cols)

    # 資料過濾邏輯
    filtered_df = df.copy()
    if selected_region != "全部":
        filtered_df = filtered_df[filtered_df['區域'] == selected_region]
    if selected_keyword != "全部":
        filtered_df = filtered_df[filtered_df[selected_keyword] == 1]

    # --- 圖表分析區 ---
    st.subheader("📊 統計趨勢與分佈")
    c1, c2 = st.columns(2)

    with c1:
        # 區域標案分佈
        region_counts = filtered_df['區域'].value_counts().reset_index()
        region_counts.columns = ['區域', '標案數量']
        fig_pie = px.pie(region_counts, values='標案數量', names='區域', title="地區標案比例", hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)

    with c2:
        # 關鍵字熱度排行
        kw_sums = filtered_df[keyword_cols].sum().reset_index()
        kw_sums.columns = ['關鍵字', '出現次數']
        kw_sums = kw_sums.sort_values(by='出現次數', ascending=True)
        fig_bar = px.bar(kw_sums, x='出現次數', y='關鍵字', orientation='h', title='關鍵字觸發熱度排行')
        st.plotly_chart(fig_bar, use_container_width=True)

    # --- 詳細資料表格 ---
    st.subheader("📋 標案明細清單")
    
    # 隱藏不必要的關鍵字 0/1 欄位，讓表格乾淨
    display_cols = ['日期', '機關名稱', '地點', '區域', '標案名稱', '預算', '成果連結', '關鍵字總計']
    st.dataframe(
        filtered_df[display_cols],
        column_config={
            "成果連結": st.column_config.LinkColumn("標案詳細連結")
        },
        use_container_width=True,
        hide_index=True
    )
else:
    st.warning("⚠️ 找不到 `data/採購網_決標彙整.xlsx` 檔案，請確認檔案已正確上傳至 GitHub 儲存庫。")