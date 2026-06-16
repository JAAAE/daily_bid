import streamlit as st
import pandas as pd
import os
import plotly.express as px

st.set_page_config(page_title="政府電子採購網決標觀測站", layout="wide")

EXCEL_PATH = os.path.join('data', '採購網_決標彙整.xlsx')

# 定義固定的關鍵字清單，與爬蟲保持絕對一致，最安全
KEYWORDS = ["測繪", "空間資訊", "測量", "製圖", "圖資", "地圖", "地形", "測製", "地理資訊", "監審", "光達", "點雲", "模型", "建模"]

@st.cache_data(ttl=1800)  # 快取調整為 30 分鐘，兼顧效能與即時性
def load_data():
    if os.path.exists(EXCEL_PATH):
        try:
            # 讀取全部資料
            df = pd.read_excel(EXCEL_PATH, sheet_name='全部彙整')
            df['日期'] = df['日期'].astype(str).str.strip()
            # 確保金額欄位轉為數值，方便做儀表板統計
            df['預算'] = pd.to_numeric(df['預算'], errors='coerce').fillna(0)
            return df
        except Exception as e:
            st.error(f"讀取 Excel 失敗: {e}")
            return None
    return None

st.title("🌐 空間資訊與測繪標案 決標觀測站")
st.caption("數據來源：政府電子採購網 (Openfun API 每日自動同步)")

df = load_data()

if df is not None:
    # --- 💡 動態取得關鍵字欄位清單（避免寫死索引） ---
    # 只要 DataFrame 欄位名稱在 KEYWORDS 清單內就抓出來
    keyword_cols = [col for col in df.columns if col in KEYWORDS]

    # --- 側邊欄篩選器 ---
    st.sidebar.header("🔍 篩選條件")
    
    # 區域篩選
    all_regions = ["全部"] + sorted(list(df['區域'].dropna().unique()))
    selected_region = st.sidebar.selectbox("選擇區域", all_regions)
    
    # 關鍵字篩選
    selected_keyword = st.sidebar.selectbox("主要關鍵字篩選", ["全部"] + keyword_cols)

    # 資料過濾邏輯
    filtered_df = df.copy()
    if selected_region != "全部":
        filtered_df = filtered_df[filtered_df['區域'] == selected_region]
    if selected_keyword != "全部":
        filtered_df = filtered_df[filtered_df[selected_keyword] == 1]

    # --- 頂部指標看板 ---
    total_tenders = len(filtered_df)
    latest_date_str = df['日期'].max()
    
    # 格式化日期顯示 (例如 20260515 -> 2026-05-15)
    if len(latest_date_str) == 8:
        formatted_date = f"{latest_date_str[:4]}-{latest_date_str[4:6]}-{latest_date_str[6:]}"
    else:
        formatted_date = latest_date_str

    # 統計當前篩選條件下的總預算 (萬)
    total_budget_wan = filtered_df['預算'].sum() / 10000

    col1, col2, col3 = st.columns(3)
    col1.metric("當前篩選標案量", f"{total_tenders} 件", help="符合左側篩選條件的標案總數")
    col2.metric("總決標預算規模", f"{total_budget_wan:,.0f} 萬元", help="當前篩選標案的預算金額加總")
    col3.metric("資料更新至 (昨日)", formatted_date)

    st.markdown("---")

    # --- 圖表分析區 ---
    st.subheader("📊 統計趨勢與分佈")
    c1, c2 = st.columns(2)

    with c1:
        # 區域標案分佈
        region_counts = filtered_df['區域'].value_counts().reset_index()
        region_counts.columns = ['區域', '標案數量']
        fig_pie = px.pie(region_counts, values='標案數量', names='區域', title="地區標案比例", hole=0.4,
                         color_discrete_sequence=px.colors.qualitative.Pastel)
        st.plotly_chart(fig_pie, use_container_width=True)

    with c2:
        # 關鍵字熱度排行
        if keyword_cols:
            kw_sums = filtered_df[keyword_cols].sum().reset_index()
            kw_sums.columns = ['關鍵字', '出現次數']
            kw_sums = kw_sums.sort_values(by='出現次數', ascending=True)
            fig_bar = px.bar(kw_sums, x='出現次數', y='關鍵字', orientation='h', title='關鍵字觸發熱度排行',
                             color='出現次數', color_continuous_scale='Blues')
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("無關鍵字數據可供繪圖")

    # --- 詳細資料表格 ---
    st.subheader("📋 標案明細清單")
    
    # 重新梳理要顯示的基本欄位，隱藏獨立的 0/1 欄位
    display_cols = ['日期', '機關名稱', '地點', '區域', '標案名稱', '預算', '成果連結', '關鍵字總計']
    
    # 確保顯示的欄位在 df 中都存在
    available_display_cols = [c for c in display_cols if c in filtered_df.columns]
    
    st.dataframe(
        filtered_df[available_display_cols],
        column_config={
            "日期": st.column_config.TextColumn("決標日期"),
            "預算": st.column_config.NumberColumn("預算金額 (元)", format="$%,d"),
            "成果連結": st.column_config.LinkColumn("標案詳細連結", display_text="檢視公告")
        },
        use_container_width=True,
        hide_index=True
    )
else:
    st.warning("⚠️ 找不到 `data/採購網_決擺彙整.xlsx` 檔案，請確認爬蟲已成功執行並將檔案推上 GitHub。")