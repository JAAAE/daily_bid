import streamlit as st
import pandas as pd
import os
import plotly.express as px

# 網頁基礎設定
st.set_page_config(page_title="政府電子採購網決標觀測站", layout="wide")

EXCEL_PATH = os.path.join('data', '採購網_決標彙整.xlsx')

# 你的 14 個標準關鍵字清單（必須與 Excel 欄位完全一致）
KEYWORDS = ["測繪", "空間資訊", "測量", "製圖", "圖資", "地圖", "地形", "測製", "地理資訊", "監審", "光達", "點雲", "模型", "建模"]

@st.cache_data(ttl=1800)  # 快取 30 分鐘
def load_data():
    if os.path.exists(EXCEL_PATH):
        try:
            # 1. 讀取全部資料
            df = pd.read_excel(EXCEL_PATH, sheet_name='全部彙整')
            
            # 2. 清洗與格式化日期
            df['日期'] = df['日期'].astype(str).str.strip()
            
            # 3. 清洗預算金額為數值
            df['預算'] = pd.to_numeric(df['預算'], errors='coerce').fillna(0)
            
            # 4. 💡 針對你提供的關鍵字結構進行強制轉型（消滅 0.0 / 1.0 浮點數）
            all_target_cols = KEYWORDS + ['關鍵字總計']
            for col in all_target_cols:
                if col in df.columns:
                    # 強制轉為數值，空值補 0，最後轉為標準整數型態
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
            return df
        except Exception as e:
            st.error(f"讀取 Excel 失敗，請檢查欄位格式。錯誤訊息: {e}")
            return None
    return None

# --- 網頁標頭 ---
st.title("🌐 空間資訊與測繪標案 決標觀測站")
st.caption("數據來源：政府電子採購網 (Openfun API 每日自動同步)")

df = load_data()

if df is not None:
    # --- 💡 動態取得現有 Excel 內相符的關鍵字欄位清單 ---
    keyword_cols = [col for col in df.columns if col in KEYWORDS]

    # --- 側邊欄篩選器 ---
    st.sidebar.header("🔍 篩選條件")
    
    # 1. 區域篩選
    all_regions = ["全部"] + sorted(list(df['區域'].dropna().unique()))
    selected_region = st.sidebar.selectbox("選擇區域", all_regions)
    
    # 2. 關鍵字篩選 (會連動影響資料與圖表)
    selected_keyword = st.sidebar.selectbox("主要關鍵字篩選", ["全部"] + keyword_cols)

    # 資料過濾邏輯
    filtered_df = df.copy()
    if selected_region != "全部":
        filtered_df = filtered_df[filtered_df['區域'] == selected_region]
    if selected_keyword != "全部":
        filtered_df = filtered_df[filtered_df[selected_keyword] == 1]

    # --- 頂部數據指標看板 ---
    total_tenders = len(filtered_df)
    latest_date_str = df['日期'].max()
    
    # 轉化日期格式 20260515 -> 2026-05-15
    if len(str(latest_date_str)) == 8:
        formatted_date = f"{latest_date_str[:4]}-{latest_date_str[4:6]}-{latest_date_str[6:]}"
    else:
        formatted_date = str(latest_date_str)

    # 計算總預算（萬元）
    total_budget_wan = filtered_df['預算'].sum() / 10000

    col1, col2, col3 = st.columns(3)
    col1.metric("當前篩選標案量", f"{total_tenders} 件", help="符合目前篩選條件的標案總數量")
    col2.metric("總決標預算規模", f"{total_budget_wan:,.0f} 萬元", help="當前篩選標案的預算加總")
    col3.metric("資料更新至 (昨日)", formatted_date)

    st.markdown("---")

    # --- 📊 圖表分析區 ---
    st.subheader("📊 統計趨勢與分佈")
    c1, c2 = st.columns(2)

    with c1:
        # 區域標案分佈圓餅圖
        region_counts = filtered_df['區域'].value_counts().reset_index()
        region_counts.columns = ['區域', '標案數量']
        fig_pie = px.pie(region_counts, values='標案數量', names='區域', title="地區標案比例", hole=0.4,
                         color_discrete_sequence=px.colors.qualitative.Pastel)
        st.plotly_chart(fig_pie, use_container_width=True)

    with c2:
        # 關鍵字熱度排行條形圖 (會直接加總你那些 0 與 1 的純整數欄位)
        if keyword_cols:
            kw_sums = filtered_df[keyword_cols].sum().reset_index()
            kw_sums.columns = ['關鍵字', '出現次數']
            kw_sums = kw_sums.sort_values(by='出現次數', ascending=True)
            fig_bar = px.bar(kw_sums, x='出現次數', y='關鍵字', orientation='h', title='關鍵字觸發熱度排行',
                             color='出現次數