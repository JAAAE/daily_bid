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
            
            # 3. 💡 強力清洗金額欄位，防止舊資料包含國字、全形或金錢符號導致 pd.to_numeric 失敗變 0
            if '預算' in df.columns:
                df['預算'] = df['預算'].astype(str).str.replace('$', '', regex=False) \
                                                    .str.replace(',', '', regex=False) \
                                                    .str.replace('元', '', regex=False) \
                                                    .str.strip()
                df['預算'] = pd.to_numeric(df['預算'], errors='coerce').fillna(0)
            
            # 4. 強制將所有關鍵字欄位與「關鍵字總計」清洗回純整數 (0 或 1)
            all_target_cols = KEYWORDS + ['關鍵字總計']
            for col in all_target_cols:
                if col in df.columns:
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
    # --- 動態取得現有 Excel 內相符的關鍵字欄位清單 ---
    keyword_cols = [col for col in df.columns if col in KEYWORDS]

    # --- 側邊欄篩選器 ---
    st.sidebar.header("🔍 篩選條件")
    
    # 1. 區域篩選
    all_regions = ["全部"] + sorted(list(df['區域'].dropna().unique()))
    selected_region = st.sidebar.selectbox("選擇區域", all_regions)
    
    # 2. 關鍵字篩選
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
        # 關鍵字熱度排行條形圖
        if keyword_cols:
            kw_sums = filtered_df[keyword_cols].sum().reset_index()
            kw_sums.columns = ['關鍵字', '出現次數']
            kw_sums = kw_sums.sort_values(by='出現次數', ascending=True)
            
            fig_bar = px.bar(kw_sums, x='出現次數', y='關鍵字', orientation='h', title='關鍵字觸發熱度排行',
                             color='出現次數', color_continuous_scale='Blues')
            
            fig_bar.update_layout(xaxis=dict(tickformat="d"))
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("無關鍵字數據可供繪圖")

    # --- 📋 詳細資料表格 ---
    st.subheader("📋 標案明細清單")
    
    # 📌 全部欄位完全展開，不進行任何刪減或精簡
    base_front = ['日期', '機關名稱', '地點', '區域', '標案名稱', '預算']
    base_back = ['成果連結', '關鍵字總計']
    
    display_cols = base_front + keyword_cols + base_back
    available_display_cols = [c for c in display_cols if c in filtered_df.columns]
    
    # 建立格式化映射
    custom_configs = {
        "日期": st.column_config.TextColumn("決標日期"),
        "預算": st.column_config.NumberColumn("預算金額 (元)", format="$%,d"),
        "成果連結": st.column_config.LinkColumn("標案詳細連結", display_text="檢視公告"),
        "關鍵字總計": st.column_config.NumberColumn("關鍵字總計", format="%d")
    }
    
    # 📌 將 14 個獨立欄位全部格式化為 %d（整數型態），畫面上會是純淨的 0 與 1
    for kw in keyword_cols:
        custom_configs[kw] = st.column_config.NumberColumn(
            kw, 
            format="%d", 
            help=f"點擊排序檢視包含【{kw}】的標案"
        )

    # 渲染原始大寬表
    st.dataframe(
        filtered_df[available_display_cols],
        column_config=custom_configs,
        use_container_width=True,
        hide_index=True
    )
else:
    st.warning("⚠️ 找不到 `data/採購網_決標彙整.xlsx` 檔案，請確認爬蟲已成功執行並將檔案推上 GitHub。")