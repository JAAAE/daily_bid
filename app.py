import os  
import time  
import random
from datetime import datetime, timedelta, timezone
import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="政府電子採購網標案(決標)", layout="wide")

KEYWORDS = ["測繪", "空間資訊", "測量", "製圖", "圖資", "地圖", "地形", "測製", "地理資訊", "監審", "光達", "點雲", "模型", "建模"]
REGIONS = {
    "北部": ['基隆市', '新北市', '臺北市', '台北市', '桃園市', '新竹縣', '新竹市'],
    "中部": ['苗栗縣', '臺中市', '台中市', '彰化縣', '雲林縣', '南投縣'],
    "南部": ['嘉義縣', '嘉義市', '臺南市', '台南市', '高雄市', '屏東縣'],
    "東部": ['宜蘭縣', '花蓮縣', '臺東縣', '台東縣'],
    "離島": ['澎湖縣', '金門縣', '連江縣']
}
CITY_TO_REGION = {city: region for region, cities in REGIONS.items() for city in cities}

# 🌟 改為純粹讀取本地/GitHub 既有的 Excel 檔案，完全不執行即時爬蟲
@st.cache_data(ttl=3600)  # 一小時重新載入一次硬碟檔案即可
def get_integrated_data():
    columns = ['日期', '機關代碼', '機關名稱', '地點', '區域', '標案名稱', '預算', '成果連結'] + KEYWORDS + ['關鍵字總計']
    
    # 1. 唯讀：讀取背景腳本或 GitHub 動作生成的歷史 Excel 檔案
    history_excel = "data/採購網_決標彙整.xlsx"
    if os.path.exists(history_excel):
        df_total = pd.read_excel(history_excel, sheet_name='全部彙整', dtype={'日期': str})
    else:
        st.error(f"找不到數據源檔案：{history_excel}，請確認背景爬蟲已成功產出檔案。")
        return pd.DataFrame(columns=columns)
        
    if not df_total.empty:
        # 清洗與校正日期格式
        df_total['日期'] = df_total['日期'].astype(str).str.replace('.0', '', regex=False).str.strip()
        df_total.sort_values(by='日期', ascending=False, inplace=True)
            
    # 2. 強制對「預算」欄位進行數字轉型，防止資料異常
    if '預算' in df_total.columns:
        df_total['預算'] = df_total['預算'].astype(str) \
                                          .str.replace('$', '', regex=False) \
                                          .str.replace(',', '', regex=False) \
                                          .str.replace('元', '', regex=False) \
                                          .str.strip()
        df_total['預算'] = pd.to_numeric(df_total['預算'], errors='coerce').fillna(0)

    # 3. 強制轉整數清洗統計欄位
    for col in (KEYWORDS + ['關鍵字總計']):
        if col in df_total.columns:
            df_total[col] = pd.to_numeric(df_total[col], errors='coerce').fillna(0).astype(int)
            
    return df_total

# --- 網頁主要渲染佈局 ---
st.title("🌐 空間資訊與測繪標案 決標觀測站")
df = get_integrated_data()

if df is not None and not df.empty:
    st.sidebar.header("🔍 篩選條件")
    selected_region = st.sidebar.selectbox("選擇區域", ["全部"] + sorted(list(df['區域'].dropna().unique())))
    selected_keyword = st.sidebar.selectbox("主要關鍵字篩選", ["全部"] + KEYWORDS)

    filtered_df = df.copy()
    if selected_region != "全部": 
        filtered_df = filtered_df[filtered_df['區域'] == selected_region]
    if selected_keyword != "全部": 
        filtered_df = filtered_df[filtered_df[selected_keyword] == 1]

    # 指標
    m1, m2, m3 = st.columns(3)
    m1.metric("當前篩選標案量", f"{len(filtered_df)} 件")
    m2.metric("總決標預算規模", f"{filtered_df['預算'].sum() / 10000:,.0f} 萬元")
    m3.metric("資料最後更新至", str(df['日期'].max()))

    # 分頁
    items_per_page = 20
    max_page = ((len(filtered_df) - 1) // items_per_page) + 1 if len(filtered_df) > 0 else 1
    if 'current_page' not in st.session_state: 
        st.session_state.current_page = 1
    
    start_idx = (st.session_state.current_page - 1) * items_per_page
    page_df = filtered_df.iloc[start_idx:start_idx + items_per_page]

    # 欄位排布格式化
    custom_configs = {
        "日期": st.column_config.TextColumn("決標日期"), 
        "預算": st.column_config.NumberColumn("預算 (元)", format="$%,d"), 
        "成果連結": st.column_config.LinkColumn("連結", display_text="檢視")
    }
    for kw in KEYWORDS: 
        custom_configs[kw] = st.column_config.NumberColumn(kw, format="%d")

    st.dataframe(page_df[['日期', '機關名稱', '地點', '區域', '標案名稱', '成果連結', '預算'] + KEYWORDS + ['關鍵字總計']], column_config=custom_configs, use_container_width=True, hide_index=True)

    # 按鈕
    btn_col1, btn_col2, btn_col3 = st.columns([1, 1, 8])
    if btn_col1.button("⬅️ 上一頁", disabled=(st.session_state.current_page == 1)):
        st.session_state.current_page -= 1
        st.rerun()
    if btn_col2.button("下一頁 ➡️", disabled=(st.session_state.current_page == max_page)):
        st.session_state.current_page += 1
        st.rerun()
    btn_col3.write(f"第 {st.session_state.current_page} / {max_page} 頁")
