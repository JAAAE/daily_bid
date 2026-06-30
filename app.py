import os  
import time  
import random
import io
from datetime import datetime, timedelta, timezone
import pandas as pd
import requests
import streamlit as st

# 1. 頁面初始化與基本設定
st.set_page_config(
    page_title="政府電子採購網標案(決標)", 
    layout="wide",
    initial_sidebar_state="expanded"
)

KEYWORDS = ["測繪", "空間資訊", "測量", "製圖", "圖資", "地圖", "地形", "測製", "地理資訊", "監審", "光達", "點雲", "模型", "建模"]
REGIONS = {
    "北部": ['基隆市', '新北市', '臺北市', '台北市', '桃園市', '新竹縣', '新竹市'],
    "中部": ['苗栗縣', '臺中市', '台中市', '彰化縣', '雲林縣', '南投縣'],
    "南部": ['嘉義縣', '嘉義市', '臺南市', '台南市', '高雄市', '屏東縣'],
    "東部": ['宜蘭縣', '花蓮縣', '臺東縣', '台東縣'],
    "離島": ['澎湖縣', '金門縣', '連江縣']
}
CITY_TO_REGION = {city: region for region, cities in REGIONS.items() for city in cities}

# 2. 資料讀取與清洗 (快取 1 小時)
@st.cache_data(ttl=3600)
def get_integrated_data():
    columns = ['日期', '機關代碼', '機關名稱', '地點', '區域', '標案名稱', '預算', '成果連結'] + KEYWORDS + ['關鍵字總計']
    history_excel = "data/採購網_決標彙整.xlsx"
    
    if os.path.exists(history_excel):
        df_total = pd.read_excel(history_excel, sheet_name='全部彙整', dtype={'日期': str})
    else:
        st.error(f"⚠️ 找不到數據源檔案：{history_excel}，請確認背景爬蟲已成功產出檔案。")
        return pd.DataFrame(columns=columns)
        
    if not df_total.empty:
        df_total['日期'] = df_total['日期'].astype(str).str.replace('.0', '', regex=False).str.strip()
        df_total.sort_values(by='日期', ascending=False, inplace=True)
            
    if '預算' in df_total.columns:
        df_total['預算'] = df_total['預算'].astype(str) \
                                          .str.replace('$', '', regex=False) \
                                          .str.replace(',', '', regex=False) \
                                          .str.replace('元', '', regex=False) \
                                          .str.strip()
        df_total['預算'] = pd.to_numeric(df_total['預算'], errors='coerce').fillna(0)

    for col in (KEYWORDS + ['關鍵字總計']):
        if col in df_total.columns:
            df_total[col] = pd.to_numeric(df_total[col], errors='coerce').fillna(0).astype(int)
            
    return df_total

# 3. Excel 匯出輔助函式
def to_excel(df_to_download):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_download.to_excel(writer, index=False, sheet_name='篩選結果')
    return output.getvalue()

# --- 網頁主視覺渲染 ---
st.title("🌐 政府電子採購網標案系統")
st.caption("標案資料範圍：2023-05-19 至今 (系統每日自動更新)")

df = get_integrated_data()

if df is not None and not df.empty:
    # 側邊欄篩選面板
    st.sidebar.header("🔍 篩選條件")
    selected_region = st.sidebar.selectbox("選擇區域", ["全部"] + sorted(list(df['區域'].dropna().unique())))
    selected_keyword = st.sidebar.selectbox("主要關鍵字篩選", ["全部"] + KEYWORDS)

    # 執行篩選過濾
    filtered_df = df.copy()
    if selected_region != "全部": 
        filtered_df = filtered_df[filtered_df['區域'] == selected_region]
    if selected_keyword != "全部": 
        filtered_df = filtered_df[filtered_df[selected_keyword] == 1]

    # 關鍵指標數據區
    st.markdown("### 📈 當前篩選統計")
    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric("篩選標案總數", f"{len(filtered_df)} 件")
    with m2:
        st.metric("總決標預算規模", f"{filtered_df['預算'].sum() / 10000:,.0f} 萬元")
    with m3:
        st.metric("資料最後更新日", str(df['日期'].max()))

    st.markdown("---")

    # 操作與下載區塊
    act_col1, act_col2 = st.columns([7, 3])
    with act_col1:
        st.markdown("### 📋 標案詳細數據")
    with act_col2:
        if not filtered_df.empty:
            export_cols = ['日期', '機關名稱', '地點', '區域', '標案名稱', '成果連結', '預算'] + KEYWORDS + ['關鍵字總計']
            excel_data = to_excel(filtered_df[export_cols])
            
            # 使用 container 右對齊下載按鈕的視覺感
            st.download_button(
                label="下載「全部篩選結果」 (Excel)",
                data=excel_data,
                file_name=f"採購網篩選結果_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.warning("⚠️ 目前篩選條件下無資料可供下載。")

    # 分頁邏輯計算
    items_per_page = 20
    max_page = ((len(filtered_df) - 1) // items_per_page) + 1 if len(filtered_df) > 0 else 1
    
    if 'current_page' not in st.session_state: 
        st.session_state.current_page = 1
    
    if st.session_state.current_page > max_page:
        st.session_state.current_page = 1
    
    start_idx = (st.session_state.current_page - 1) * items_per_page
    page_df = filtered_df.iloc[start_idx:start_idx + items_per_page]

    # 資料表格美化與格式化配置
    custom_configs = {
        "日期": st.column_config.TextColumn("決標日期", width="medium"), 
        "預算": st.column_config.NumberColumn("預算 (元)", format="$%,d", width="medium"), 
        "成果連結": st.column_config.LinkColumn("標案連結", display_text="🌐 點我檢視", width="small")
    }
    for kw in KEYWORDS: 
        custom_configs[kw] = st.column_config.NumberColumn(kw, format="%d", width="small")

    # 顯示當頁數據
    display_cols = ['日期', '機關名稱', '地點', '區域', '標案名稱', '成果連結', '預算'] + KEYWORDS + ['關鍵字總計']
    st.dataframe(
        page_df[display_cols], 
        column_config=custom_configs, 
        use_container_width=True, 
        hide_index=True
    )

    # 分頁按鈕與頁碼區塊
    st.markdown(" ") # 留白
    btn_col1, btn_col2, btn_col3 = st.columns([1.5, 1.5, 7])
    
    with btn_col1:
        st.button(
            "⬅️ 上一頁", 
            disabled=(st.session_state.current_page == 1), 
            use_container_width=True,
            on_click=lambda: st.session_state.update(current_page=st.session_state.current_page - 1)
        )
    with btn_col2:
        st.button(
            "下一頁 ➡️", 
            disabled=(st.session_state.current_page == max_page), 
            use_container_width=True,
            on_click=lambda: st.session_state.update(current_page=st.session_state.current_page + 1)
        )
    with btn_col3:
        st.markdown(f"##### <p style='text-align: right; color: gray; margin-top: 5px;'>第 {st.session_state.current_page} / {max_page} 頁</p>", unsafe_html=True)