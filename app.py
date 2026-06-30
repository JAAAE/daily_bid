import os  
import time  
import random
import io
from datetime import datetime, timedelta, timezone
import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="政府電子採購網標案(決標)", layout="wide")

# RWD 
st.html("""
    <style>
        /* 1. 調整頂部留白，保留足夠空間給標題 */
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 0rem !important;
            padding-left: 2rem !important;
            padding-right: 2rem !important;
        }
        
        /* 2. 移除會吃掉標題的隱藏區塊壓縮，改用單純減少邊距 */
        div[data-testid="stForm"] > div,
        div[data-testid="stMarkdownContainer"] {
            margin-bottom: -0.2rem !important;
        }

        /* 3. 精確壓扁指標卡片容器的內部留白 */
        div[data-testid="stContainer"] {
            padding: 0.4rem 1rem !important;
            margin-bottom: 0.2rem !important;
        }
        
        /* 4. 縮小下載按鈕元件的上下外邊距 */
        div[data-testid="stDownloadButton"] {
            margin-top: -0.2rem !important;
            margin-bottom: -0.2rem !important;
        }
    </style>
""")

KEYWORDS = ["測繪", "空間資訊", "測量", "製圖", "圖資", "地圖", "地形", "測製", "地理資訊", "監審", "光達", "點雲", "模型", "建模"]
REGIONS = {
    "北部": ['基隆市', '新北市', '臺北市', '台北市', '桃園市', '新竹縣', '新竹市'],
    "中部": ['苗栗縣', '臺中市', '台中市', '彰化縣', '雲林縣', '南投縣'],
    "南部": ['嘉義縣', '嘉義市', '臺南市', '台南市', '高雄市', '屏東縣'],
    "東部": ['宜蘭縣', '花蓮縣', '臺東縣', '台東縣'],
    "離島": ['澎湖縣', '金門縣', '連江縣']
}
CITY_TO_REGION = {city: region for region, cities in REGIONS.items() for city in cities}

@st.cache_data(ttl=3600)  
def get_integrated_data():
    columns = ['日期', '機關代碼', '機關名稱', '地點', '區域', '標案名稱', '預算', '成果連結'] + KEYWORDS + ['關鍵字總計']
    history_excel = "data/採購網_決標彙整.xlsx"
    
    if os.path.exists(history_excel):
        df_total = pd.read_excel(history_excel, sheet_name='全部彙整', dtype={'日期': str})
    else:
        st.error(f"找不到數據源檔案：{history_excel}，請確認背景爬蟲已成功產出檔案。")
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

def to_excel(df_to_download):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_download.to_excel(writer, index=False, sheet_name='篩選結果')
    return output.getvalue()

# --- 網頁主要佈局 ---
st.markdown("### 🌐 政府電子採購網標案(決標，從20230519至今，每天更新)")
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
    with st.container(border=True):
        m1, m2, m3 = st.columns(3)
        
        total_bids = f"{len(filtered_df):,} 件"
        total_budget = filtered_df['預算'].sum()
        budget_text = f"{total_budget / 100000000:,.2f} 億元" if total_budget >= 100000000 else f"{total_budget / 1000:,.0f} 萬元"
        raw_date = str(df['日期'].max())
        formatted_date = f"{raw_date[:4]}-{raw_date[4:6]}-{raw_date[6:]}" if len(raw_date) == 8 else raw_date

        with m1:
            st.html(f'<div style="font-size: 13px; color: #888888;">📊 目前篩選標案量</div><div style="font-size: 20px; font-weight: bold; line-height:1.2;">{total_bids}</div>')
        with m2:
            st.html(f'<div style="font-size: 13px; color: #888888;">💰 總決標預算規模</div><div style="font-size: 20px; font-weight: bold; line-height:1.2;">{budget_text}</div>')
        with m3:
            st.html(f'<div style="font-size: 13px; color: #888888;">📅 資料最後更新至</div><div style="font-size: 20px; font-weight: bold; line-height:1.2;">{formatted_date}</div>')

    # 下載按鈕與分頁資訊（併入同一橫列，節省垂直空間）
    btn_cols = st.columns([4, 6])
    with btn_cols[0]:
        if not filtered_df.empty:
            export_cols = ['日期', '機關名稱', '地點', '區域', '標案名稱', '成果連結', '預算'] + KEYWORDS + ['關鍵字總計']
            excel_data = to_excel(filtered_df[export_cols])
            st.download_button(
                label="📥 下載「全部篩選結果」 (Excel)",
                data=excel_data,
                file_name=f"採購網篩選結果_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ 無資料供下載。")

    # 分頁邏輯
    items_per_page = 20
    max_page = ((len(filtered_df) - 1) // items_per_page) + 1 if len(filtered_df) > 0 else 1
    if 'current_page' not in st.session_state: 
        st.session_state.current_page = 1
    if st.session_state.current_page > max_page:
        st.session_state.current_page = 1
    
    start_idx = (st.session_state.current_page - 1) * items_per_page
    page_df = filtered_df.iloc[start_idx:start_idx + items_per_page]

    # 欄位
    custom_configs = {
        "日期": st.column_config.TextColumn("決標日期", width="small"), 
        "機關名稱": st.column_config.TextColumn("機關名稱", width="medium"), 
        "地點": st.column_config.TextColumn("地點", width="small"), 
        "區域": st.column_config.TextColumn("區域", width="small"), 
        "標案名稱": st.column_config.TextColumn("標案名稱", width="large"), 
        "預算": st.column_config.NumberColumn("預算 (元)", format="$%,d", width="medium"), 
        "成果連結": st.column_config.LinkColumn("連結", display_text="檢視", width="small"),
        "關鍵字總計": st.column_config.NumberColumn("總計", format="%d", width="small")
    }
    for kw in KEYWORDS: 
        custom_configs[kw] = st.column_config.NumberColumn(kw, format="%d", width="small")

    # RWD 
    display_cols = ['日期', '機關名稱', '地點', '區域', '標案名稱', '成果連結', '預算'] + KEYWORDS + ['關鍵字總計']
    st.dataframe(
        page_df[display_cols], 
        column_config=custom_configs, 
        use_container_width=True, 
        hide_index=True,
        height=int(st.session_state.get('win_height', 400) if 'win_height' in st.session_state else 420) 
    )

    # 頁碼控制按鈕（與底部完美貼合）
    nav_col1, nav_col2, nav_col3 = st.columns([1.5, 1.5, 7])
    if nav_col1.button("⬅️ 上一頁", disabled=(st.session_state.current_page == 1), use_container_width=True):
        st.session_state.current_page -= 1
        st.rerun()
    if nav_col2.button("下一頁 ➡️", disabled=(st.session_state.current_page == max_page), use_container_width=True):
        st.session_state.current_page += 1
        st.rerun()
    with nav_col3:
        st.html(f"""
            <p style='text-align: right; color: #888888; font-size: 14px; margin: 0; padding-top: 6px;'>
                第 {st.session_state.current_page} / {max_page} 頁
            </p>
        """)