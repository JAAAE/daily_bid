import os  
import time  
import random
from datetime import datetime, timedelta, timezone
import pandas as pd
import requests
import streamlit as st
import io
import base64

st.set_page_config(page_title="政府電子採購網決標觀測站", layout="wide")

KEYWORDS = ["測繪", "空間資訊", "測量", "製圖", "圖資", "地圖", "地形", "測製", "地理資訊", "監審", "光達", "點雲", "模型", "建模"]
REGIONS = {
    "北部": ['基隆市', '新北市', '臺北市', '台北市', '桃園市', '新竹縣', '新竹市'],
    "中部": ['苗栗縣', '臺中市', '台中市', '彰化縣', '雲林縣', '南投縣'],
    "南部": ['嘉義縣', '嘉義市', '臺南市', '台南市', '高雄市', '屏東縣'],
    "東部": ['宜蘭縣', '花蓮縣', '臺東縣', '台東縣'],
    "離島": ['澎湖縣', '金門縣', '連江縣']
}
CITY_TO_REGION = {city: region for region, cities in REGIONS.items() for city in cities}

# --- GitHub 倉庫配置 ---
REPO_OWNER = "JAAAE"
REPO_NAME = "daily_bid"
FILE_PATH = "data/採購網_決標彙整.xlsx"

def fetch_url_content(url):
    try:
        time.sleep(random.uniform(0.05, 0.1))
        response = requests.get(url, timeout=3) # 拉短超時防止掛起
        if response.status_code == 200: 
            return response.json()
    except: 
        pass
    return None

def fetch_data_for_date(date):
    url = f"https://pcc-api.openfun.app/api/listbydate?date={date}"
    try:
        response = requests.get(url, timeout=4)
        if response.status_code == 200: 
            return response.json()
    except: 
        pass
    return None

def crawl_live_data(date_list):
    """ 在 Streamlit 背景執行即時爬取 """
    all_rows = []
    for d_str in date_list:
        data = fetch_data_for_date(d_str)
        if not data or 'records' not in data or not data['records']: 
            continue
        
        award_records = [r for r in data['records'] if r.get('brief', {}).get('type') == "決標公告"]
        for record in award_records:
            tender_name = record.get('brief', {}).get('title', '')
            found_flags = [1 if k in tender_name else 0 for k in KEYWORDS]
            if sum(found_flags) == 0: 
                continue
            
            tender_url = record.get('tender_api_url', '')
            if not tender_url:
                continue
                
            content = fetch_url_content(tender_url)
            
            # 安全防禦鎖：確保 records 有長度才能讀取 [0]
            if content and 'records' in content and isinstance(content['records'], list) and len(content['records']) > 0:
                block = content['records'][0]
                if not block or 'detail' not in block:
                    continue
                
                detail = block.get('detail', {})
                place = detail.get('機關資料:機關地址', '')
                place_substring = place[4:7] if place and len(place) >= 7 else ""
                
                raw_price = detail.get('採購資料:預算金額') or detail.get('採購資料:採購金額') or ""
                price = raw_price.replace(',', '').replace('元', '').strip() if isinstance(raw_price, str) else raw_price
                
                base = [d_str, detail.get('機關資料:機關代碼', ''), detail.get('機關資料:機關名稱', ''), 
                        place_substring, CITY_TO_REGION.get(place_substring, "其他"), tender_name, price, detail.get('url', '')]
                all_rows.append(base + found_flags + [sum(found_flags)])
    return all_rows

def push_excel_to_github(df_total, target_stats_cols):
    """將更新後的多分頁 Excel 強制推回倉庫，解開萬年死鎖"""
    if "GITHUB_TOKEN" not in st.secrets:
        return

    token = st.secrets["GITHUB_TOKEN"]
    url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/{FILE_PATH}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}

    try:
        res = requests.get(url, headers=headers)
        sha = res.json().get("sha", "") if res.status_code == 200 else ""

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_total.to_excel(writer, sheet_name='全部彙整', index=False)
            for region_name, cities in REGIONS.items():
                region_df = df_total[df_total['區域'] == region_name].copy()
                if not region_df.empty:
                    for col in target_stats_cols:
                        if col in region_df.columns:
                            region_df[col] = region_df[col].astype(int)
                    region_df.to_excel(writer, sheet_name=region_name, index=False)
        
        excel_binary = output.getvalue()
        base64_content = base64.b64encode(excel_binary).decode("utf-8")

        payload = {
            "message": f"🤖 Streamlit 自動越過死鎖同步 Excel: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "content": base64_content,
            "branch": "main"
        }
        if sha: 
            payload["sha"] = sha

        requests.put(url, headers=headers, json=payload)
    except Exception as e:
        print(f"⚠️ GitHub 同步失敗: {e}")

@st.cache_data(ttl=2) # 💡 設定短快取，點開網頁立即出發重新校正
def get_integrated_data():
    columns = ['日期', '機關代碼', '機關名稱', '地點', '區域', '標案名稱', '預算', '成果連結'] + KEYWORDS + ['關鍵字總計']
    target_stats_cols = KEYWORDS + ['關鍵字總計']
    
    # 建立本地目錄防禦
    os.makedirs("data", exist_ok=True)
    
    history_excel = "data/採購網_決標彙整.xlsx"
    if os.path.exists(history_excel):
        try:
            df_total = pd.read_excel(history_excel, sheet_name='全部彙整', dtype={'日期': str})
        except:
            df_total = pd.DataFrame(columns=columns)
    else:
        df_total = pd.DataFrame(columns=columns)
        
    tw_tz = timezone(timedelta(hours=8))
    today_dt = datetime.now(tw_tz)
    
    start_date_str = "20260515"
    if not df_total.empty:
        df_total['日期'] = df_total['日期'].astype(str).str.replace('.0', '', regex=False).str.strip()
        start_date_str = (datetime.strptime(df_total['日期'].max(), '%Y%m%d') + timedelta(days=1)).strftime('%Y%m%d')
        
    start_dt = datetime.strptime(start_date_str, '%Y%m%d')
    
    # 💡 ✨ 終極暴力除錯核心：只要落後今天，就強制拉出一整條直到今天的日期清單，絕不被單一日期卡死
    if start_dt <= today_dt.replace(tzinfo=None):
        date_list = []
        curr = start_dt
        while curr <= today_dt.replace(tzinfo=None):
            date_list.append(curr.strftime('%Y%m%d'))
            curr += timedelta(days=1)
            
        print(f"Streamlit 背景補齊計畫啟動，日期清單: {date_list}")
        new_rows = crawl_live_data(date_list)
        
        if new_rows:
            df_new = pd.DataFrame(new_rows, columns=columns)
            df_total = pd.concat([df_total, df_new], ignore_index=True)
            df_total.drop_duplicates(subset=['標案名稱', '成果連結'], keep='first', inplace=True)
        
        # 💡 ✨ 關鍵修正：只要觸發了日期填補，不論當天有沒有撈到新案子，我們都強制同步回推 GitHub。
        # 這會重寫 Excel，強迫更新修改時間與 Sha 憑證，歷史指針將會瞬間跨越 5/21 阻礙！
        df_total.sort_values(by='日期', ascending=False, inplace=True)
        
        # 資料型態清洗 (在同步前先做好)
        if '預算' in df_total.columns:
            df_total['預算'] = df_total['預算'].astype(str) \
                                              .str.replace('$', '', regex=False) \
                                              .str.replace(',', '', regex=False) \
                                              .str.replace('元', '', regex=False) \
                                              .str.strip()
            df_total['預算'] = pd.to_numeric(df_total['預算'], errors='coerce').fillna(0)

        for col in target_stats_cols:
            if col in df_total.columns:
                df_total[col] = pd.to_numeric(df_total[col], errors='coerce').fillna(0).astype(int)
        
        # 強制同步至 GitHub 倉庫
        push_excel_to_github(df_total, target_stats_cols)
    else:
        # 如果沒落後進度，仍要確保型態正確
        if '預算' in df_total.columns:
            df_total['預算'] = pd.to_numeric(df_total['預算'], errors='coerce').fillna(0)
        for col in target_stats_cols:
            if col in df_total.columns:
                df_total[col] = pd.to_numeric(df_total[col], errors='coerce').fillna(0).astype(int)
        df_total.sort_values(by='日期', ascending=False, inplace=True)
            
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
    st.columns(3)[0].metric("當前篩選標案量", f"{len(filtered_df)} 件")
    st.columns(3)[1].metric("總決標預算規模", f"{filtered_df['預算'].sum() / 10000:,.0f} 萬元")
    st.columns(3)[2].metric("最新觀測日期", str(df['日期'].max()))

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