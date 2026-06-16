import os
import time
import random
from datetime import datetime, timedelta
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# --- 配置設定 ---
DATA_DIR = 'data'
excel_path = os.path.join(DATA_DIR, '採購網_決標彙整.xlsx')
KEYWORDS = ["測繪", "空間資訊", "測量", "製圖", "圖資", "地圖", "地形", "測製", "地理資訊", "監審", "光達", "點雲", "模型", "建模"]
REGIONS = {
    "北部": ['基隆市', '新北市', '臺北市', '台北市', '桃園市', '新竹縣', '新竹市'],
    "中部": ['苗栗縣', '臺中市', '台中市', '彰化縣', '雲林縣', '南投縣'],
    "南部": ['嘉義縣', '嘉義市', '臺南市', '台南市', '高雄市', '屏東縣'],
    "東部": ['宜蘭縣', '花蓮縣', '臺東縣', '台東縣'],
    "離島": ['澎湖縣', '金門縣', '連江縣']
}
CITY_TO_REGION = {city: region for region, cities in REGIONS.items() for city in cities}

def get_session():
    session = requests.Session()
    retry_strategy = Retry(
        total=5, backoff_factor=2,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"]
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session

session = get_session()

def fetch_url_content(url):
    try:
        time.sleep(random.uniform(0.1, 0.2))
        response = session.get(url, timeout=3)
        if response.status_code == 200:
            return response.json()
    except Exception:
        pass
    return None

def process_data_for_date(date_str):
    url = f"https://pcc-api.openfun.app/api/listbydate?date={date_str}"
    try:
        response = session.get(url, timeout=3)
        if response.status_code != 200:
            return []
        data = response.json()
    except Exception as e:
        print(f"日期 {date_str} 抓取失敗: {e}")
        return []

    if not data or 'records' not in data:
        return []

    processed_rows = []
    award_records = [r for r in data['records'] if r.get('brief', {}).get('type') == "決標公告"]

    for record in award_records:
        brief = record.get('brief', {})
        tender_name = brief.get('title', '')
        
        # --- 💡 關鍵字 0/1 統計邏輯 ---
        found_flags = [1 if k in tender_name else 0 for k in KEYWORDS]
        row_sum = sum(found_flags)
        if row_sum == 0:
            continue

        tender_url = record.get('tender_api_url', '')
        content = fetch_url_content(tender_url)
        agency_code, agency_name, price, link2, place_substring, region = "", "", "", "", "", "其他"

        if content and 'records' in content and content['records']:
            detail = content['records'][0].get('detail', {})
            agency_code = detail.get('機關資料:機關代碼', '')
            agency_name = detail.get('機關資料:機關名稱', '')
            link2 = detail.get('url', '')
            price = detail.get('採購資料:預算金額', '')
            place = detail.get('機關資料:機關地址', '')
            place_substring = place[4:7] if place else ""
            region = CITY_TO_REGION.get(place_substring, "其他")

        base_data = [date_str, agency_code, agency_name, place_substring, region, tender_name, price, link2]
        processed_rows.append(base_data + found_flags + [row_sum])
    
    print(f"Done: {date_str} ({len(processed_rows)} rows)")
    return processed_rows

def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    columns = ['日期', '機關代碼', '機關名稱', '地點', '區域', '標案名稱', '預算', '成果連結'] + KEYWORDS + ['關鍵字總計']
    
    start_date_str = "20260515"
    df_old = pd.DataFrame(columns=columns)
    
    # 1. 自動偵測歷史 Excel 進度
    if os.path.exists(excel_path):
        print("偵測到現有歷史 Excel，讀取進度中...")
        try:
            df_old = pd.read_excel(excel_path, sheet_name='全部彙整')
            df_old = df_old.reindex(columns=columns)
            if not df_old.empty:
                max_date_str = str(df_old['日期'].max()).strip()
                max_date = datetime.strptime(max_date_str, '%Y%m%d')
                start_date_str = (max_date + timedelta(days=1)).strftime('%Y%m%d')
                print(f"📈 歷史檔案最後更新到: {max_date_str}，將從隔日 {start_date_str} 開始自動追趕進度！")
        except Exception as e:
            print(f"歷史資料讀取失敗，將採用預設設定。錯誤: {e}")

    # 爬取終點：昨天
    yesterday_dt = datetime.now() - timedelta(days=1)
    end_date_str = yesterday_dt.strftime('%Y%m%d')
    start_dt = datetime.strptime(start_date_str, '%Y%m%d')
    
    if start_dt > yesterday_dt:
        print("✨ 資料庫已是最新狀態，無需填補！")
        return

    # 2. 迴圈補齊資料
    print(f"🚀 開始自動補齊中斷日期：自 {start_date_str} 至 {end_date_str}")
    all_new_rows = []
    curr = start_dt
    while curr <= yesterday_dt:
        date_str = curr.strftime('%Y%m%d')
        rows = process_data_for_date(date_str)
        all_new_rows.extend(rows)
        curr += timedelta(days=1)
        
    df_new = pd.DataFrame(all_new_rows, columns=columns)
    
    # --- 💡 強制確保新資料的關鍵字欄位是 int 型態，避免轉為 float ---
    for kw in KEYWORDS + ['關鍵字總計']:
        if kw in df_new.columns:
            df_new[kw] = df_new[kw].astype(int)

    # 3. 合併新舊資料
    df_total = pd.concat([df_old, df_new], ignore_index=True)

    # 4. 全域去重與清理
    df_total.drop_duplicates(subset=['標案名稱', '成果連結'], keep='first', inplace=True)
    df_total['日期'] = df_total['日期'].astype(str)
    
    # 再次確認合併後的統計欄位都是整數格式
    for kw in KEYWORDS + ['關鍵字總計']:
        df_total[kw] = pd.to_numeric(df_total[kw], errors='coerce').fillna(0).astype(int)
        
    df_total.sort_values(by='日期', ascending=False, inplace=True)

    # 5. 重新寫入 Excel（包含分區分頁）
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df_total.to_excel(writer, sheet_name='全部彙整', index=False)
        for region_name in REGIONS.keys():
            region_df = df_total[df_total['區域'] == region_name]
            if not region_df.empty:
                region_df.to_excel(writer, sheet_name=region_name, index=False)
                
    print(f"🎉 成功！Excel 檔案已完整填補，關鍵字獨立統計欄位無誤：{excel_path}")

if __name__ == "__main__":
    main()