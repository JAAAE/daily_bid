import os
import time
import random
from datetime import datetime, timedelta, timezone # 💡 導入 timezone 解決時差
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# --- 配置設定（完全維持你原本的 data 資料夾與 14 個關鍵字） ---
DATA_DIR = 'data'
excel_path = os.path.join(DATA_DIR, '採購網_決標彙整.xlsx')
KEYWORDS = ["測繪", "空間資訊", "測量", "製圖", "圖資", "地圖", "地形", "測製", "地理資訊", "監審", "光達", "點雲", "模型", "建模"] 

regions = {
    "北部": ['基隆市', '新北市', '臺北市', '台北市', '桃園市', '新竹縣', '新竹市'],
    "中部": ['苗栗縣', '臺中市', '台中市', '彰化縣', '雲林縣', '南投縣'],
    "南部": ['嘉義縣', '嘉義市', '臺南市', '台南市', '高雄市', '屏東縣'],
    "東部": ['宜蘭縣', '花蓮縣', '臺東縣', '台東縣'],
    "離島": ['澎湖縣', '金門縣', '連江縣']
}
city_to_region = {city: region for region, cities in regions.items() for city in cities}

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
        # 💡 將詳細頁超時由 1 秒放寬到 8 秒，防止國外虛擬機跨海抓取時因網路延遲集體斷線變空白
        response = session.get(url, timeout=8)
        if response.status_code == 200:
            return response.json()
    except Exception:
        pass
    return None

def fetch_data_for_date(date):
    url = f"https://pcc-api.openfun.app/api/listbydate?date={date}"
    try:
        response = session.get(url, timeout=8) # 💡 同步放寬到 8 秒
        if response.status_code == 200:
            return response.json()
    except Exception as e:
        print(f"❌ 日期 {date} 總表抓取失敗: {e}")
    return None

def process_data_for_date(date_str):
    """ 完全維持你原本的單日處理邏輯與精確縮排 """
    data = fetch_data_for_date(date_str)
    if not data or 'records' not in data:
        return []

    processed_rows = []
    award_records = [r for r in data['records'] if r.get('brief', {}).get('type') == "決標公告"]

    for record in award_records:
        brief = record.get('brief', {})
        tender_name = brief.get('title', '')
        
        found_flags = [1 if k in tender_name else 0 for k in KEYWORDS]
        row_sum = sum(found_flags)
        
        if row_sum == 0:
            continue

        tender_url = record.get('tender_api_url', '')
        content = fetch_url_content(tender_url)

        # 💡 確保明細抓取成功才寫入，防範去重時誤殺
        if content and 'records' in content and content['records']:
            detail = content['records'][0].get('detail', {})
            agency_code = detail.get('機關資料:機關代碼', '')
            agency_name = detail.get('機關資料:機關名稱', '')
            link2 = detail.get('url', '')
            place = detail.get('機關資料:機關地址', '')
            place_substring = place[4:7] if place else ""
            region_name = city_to_region.get(place_substring, '其他')

            raw_price = detail.get('採購資料:預算金額') or \
                        detail.get('採購資料:採購金額') or \
                        detail.get('採購資料:總預算金額') or \
                        detail.get('招標資料:預算金額') or \
                        detail.get('採購資料:預估金額') or ""
            
            if isinstance(raw_price, str):
                price = raw_price.replace(',', '').replace('元', '').replace('$', '').strip()
            else:
                price = raw_price

            base_data = [date_str, agency_code, agency_name, place_substring, region_name, tender_name, price, link2]
            processed_rows.append(base_data + found_flags + [row_sum])
    
    print(f"📅 日期 {date_str} 掃描完畢，命中空間資訊標案: {len(processed_rows)} 筆")
    return processed_rows

def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    columns = ['日期', '機關代碼', '機關名稱', '地點', '區域', '標案名稱', '預算', '成果連結'] + KEYWORDS + ['關鍵字總計']
    
    # 💡 鎖定台灣時區 (UTC+8) 計算今天，消滅雲端時差
    tw_tz = timezone(timedelta(hours=8))
    today_dt = datetime.now(tw_tz)
    
    # 💡 預設安全起爬點：如果讀不到歷史 Excel，預設至少往前推 5 天續爬，絕對不會秒退
    start_date_str = (today_dt - timedelta(days=5)).strftime('%Y%m%d')
    df_old = pd.DataFrame(columns=columns)
    
    # 自動偵測歷史 Excel 進度
    if os.path.exists(excel_path):
        try:
            df_old = pd.read_excel(excel_path, sheet_name='全部彙整', dtype={'日期': str})
            df_old = df_old.reindex(columns=columns)
            if not df_old.empty:
                df_old['日期'] = df_old['日期'].astype(str).str.replace('.0', '', regex=False).str.strip()
                max_date_str = df_old['日期'].max()
                max_dt = datetime.strptime(max_date_str, '%Y%m%d')
                start_date_str = (max_dt + timedelta(days=1)).strftime('%Y%m%d')
                print(f"📈 歷史檔案最後日期為: {max_date_str}，下一階段起爬點: {start_date_str}")
        except Exception as e:
            print(f"⚠️ 歷史進度解析失敗，採用安全預設起爬點。原因: {e}")

    end_date_str = today_dt.strftime('%Y%m%d')
    start_dt = datetime.strptime(start_date_str, '%Y%m%d')
    
    # 如果計算出來的起爬點大於今天，才不需要填補
    if start_dt > today_dt.replace(tzinfo=None):
        print("✨ 資料庫已是最新狀態，無需填補！")
        return

    # 建立日期清單
    date_list = []
    curr = start_dt
    while curr <= today_dt.replace(tzinfo=None):
        date_list.append(curr.strftime('%Y%m%d'))
        curr += timedelta(days=1)

    print(f"🚀 準備爬取的日期範圍: {date_list}")

    # 維持你原本最習慣、在本機跑得最順的單執行緒安全迴圈架構
    all_data = []
    for d_str in date_list:
        res = process_data_for_date(d_str)
        all_data.extend(res)

    df_new = pd.DataFrame(all_data, columns=columns)
    
    target_stats_cols = KEYWORDS + ['關鍵字總計']
    for col in target_stats_cols:
        if col in df_new.columns:
            df_new[col] = pd.to_numeric(df_new[col], errors='coerce').fillna(0).astype('Int64')

    if not df_old.empty:
        df_old['日期'] = df_old['日期'].astype(str).str.replace('.0', '', regex=False).str.strip()
    if not df_new.empty:
        df_new['日期'] = df_new['日期'].astype(str).str.strip()

    # 合併歷史與全新數據
    df_total = pd.concat([df_old, df_new], ignore_index=True)
    
    df_total['日期'] = df_total['日期'].astype(str).str.strip()
    df_total['標案名稱'] = df_total['標案名稱'].astype(str).str.strip()
    df_total['成果連結'] = df_total['成果連結'].astype(str).str.strip()
    
    # 雙重安全去重基準
    df_total.drop_duplicates(subset=['日期', '標案名稱', '成果連結'], keep='first', inplace=True)

    for col in target_stats_cols:
        if col in df_total.columns:
            df_total[col] = pd.to_numeric(df_total[col], errors='coerce').fillna(0).astype(int)
            
    df_total.sort_values(by='日期', ascending=False, inplace=True)

    # 重新寫入 Excel 各分頁
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df_total.to_excel(writer, sheet_name='全部彙整', index=False)
        for region_name, cities in regions.items():
            region_df = df_total[df_total['區域'] == region_name].copy()
            if not region_df.empty:
                for col in target_stats_cols:
                    if col in region_df.columns:
                        region_df[col] = region_df[col].astype(int)
                region_df.to_excel(writer, sheet_name=region_name, index=False)

    print(f"🎉 成功！Excel 資料已更新完畢，統計欄位皆維持純整數格式：{excel_path}")

if __name__ == "__main__":
    main()