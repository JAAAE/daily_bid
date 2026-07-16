import os
import sys
import time
import codecs
import random
import shutil
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

sys.stdout = codecs.getwriter("utf-8")(sys.stdout.detach())
sys.stderr = codecs.getwriter("utf-8")(sys.stderr.detach())

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
excel_path = os.path.join(DATA_DIR, '採購網_決標彙整.xlsx')

BACKUP_PATH = os.path.join(BASE_DIR, "latest_crawl_output.xlsx")

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
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
        "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1"
    }
    session.headers.update(headers)
    
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
        time.sleep(random.uniform(0.5, 1.5))  
        
        if url.startswith("http://"):
            url = url.replace("http://", "https://")
            
        response = session.get(url, timeout=15)
        
        if response.status_code == 200:
            try:
                return response.json()
            except requests.exceptions.JSONDecodeError:
                print(f"❌ [格式錯誤] 網址回傳的不是 JSON。前 100 字元: {response.text[:100]}")
                return None
        else:
            print(f"⚠️ [狀態碼 {response.status_code}] 請求被拒或無效: {url}")
            return None
            
    except requests.exceptions.Timeout:
        print(f"⏳ [連線逾時] 伺服器超過 15 秒未回應: {url}")
    except Exception as e:
        print(f"🚨 [未知錯誤] 連線失敗: {e} | URL: {url}")
        
    return None

def fetch_data_for_date(date):
    url = f"https://pcc-api.openfun.app/api/listbydate?date={date}"
    try:
        response = session.get(url, timeout=10)
        if response.status_code == 200:
            return response.json()
    except Exception as e:
        print(f"日期 {date} 抓取失敗: {e}")
    return None

def process_data_for_date(date_str):
    data = fetch_data_for_date(date_str)
    if not data or 'records' not in data or not data['records']:
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

        agency_code, agency_name, price, link2, place_substring, region = "", "", "", "", "", "其他"

        if content and 'records' in content and content['records']:
            target_block = None
            
            for b in content['records']:
                block_type = b.get('detail', {}).get('type', '')
                if block_type == "決標公告":
                    target_block = b
                    break
            
            if not target_block:
                continue
                
            detail = target_block.get('detail', {})
            agency_code = target_block.get('detail', {}).get('機關資料:機關代碼', '')
            agency_name = target_block.get('detail', {}).get('機關資料:機關名稱', '')
            link2 = target_block.get('detail', {}).get('url', '')
            place = target_block.get('detail', {}).get('機關資料:機關地址', '')
            
            if place:
                place_clean = ''.join([ch for ch in str(place) if not ch.isdigit()]).strip()
                place_substring = place_clean[:3]
                region = CITY_TO_REGION.get(place_substring, "其他")

  
            
            
            raw_price = ""
            
            price_keywords = ["決標金額", "採購金額", "預算金額", "預估金額", "金額"]
            
            best_match_idx = len(price_keywords)

            for key, value in detail.items():
                if value:  
                    for idx, kw in enumerate(price_keywords):
                        if kw in key and idx < best_match_idx:
                            raw_price = str(value)
                            best_match_idx = idx
                            break 

            if raw_price:
                price = raw_price.replace(',', '').replace('元', '').replace('$', '').strip()
            else:
                price = "0"  
            
            if isinstance(raw_price, str):
                price = raw_price.replace(',', '').replace('元', '').replace('$', '').strip()
            else:
                price = raw_price

     
        base_data = [date_str, agency_code, agency_name, place_substring, region, tender_name, price, link2]
        processed_rows.append(base_data + found_flags + [row_sum])
    
    print(f"Done: {date_str} ({len(processed_rows)} rows)")
    return processed_rows

def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    columns = ['日期', '機關代碼', '機關名稱', '地點', '區域', '標案名稱', '預算', '成果連結'] + KEYWORDS + ['關鍵字總計']
    target_stats_cols = KEYWORDS + ['關鍵字總計']
    
    start_date_str = "20260515"
    df_old = pd.DataFrame(columns=columns)
    
    if os.path.exists(excel_path):
        try:
            df_old = pd.read_excel(excel_path, sheet_name='全部彙整', dtype={'日期': str})
            df_old = df_old.reindex(columns=columns)
            
            if not df_old.empty:
                df_old = df_old.dropna(subset=['日期'])
                df_old['日期'] = df_old['日期'].astype(str).str.replace('.0', '', regex=False).str.strip()
                
                if not df_old.empty:
                    max_date_str = df_old['日期'].max()
                    max_dt = datetime.strptime(max_date_str, '%Y%m%d')
                    start_date_str = (max_dt + timedelta(days=1)).strftime('%Y%m%d')
                    print(f"歷史檔案最後真實更新到: {max_date_str}，下一階段將從 {start_date_str} 開始自動補齊！")
        except Exception as e:
            print(f"歷史資料解析失敗: {e}")

    tw_tz = timezone(timedelta(hours=8))
    today_dt = datetime.now(tw_tz)
    
    end_date_str = today_dt.strftime('%Y%m%d')
    start_dt = datetime.strptime(start_date_str, '%Y%m%d')
    
    if start_dt > today_dt.replace(tzinfo=None):
        print("資料庫已是最新狀態，直接結束！")
        if os.path.exists(excel_path):
            shutil.copy(excel_path, BACKUP_PATH)
        return

    print(f"開始準備填補中斷日期：自 {start_date_str} 至 {end_date_str}")
    date_list = []
    curr = start_dt
    while curr <= today_dt.replace(tzinfo=None):
        date_list.append(curr.strftime('%Y%m%d'))
        curr += timedelta(days=1)

    all_new_rows = []
    with ThreadPoolExecutor(max_workers=2) as executor:
        results = list(executor.map(process_data_for_date, date_list))
        for res in results:
            all_new_rows.extend(res)
        
    df_new = pd.DataFrame(all_new_rows, columns=columns)
    
    if not df_old.empty:
        df_old['日期'] = df_old['日期'].astype(str).str.replace('.0', '', regex=False).str.strip()
    if not df_new.empty:
        df_new['日期'] = df_new['日期'].astype(str).str.strip()

    df_total = pd.concat([df_old, df_new], ignore_index=True)
    
    df_total['標案名稱'] = df_total['標案名稱'].astype(str).str.strip()
    df_total['成果連結'] = df_total['成果連結'].astype(str).str.strip()
    df_total.drop_duplicates(subset=['標案名稱', '成果連結'], keep='first', inplace=True)
    
    if '預算' in df_total.columns:
        df_total['預算'] = df_total['預算'].astype(str).str.replace('$', '', regex=False).str.replace(',', '', regex=False).str.replace('元', '', regex=False).str.strip()
        df_total['預算'] = pd.to_numeric(df_total['預算'], errors='coerce').fillna(0)

    for col in target_stats_cols:
        if col in df_total.columns:
            df_total[col] = pd.to_numeric(df_total[col], errors='coerce').fillna(0).astype(int)
            
    df_total.sort_values(by='日期', ascending=False, inplace=True)

    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df_total.to_excel(writer, sheet_name='全部彙整', index=False)
        for region_name in REGIONS.keys():
            region_df = df_total[df_total['區域'] == region_name].copy()
            if not region_df.empty:
                for col in target_stats_cols:
                    if col in region_df.columns:
                        region_df[col] = region_df[col].astype(int)
                region_df.to_excel(writer, sheet_name=region_name, index=False)
                
    print(f"成功！Excel 資料已更新完畢（已實現安全 Append 追加）：{excel_path}")

    try:
        shutil.copy(excel_path, BACKUP_PATH)
        print(f"[Python 內部正名] 英文備份檔已就緒: {BACKUP_PATH}")
    except Exception as e:
        print(f"備份英文檔失敗: {e}")

if __name__ == "__main__":
    main()
