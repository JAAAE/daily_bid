import os
import time
import random
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# --- 💡 配置設定（修正 Z 槽死角，改為專案相對路徑，方便 Git 管理） ---
DATA_DIR = 'data'
excel_path = os.path.join(DATA_DIR, '採購網_決標彙整.xlsx')
KEYWORDS = ["測繪", "空間資訊", "測量", "製圖", "圖資", "地圖", "地形", "測製", "地理資訊", "監審", "光達", "點雲", "模型", "建模"] 

# 區域分類字典
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
        total=5,
        backoff_factor=2,
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
        # 💡 將超時從 1 秒放寬到 5 秒，避免政府 API 不穩導致頻繁下載失敗變空白
        response = session.get(url, timeout=5)
        if response.status_code == 200:
            return response.json()
    except Exception:
        pass
    return None

def fetch_data_for_date(date):
    url = f"https://pcc-api.openfun.app/api/listbydate?date={date}"
    try:
        response = session.get(url, timeout=5)
        if response.status_code == 200:
            return response.json()
    except Exception as e:
        print(f"日期 {date} 抓取失敗: {e}")
    return None

def process_data_for_date(date_str):
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

        # 💡 ✨ 關鍵縮排修正：只有詳細資料「成功抓到」才寫入，防止空資料造成去重時誤殺最新數據
        if content and 'records' in content and content['records']:
            detail = content['records'][0].get('detail', {})
            agency_code = detail.get('機關資料:機關代碼', '')
            agency_name = detail.get('機關資料:機關名稱', '')
            link2 = detail.get('url', '')
            place = detail.get('機關資料:機關地址', '')
            place_substring = place[4:7] if place else ""
            region_name = city_to_region.get(place_substring, '其他')

            # 💡 多重欄位防禦型抓取預算金額，避免金額全變 0
            raw_price = detail.get('採購資料:預算金額') or \
                        detail.get('採購資料:採購金額') or \
                        detail.get('採購資料:總預算金額') or \
                        detail.get('招標資料:預算金額') or \
                        detail.get('採購資料:預估金額') or ""
            
            if isinstance(raw_price, str):
                price = raw_price.replace(',', '').replace('元', '').replace('$', '').strip()
            else:
                price = raw_price

            # 保持你原始排列，直接把「區域」安排在「地點」後面
            base_data = [date_str, agency_code, agency_name, place_substring, region_name, tender_name, price, link2]
            processed_rows.append(base_data + found_flags + [row_sum])
    
    print(f"Done: {date_str} ({len(processed_rows)} rows)")
    return processed_rows


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    columns = ['日期', '機關代碼', '機關名稱', '地點', '區域', '標案名稱', '預算', '成果連結'] + KEYWORDS + ['關鍵字總計']
    
    # 預設歷史起點
    start_date_str = "20260515"
    df_old = pd.DataFrame(columns=columns)
    
    # 💡 ✨ 自動續爬機制：讀取專案現有 Excel，從最後一天的隔天開始自動接下去補
    if os.path.exists(excel_path):
        print("偵測到現有歷史 Excel，讀取進度中...")
        try:
            # 強制將日期讀取為字串，避免點零（.0）
            df_old = pd.read_excel(excel_path, sheet_name='全部彙整', dtype={'日期': str})
            df_old = df_old.reindex(columns=columns)
            if not df_old.empty:
                df_old['日期'] = df_old['日期'].astype(str).str.replace('.0', '', regex=False).str.strip()
                max_date_str = df_old['日期'].max()
                max_dt = datetime.strptime(max_date_str, '%Y%m%d')
                start_date_str = (max_dt + timedelta(days=1)).strftime('%Y%m%d')
                print(f"📈 歷史檔案最後真實更新到: {max_date_str}，將從 {start_date_str} 開始自動補齊！")
        except Exception as e:
            print(f"歷史資料解析失敗，將採用預設設定。錯誤原因: {e}")

    # 💡 強制設定為台灣時區 (UTC+8)，避免 GitHub 伺服器因時差少抓一天
    tw_tz = timezone(timedelta(hours=8))
    today_dt = datetime.now(tw_tz)
    
    end_date_str = today_dt.strftime('%Y%m%d')
    start_dt = datetime.strptime(start_date_str, '%Y%m%d')
    
    # 安全比較
    if start_dt > today_dt.replace(tzinfo=None):
        print("✨ 資料庫已是最新狀態，無需填補！")
        return

    print(f"🚀 開始準備填補中斷日期：自 {start_date_str} 至 {end_date_str}")
    date_list = []
    curr = start_dt
    while curr <= today_dt.replace(tzinfo=None):
        date_list.append(curr.strftime('%Y%m%d'))
        curr += timedelta(days=1)

    # 💡 保留你原本的多執行緒平行爬取架構，並將 workers 調至 2 加速中斷日期的補齊
    all_data = []
    with ThreadPoolExecutor(max_workers=2) as executor:
        results = list(executor.map(process_data_for_date, date_list))
        for res in results:
            all_data.extend(res)

    df_new = pd.DataFrame(all_data, columns=columns)
    
    # 預清洗新資料型態
    target_stats_cols = KEYWORDS + ['關鍵字總計']
    for col in target_stats_cols:
        if col in df_new.columns:
            df_new[col] = pd.to_numeric(df_new[col], errors='coerce').fillna(0).astype('Int64')

    # 確保合併前新舊資料的「日期」欄位型態百分之百絕對一致 (全文字)
    if not df_old.empty:
        df_old['日期'] = df_old['日期'].astype(str).str.replace('.0', '', regex=False).str.strip()
    if not df_new.empty:
        df_new['日期'] = df_new['日期'].astype(str).str.strip()

    # 合併舊資料與新填補資料
    df_total = pd.concat([df_old, df_new], ignore_index=True)
    
    # 精確去重
    df_total['標案名稱'] = df_total['標案名稱'].astype(str).str.strip()
    df_total['成果連結'] = df_total['成果連結'].astype(str).str.strip()
    df_total.drop_duplicates(subset=['標案名稱', '成果連結'], keep='first', inplace=True)

    # 💡 終極清洗：強制將所有 14 個關鍵字與總計欄位鎖定為純整數格式（切除 1.0 的小數點）
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

    print(f"🎉 成功！Excel 資料已完整填補並更新至最新日期：{excel_path}")

if __name__ == "__main__":
    main()