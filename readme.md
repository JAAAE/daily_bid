# Daily Bid (每日標案自動爬取與視覺化系統) 

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.28%2B-FF4B4B.svg)](https://streamlit.io/)
[![GitHub Actions](https://img.shields.io/badge/Automation-GitHub%20Actions-blueviolet.svg)](https://github.com/features/actions)

**Daily Bid** 是一個自動化政府採購網標案爬取、資料處理與動態儀表板分析系統。系統支援每日定時自動爬取政府公開招標資訊、清洗並儲存至 Excel 檔案，並透過 **Streamlit** 提供視覺化搜尋與統計分析介面，方便團隊隨時檢索與掌握最新標案動態。

---

## 系統架構與功能亮點 (Features)

- **自動化標案爬蟲 (`crawler.py`)**：自動爬取政府電子採購網最新公告，擷取標案名稱、招標單位、預算金額、發布日期與詳細連結。
- **Streamlit 視覺化儀表板 (`app.py`)**：
  - **動態關鍵字搜尋與篩選**：支援關鍵字比對、預算金額區間過濾。
  - **Plotly 圖表分析**：提供標案數量趨勢、單位招標排行及預算分佈等互動式圖表。
- **GitHub Actions 排程自動化 (`.github/workflows/daily_crawl.yml`)**：每日定時觸發爬蟲任務，自動更新標案資料庫。
- **數據儲存 (`data/`)**：將每日爬取的數據清洗後整理成結構化的 `.xlsx` 檔案，方便輸出。
- **DevContainer 支援 (`.devcontainer/`)**：即用的開發環境設定，支援 VS Code / GitHub Codespaces快速部署。

---

## 專案結構 (Project Structure)

```text
daily_bid/
├── .devcontainer/
│   └── devcontainer.json     # VS Code DevContainer 開發環境設定
├── .github/
│   └── workflows/
│       └── daily_crawl.yml   # GitHub Actions 每日自動爬取排程
├── data/
│   └── .xlsx                 # 標案數據儲存檔 (Excel 格式)
├── app.py                    # Streamlit 視覺化 Web 介面
├── crawler.py                # 標案爬蟲核心邏輯
├── requirements.txt          # Python 依賴套件套件包
└── README.md                 # 專案說明文件
