# FH-NEWS — 金管會月報 Excel 入庫工具

把金管會（banking.gov.tw）每月「單獨公布的 Excel」自動抓下來，
彙整進一個 **SQLite** 資料庫，方便後續查詢與分析。

> ⚠️ **網路限制**：Claude Code on the web 的雲端環境預設網路政策**擋住
> `banking.gov.tw`**，所以爬取（`run` / `crawl`）需要在**你自己的本機**執行。
> 若要在 web 環境直接跑，需先調整環境的網路政策放行該網域
> （見 https://code.claude.com/docs/en/claude-code-on-the-web ）。
> 純解析入庫（`load`）不需網路，任何環境都能跑。

## 安裝

```bash
pip install -r requirements.txt
cp config.example.yaml config.yaml   # 再依需求編輯
```

## 設定（config.yaml）

| 欄位 | 說明 |
|------|------|
| `start_urls` | 起始頁。可放單篇公告頁，或「列表頁」。 |
| `follow_article_links` | 是否自動跟著頁面上「其他月份公告」連結往下爬一層（抓整個系列時開啟）。 |
| `raw_dir` | 下載的 Excel 存放目錄（預設 `data/raw`，已 gitignore）。 |
| `db_path` | SQLite 檔名（預設 `fsc_news.sqlite`）。 |
| `request_delay` | 每次請求間隔秒數，對政府網站客氣一點。 |

## 使用

```bash
# 1) 先檢查能不能抓到 Excel 連結（不下載，純列出）
python -m fsc crawl --config config.yaml

# 2) 完整流程：爬取 → 下載 → 入庫
python -m fsc run --config config.yaml

# 3) 若已手動把 Excel 放進 data/raw，只做入庫（免連網）
python -m fsc load --config config.yaml
```

## 資料庫結構

採「無損長表（long format）」設計——因為金管會各月 Excel 的欄位、標題、
合併儲存格常隨年度變動，固定 schema 容易壞。先把資料**完整、不失真**地收進來，
正規化留到 SQL 端做。

- **`source_files`** — 檔案登記表
  `id, url, filename, sha256(去重用), period('YYYY-MM'), size_bytes, downloaded_at`
- **`cells`** — 每個工作表的每一格
  `file_id, sheet, row, col, header(欄標題), value`

`period` 會自動從檔名／標題推算，支援民國（`10805` / `108年5月`）與西元（`201905`）。

### 範例查詢

```sql
-- 看有哪些月份、各收了多少格資料
SELECT f.period, COUNT(*) AS cells
FROM cells c JOIN source_files f ON f.id = c.file_id
GROUP BY f.period ORDER BY f.period;

-- 把某個欄位（例如「存款總額」）跨月份拉出來
SELECT f.period, c.value
FROM cells c JOIN source_files f ON f.id = c.file_id
WHERE c.header LIKE '%存款%' ORDER BY f.period;
```

## 畫圖與修改

SQLite 只負責「存資料」，要不要修改、怎麼畫圖都很自由：

- **修改**：`UPDATE/INSERT/DELETE/ALTER`，或用 [DB Browser for SQLite](https://sqlitebrowser.org/)
  圖形介面像 Excel 一樣編輯。重跑 `run` 靠 sha256 去重，不會覆蓋你手改過的舊資料。
- **畫圖**：Python（pandas + matplotlib/plotly）、Excel 匯入、或 BI 工具
  （Power BI / Tableau / Metabase / Grafana）皆可直接連 SQLite。

附一支範例（先 `pip install matplotlib`）：

```bash
python examples/plot_trend.py "存款" --db fsc_news.sqlite   # 輸出 trend.png
```

## 下一步（需要實際檔案才能做）

`cells` 是通用長表，能容納任何格式。等本機跑過、看到真實 Excel 後，
建議再加一支「正規化」腳本，把 `cells` 轉成乾淨的寬表
（例如 `bank_deposits(period, bank, deposit, loan)`），分析會更順手。
各月格式對齊的邏輯通常是整個專案最花時間的部分。

## 專案結構

```
fsc/
  crawler.py     走訪頁面、找 Excel 連結（含反爬蟲 headers、重試）
  downloader.py  下載附件、sha256 去重
  parser.py      讀 Excel → 長表格子；檔名/標題推算年月
  database.py    SQLite schema 與寫入
  __main__.py    CLI：crawl / run / load
```