# FH-NEWS — 金管會月報 Excel 入庫工具

把金管會每月公布的報表（Excel / ZIP）自動抓下來，彙整進一個
**SQLite** 資料庫，方便後續查詢、畫圖與分析。

支援兩種抓取模式（`run` 依設定檔自動選擇）：

- **網址範本模式**（建議）：對 `fsc.gov.tw` 這種**固定命名規則**的檔案，
  照「年月」直接套網址逐月下載，免爬頁面。ZIP 會自動解壓讀裡面的 Excel/CSV。
- **爬頁面模式**：對 `banking.gov.tw` 這種把連結掛在公告頁的情況，
  從 `start_urls` 找出 Excel 連結再下載。

> ⚠️ **網路限制**：Claude Code on the web 的雲端環境預設**擋住政府網域**，
> 所以 `run` / `crawl` 需在**你自己的本機**執行。純解析入庫（`load`）免連網。

## 安裝

```bash
pip install -r requirements.txt
cp config.example.yaml config.yaml   # Windows: copy config.example.yaml config.yaml
```

## 設定（config.yaml）

**網址範本模式**（電子支付帳戶揭露報表，已預設好）：

| 欄位 | 說明 |
|------|------|
| `url_template.patterns` | 網址範本清單，用 `{period}` 代表年月。**不同月份檔名後綴不同**，逐月依序嘗試，第一個抓到真檔案的就用。 |
| `url_template.start` | 起始 `[民國年, 月]`，例 `[107, 4]`。 |
| `url_template.end` | 結束 `[民國年, 月]`；留 `null` 代表自動抓到今天。 |

> `{period}` = 民國年接月份、**月份不補零**：107年4月→`1074`、107年10月→`10710`。
> 不存在的月份（伺服器回首頁的軟性 404）會自動略過。
> 已知檔名變體：`電子支付帳戶重要資訊揭露` 與 `電子支付帳戶重要資訊揭露(ODS&xlsx格式)`；
> 若仍有月份抓不到，用 `python -m fsc inspect --url <真實網址>` 找出新樣式再加進 `patterns`。

**爬頁面模式**：`start_urls`、`follow_article_links`。
**共用**：`raw_dir`、`db_path`、`request_delay`、`verify_ssl`。

## 使用

```bash
# 完整流程：下載 → （ZIP 解壓）→ 入庫
python -m fsc run --config config.yaml

# 爬頁面模式可先檢查抓不抓得到連結（不下載）
python -m fsc crawl --config config.yaml

# 若已手動把檔案放進 data/raw，只做入庫（免連網）
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