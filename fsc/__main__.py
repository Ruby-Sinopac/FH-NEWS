"""命令列入口。

用法：
  python -m fsc run     [--config config.yaml]   # 下載 + 入庫（完整流程）
  python -m fsc crawl   [--config config.yaml]   # 只列出找到的 Excel 連結（檢查用）
  python -m fsc load    [--config config.yaml]   # 只把 data/raw 下已下載的檔案入庫

run 有兩種模式，依設定檔自動選擇：
  - 有 url_template → 「網址範本」模式：照年月套網址逐月下載（適用 fsc.gov.tw）
  - 否則           → 「爬頁面」模式：從 start_urls 找 Excel 連結
"""

from __future__ import annotations

import argparse
import datetime as _dt
import glob
import os
import sys
import time
from urllib.parse import urlparse

import yaml

from . import crawler, downloader, parser, database


def _iter_periods(start: list[int], end: list[int] | None) -> list[tuple[str, int, int]]:
    """產生 (‘YYYY-MM’, 民國年, 月) 清單，從 start 到 end（含）。

    end 為 None 時自動取「今天」的民國年月。
    """
    sy, sm = start
    if end:
        ey, em = end
    else:
        today = _dt.date.today()
        ey, em = today.year - 1911, today.month
    out: list[tuple[str, int, int]] = []
    y, m = sy, sm
    while (y, m) <= (ey, em):
        out.append((f"{y + 1911}-{m:02d}", y, m))
        m += 1
        if m > 12:
            m, y = 1, y + 1
    return out


def _period_codes(y: int, m: int) -> list[str]:
    """一個月份的候選代碼：不補零與補零兩種（FSC 不同年份規則不一致）。

    例：107年4月 → ['1074', '10704']；114年10月 → ['11410']（兩者相同，去重）。
    """
    return list(dict.fromkeys([f"{y}{m}", f"{y}{m:02d}"]))


def _build_url(pattern: str, code: str) -> str:
    """把 {period} 代入並把路徑中的 & 編成 %26（與瀏覽器一致；無 query 才處理）。"""
    url = pattern.replace("{period}", code)
    if "?" not in url:
        url = url.replace("&", "%26")
    return url


def _template_patterns(tpl: dict) -> list[str]:
    """取得網址範本清單。支援單一 pattern 或多個 patterns（檔名變體）。"""
    pats: list[str] = []
    if tpl.get("patterns"):
        pats = list(tpl["patterns"])
    elif tpl.get("pattern"):
        pats = [tpl["pattern"]]
    if not pats:
        sys.exit("url_template 需要 pattern 或 patterns。")
    for p in pats:
        if "{period}" not in p:
            sys.exit(f"範本必須包含 {{period}} 佔位符：{p}")
    return pats


def _run_template(cfg: dict) -> None:
    """網址範本模式：照年月逐一套網址下載並入庫。

    每個月份會依序嘗試多個檔名變體（patterns），第一個抓到真檔案的就採用。
    """
    tpl = cfg["url_template"]
    patterns = _template_patterns(tpl)
    periods = _iter_periods(tpl.get("start", [107, 4]), tpl.get("end"))
    raw_dir = cfg.get("raw_dir", "data/raw")
    db_path = cfg.get("db_path", "fsc_news.sqlite")
    delay = float(cfg.get("request_delay", 1.5))
    session = crawler.make_session(verify_ssl=cfg.get("verify_ssl", True))
    base = f"{urlparse(patterns[0]).scheme}://{urlparse(patterns[0]).netloc}/"
    crawler.warm_up(session, base)

    print(f"範本模式：{len(periods)} 個月份（{periods[0][1]} ~ {periods[-1][1]}）"
          f"，每月嘗試 {len(patterns)} 種網址 × 補零/不補零代碼")
    with database.connect(db_path) as conn:
        database.init_db(conn)
        new = skip = miss = fail = 0
        for ym, y, mth in periods:
            codes = _period_codes(y, mth)
            dl = None
            url = ""
            for pat in patterns:
                for code in codes:
                    url = _build_url(pat, code)
                    try:
                        # 多樣式嘗試：HTML（軟性404）快速換下一個；html_retries=1
                        dl = downloader.download(
                            session, url, raw_dir, delay=delay, html_retries=1
                        )
                        break
                    except downloader.NotFound:
                        time.sleep(delay)  # 換下一個樣式前稍候，對伺服器客氣
                        continue
                    except RuntimeError as exc:
                        fail += 1
                        print(f"  ✗ {ym}：{exc}")
                        dl = "FAIL"
                        break
                if dl is not None:
                    break
            if dl == "FAIL":
                continue
            if dl is None:
                miss += 1
                print(f"  · {ym}（{'/'.join(codes)}）不存在，略過")
                continue
            if database.file_exists(conn, dl.sha256):
                skip += 1
                print(f"  = {ym} 已存在，跳過")
                continue
            try:
                cells = parser.read_cells(dl.path)
            except RuntimeError as exc:
                fail += 1
                print(f"  ✗ {ym} 解析失敗：{exc}")
                continue
            database.insert_file(
                conn, url=url, filename=dl.filename, sha256=dl.sha256,
                period=ym, size_bytes=dl.size, cells=cells,
            )
            new += 1
            print(f"  ✓ {ym}  {dl.filename}  格子數={len(cells)}")
        print(f"\n完成：新增 {new}、已存在 {skip}、不存在 {miss}、失敗 {fail}。資料庫：{db_path}")


def _load_config(path: str) -> dict:
    if not os.path.exists(path):
        sys.exit(f"找不到設定檔 {path}（可從 config.example.yaml 複製一份）")
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def _collect_excel_links(cfg: dict) -> list[crawler.Link]:
    session = crawler.make_session(verify_ssl=cfg.get("verify_ssl", True))
    start_urls: list[str] = list(cfg.get("start_urls", []))
    follow = cfg.get("follow_article_links", False)
    delay = float(cfg.get("request_delay", 1.5))

    pages = list(start_urls)
    visited: set[str] = set()
    excel_links: dict[str, crawler.Link] = {}

    import time

    while pages:
        url = pages.pop(0)
        if url in visited:
            continue
        visited.add(url)
        print(f"  ▶ 讀取頁面：{url}")
        try:
            html = crawler.fetch_html(session, url)
        except RuntimeError as exc:
            print(f"    ⚠ {exc}")
            continue

        for link in crawler.find_excel_links(html, url):
            excel_links.setdefault(link.url, link)

        if follow:
            for art in crawler.find_article_links(html, url):
                if art.url not in visited:
                    pages.append(art.url)
        time.sleep(delay)

    return list(excel_links.values())


def cmd_crawl(cfg: dict) -> None:
    links = _collect_excel_links(cfg)
    print(f"\n找到 {len(links)} 個 Excel 連結：")
    for ln in links:
        print(f"  - {ln.text}\n    {ln.url}")
    if not links:
        print("  （沒找到。請確認頁面是否能存取，或把『列表頁』加進 start_urls，"
              "並視情況開啟 follow_article_links）")


def cmd_inspect(cfg: dict, *, url: str | None, period: str | None) -> None:
    """診斷單一網址：印出狀態、標頭、cookie 與 HTML 內文前段，找出被擋原因。"""
    import requests

    if not url:
        tpl = cfg.get("url_template")
        if not (tpl and period):
            sys.exit("請給 --url <網址>，或 --period <期間代碼，如 11408>（需設定 url_template）。")
        pats = _template_patterns(tpl)
        url = _build_url(pats[0], period)
        if len(pats) > 1:
            print("（注意：此期間有多種檔名變體，inspect 只檢視第 1 種；"
                  "要檢視其他種請用 --url 指定完整網址）\n")

    session = crawler.make_session(verify_ssl=cfg.get("verify_ssl", True))
    base = f"{urlparse(url).scheme}://{urlparse(url).netloc}/"
    crawler.warm_up(session, base)
    headers = {"Referer": base, "Accept": "application/zip,application/octet-stream,*/*;q=0.8"}

    print(f"網址：{url}\n")
    try:
        resp = session.get(url, timeout=60, headers=headers)
    except requests.exceptions.SSLError:
        crawler.disable_ssl_verify(session)
        resp = session.get(url, timeout=60, headers=headers)

    content = resp.content
    print(f"HTTP 狀態：{resp.status_code}")
    print(f"內容類型：{resp.headers.get('Content-Type', '-')}")
    print(f"內容長度：{resp.headers.get('Content-Length', len(content))} bytes（實收 {len(content)}）")
    print(f"判定：{downloader.sniff(content)}")
    print(f"session cookie：{dict(session.cookies)}")
    print("\n— 回應標頭 —")
    for k, v in resp.headers.items():
        print(f"  {k}: {v}")
    print("\n— 內文前 800 字 —")
    try:
        text = content.decode(resp.apparent_encoding or "utf-8", errors="replace")
    except Exception:
        text = repr(content[:800])
    print(text[:800])


def cmd_probe(cfg: dict) -> None:
    """診斷：逐月檢查 url_template 的每個網址實際回傳什麼（不寫入資料庫）。"""
    tpl = cfg.get("url_template")
    if not tpl:
        sys.exit("probe 只適用 url_template 模式，請先在 config.yaml 設定 url_template。")
    import time

    patterns = _template_patterns(tpl)
    periods = _iter_periods(tpl.get("start", [107, 4]), tpl.get("end"))
    delay = float(cfg.get("request_delay", 1.5))
    session = crawler.make_session(verify_ssl=cfg.get("verify_ssl", True))
    base = f"{urlparse(patterns[0]).scheme}://{urlparse(patterns[0]).netloc}/"
    crawler.warm_up(session, base)

    print(f"探測 {len(periods)} 個月份（{periods[0][1]} ~ {periods[-1][1]}）"
          f"，每月嘗試 {len(patterns)} 種網址 × 補零/不補零：\n")
    good: list[str] = []
    first = True
    for ym, y, mth in periods:
        codes = _period_codes(y, mth)
        hit = None
        last = (0, "", "", "-", 0, "ERR")
        for vi, pat in enumerate(patterns):
            for code in codes:
                if not first:
                    time.sleep(delay)
                first = False
                status, ctype, size, kind = _probe_one(session, _build_url(pat, code))
                last = (vi, code, status, ctype, size, kind)
                if kind in ("zip", "ole2"):
                    hit = last
                    break
            if hit:
                break
        vi, code, status, ctype, size, kind = hit or last
        if kind in ("zip", "ole2"):
            good.append(ym)
        flag = {"zip": "✓ 真檔案", "ole2": "✓ 舊版xls", "html": "✗ 找不到(回首頁)",
                "pdf": "PDF", "unknown": "? 未知", "ERR": "✗ 連線失敗"}.get(kind, kind)
        tag = f"樣式{vi + 1}/{code}" if len(patterns) > 1 else code
        print(f"  {ym}  HTTP {status}  {ctype or '-':<20} "
              f"{size:>9} bytes  {flag}  [{tag}]")
    print(f"\n可用月份（{len(good)}）：", "、".join(good) if good else "（無）")
    if not good:
        print("→ 沒有任何月份回傳真檔案。請把上面幾行與某個月份的真實下載連結貼給我，"
              "我據此調整 patterns。")


def _probe_one(session, url: str):
    """回傳 (status, content_type, size, kind)。"""
    import requests

    base = f"{urlparse(url).scheme}://{urlparse(url).netloc}/"
    headers = {"Referer": base, "Accept": "application/zip,application/octet-stream,*/*;q=0.8"}
    for _ in range(2):
        try:
            resp = session.get(url, timeout=30, stream=True, headers=headers)
            ctype = resp.headers.get("Content-Type", "").split(";")[0]
            chunk = next(resp.iter_content(2048), b"") or b""
            clen = resp.headers.get("Content-Length")
            size = int(clen) if clen and clen.isdigit() else len(chunk)
            resp.close()
            if resp.status_code in (404, 410):
                return resp.status_code, ctype, size, "html"
            return resp.status_code, ctype, size, downloader.sniff(chunk)
        except requests.exceptions.SSLError:
            if session.verify:
                crawler.disable_ssl_verify(session)
                continue
            return "-", "", 0, "ERR"
        except requests.RequestException:
            return "-", "", 0, "ERR"
    return "-", "", 0, "ERR"


def cmd_run(cfg: dict) -> None:
    if cfg.get("url_template"):
        return _run_template(cfg)
    raw_dir = cfg.get("raw_dir", "data/raw")
    db_path = cfg.get("db_path", "fsc_news.sqlite")
    delay = float(cfg.get("request_delay", 1.5))

    links = _collect_excel_links(cfg)
    print(f"\n共 {len(links)} 個 Excel 連結，開始下載 + 入庫…")
    session = crawler.make_session(verify_ssl=cfg.get("verify_ssl", True))

    with database.connect(db_path) as conn:
        database.init_db(conn)
        new, skip, fail = 0, 0, 0
        for ln in links:
            try:
                dl = downloader.download(session, ln.url, raw_dir, delay=delay)
            except RuntimeError as exc:
                print(f"  ✗ {exc}")
                fail += 1
                continue
            if database.file_exists(conn, dl.sha256):
                skip += 1
                continue
            try:
                cells = parser.read_cells(dl.path)
            except RuntimeError as exc:
                print(f"  ✗ 解析失敗：{exc}")
                fail += 1
                continue
            period = parser.guess_period(dl.filename, ln.text)
            database.insert_file(
                conn,
                url=ln.url,
                filename=dl.filename,
                sha256=dl.sha256,
                period=period,
                size_bytes=dl.size,
                cells=cells,
            )
            new += 1
            print(f"  ✓ {dl.filename}  期間={period or '未知'}  格子數={len(cells)}")
        print(f"\n完成：新增 {new}、跳過(已存在) {skip}、失敗 {fail}。資料庫：{db_path}")


def cmd_load(cfg: dict) -> None:
    """把 raw_dir 下已存在的 Excel 直接入庫（不連網）。"""
    import hashlib

    raw_dir = cfg.get("raw_dir", "data/raw")
    db_path = cfg.get("db_path", "fsc_news.sqlite")
    files = sorted(
        glob.glob(os.path.join(raw_dir, "*.xls*"))
        + glob.glob(os.path.join(raw_dir, "*.csv"))
    )
    print(f"在 {raw_dir} 找到 {len(files)} 個檔案。")

    with database.connect(db_path) as conn:
        database.init_db(conn)
        new, skip, fail = 0, 0, 0
        for path in files:
            with open(path, "rb") as f:
                sha = hashlib.sha256(f.read()).hexdigest()
            if database.file_exists(conn, sha):
                skip += 1
                continue
            try:
                cells = parser.read_cells(path)
            except RuntimeError as exc:
                print(f"  ✗ {exc}")
                fail += 1
                continue
            fn = os.path.basename(path)
            period = parser.guess_period(fn)
            database.insert_file(
                conn, url="", filename=fn, sha256=sha, period=period,
                size_bytes=os.path.getsize(path), cells=cells,
            )
            new += 1
            print(f"  ✓ {fn}  期間={period or '未知'}  格子數={len(cells)}")
        print(f"\n完成：新增 {new}、跳過 {skip}、失敗 {fail}。資料庫：{db_path}")


def main(argv: list[str] | None = None) -> None:
    ap = argparse.ArgumentParser(prog="fsc", description="金管會月報 Excel 爬取入庫工具")
    ap.add_argument("command", choices=["run", "probe", "inspect", "crawl", "load"],
                    help="動作（probe=逐月診斷；inspect=檢視單一網址完整回應）")
    ap.add_argument("--config", default="config.yaml", help="設定檔路徑")
    ap.add_argument("--url", help="inspect 用：要檢視的完整網址")
    ap.add_argument("--period", help="inspect 用：期間代碼（如 11408），依範本組出網址")
    args = ap.parse_args(argv)

    cfg = _load_config(args.config)
    if args.command == "inspect":
        return cmd_inspect(cfg, url=args.url, period=args.period)
    {"run": cmd_run, "probe": cmd_probe, "crawl": cmd_crawl,
     "load": cmd_load}[args.command](cfg)


if __name__ == "__main__":
    main()
