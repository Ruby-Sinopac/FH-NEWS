"""命令列入口。

用法：
  python -m fsc crawl   [--config config.yaml]   # 只列出找到的 Excel 連結（檢查用）
  python -m fsc run     [--config config.yaml]   # 爬取 + 下載 + 入庫（完整流程）
  python -m fsc load    [--config config.yaml]   # 只把 data/raw 下已下載的檔案入庫
"""

from __future__ import annotations

import argparse
import glob
import os
import sys

import yaml

from . import crawler, downloader, parser, database


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


def cmd_run(cfg: dict) -> None:
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
    ap.add_argument("command", choices=["crawl", "run", "load"], help="要執行的動作")
    ap.add_argument("--config", default="config.yaml", help="設定檔路徑")
    args = ap.parse_args(argv)

    cfg = _load_config(args.config)
    {"crawl": cmd_crawl, "run": cmd_run, "load": cmd_load}[args.command](cfg)


if __name__ == "__main__":
    main()
