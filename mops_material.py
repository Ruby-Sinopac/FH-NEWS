#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""公開資訊觀測站(MOPS) — 各金控重大訊息彙整。
需求：pip install requests pandas lxml openpyxl

說明：
- 重大訊息用「公司代號」查；子公司的重訊一般由金控母公司代為公告，
  所以預設只查 7 家金控代號即可涵蓋。若旗下有單獨上市的子公司，
  在 COMPANIES 自行加上代號。
- MOPS 回傳的是 HTML 表格，這裡用 pandas.read_html 解析。
- 一定要帶瀏覽器 User-Agent，否則會被回 403。
"""

import datetime
import io
import time

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 設定 ───────────────────────────────────────────────────
COMPANIES = {
    "2881": "富邦金",
    "2882": "國泰金",
    "2884": "玉山金",
    "2885": "元大金",
    "2887": "台新新光金",
    "2890": "永豐金",
    "2891": "中信金",
    # 旗下單獨上市的子公司要查就加在這裡，例如：
    # "2809": "京城銀行",
    # "2867": "三商美邦人壽",
}
DAYS_BACK = 30   # 往前抓幾天的重訊
# ───────────────────────────────────────────────────────────

MOPS = "https://mopsov.twse.com.tw/mops/web/ajax_t05st01"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Referer": "https://mopsov.twse.com.tw/mops/web/t05st01",
    "Origin": "https://mopsov.twse.com.tw",
    "Content-Type": "application/x-www-form-urlencoded",
}


def roc_year(d):
    """西元日期 → 民國年。"""
    return d.year - 1911


def parse_roc_date(s):
    """民國日期字串（115/06/10 或 115/6/10）→ datetime.date；失敗回 None。"""
    try:
        y, m, d = (int(x) for x in str(s).strip().split("/"))
        return datetime.date(y + 1911, m, d)
    except Exception:
        return None


def fetch_material(co_id, year):
    """抓某公司某民國年的重大訊息，回傳 list[dict]：{date, time, subject}。"""
    data = {
        "encodeURIComponent": "1", "step": "1", "firstin": "1", "off": "1",
        "TYPEK": "all", "co_id": co_id, "year": str(year),
    }
    r = requests.post(MOPS, data=data, headers=HEADERS, timeout=30)
    r.raise_for_status()
    r.encoding = "utf-8"
    try:
        tables = pd.read_html(io.StringIO(r.text))
    except ValueError:
        return []                                  # 查無表格 = 沒資料

    # 找出含「主旨」欄、且筆數最多的那張表
    target = None
    for t in tables:
        cols = [str(c) for c in t.columns]
        if any("主旨" in c for c in cols):
            if target is None or len(t) > len(target):
                target = t
    if target is None:
        return []

    # 欄名容錯（MOPS 偶有微調）
    def pick(cols, *keys):
        for c in cols:
            if any(k in str(c) for k in keys):
                return c
        return None

    cols = list(target.columns)
    c_date = pick(cols, "發言日期", "日期")
    c_time = pick(cols, "發言時間", "時間")
    c_subj = pick(cols, "主旨")

    out = []
    for _, row in target.iterrows():
        subj = str(row.get(c_subj, "")).strip()
        if not subj or subj == "nan":
            continue
        out.append({
            "date": parse_roc_date(row.get(c_date, "")),
            "time": str(row.get(c_time, "")).strip(),
            "subject": subj,
        })
    return out


def build_excel(all_rows, start_date, end_date, out_path):
    hf = PatternFill("solid", fgColor="1F4E78")
    hfont = Font(color="FFFFFF", bold=True)
    tfont = Font(bold=True, size=14)
    lfont = Font(bold=True, color="555555")
    wrap = Alignment(wrap_text=True, vertical="top")
    thin = Side(style="thin", color="D9D9D9")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = Workbook()
    ws = wb.active
    ws.title = "重大訊息"
    ws["A1"] = "金控重大訊息彙整（公開資訊觀測站）"
    ws["A1"].font = tfont
    ws["A2"] = f"資料區間：{start_date} ~ {end_date}（共 {len(all_rows)} 則）"
    ws["A2"].font = lfont

    cols = [("代號", 8), ("公司", 14), ("發言日期", 12), ("時間", 8), ("主旨", 80)]
    for c, (h, w) in enumerate(cols, start=1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.fill, cell.font, cell.border = hf, hfont, bd
        ws.column_dimensions[get_column_letter(c)].width = w

    rr = 5
    if not all_rows:
        ws.cell(row=rr, column=1, value="（本期間查無重大訊息）")
    for n in sorted(all_rows, key=lambda x: (x["date"] or datetime.date.min, x["time"]),
                    reverse=True):
        ws.cell(row=rr, column=1, value=n["co_id"]).border = bd
        ws.cell(row=rr, column=2, value=n["name"]).border = bd
        ws.cell(row=rr, column=3,
                value=n["date"].isoformat() if n["date"] else "").border = bd
        ws.cell(row=rr, column=4, value=n["time"]).border = bd
        c5 = ws.cell(row=rr, column=5, value=n["subject"])
        c5.alignment, c5.border = wrap, bd
        rr += 1
    ws.freeze_panes = "A5"
    wb.save(out_path)


def main():
    today = datetime.date.today()
    cutoff = today - datetime.timedelta(days=DAYS_BACK)
    start_date, end_date = cutoff.isoformat(), today.isoformat()
    out_path = f"金控重訊_{today.strftime('%Y-%m')}.xlsx"

    # 區間可能橫跨民國年（例如一月初），把涉及的民國年都查一遍
    years = sorted({roc_year(cutoff), roc_year(today)}, reverse=True)

    print(f"抓取區間：{start_date} ~ {end_date}（民國年 {years}）")
    all_rows = []
    for co_id, name in COMPANIES.items():
        print(f"• {co_id} {name}")
        got = 0
        for y in years:
            try:
                rows = fetch_material(co_id, y)
            except Exception as e:
                print(f"    民國{y}年 失敗：{e}")
                rows = []
            for n in rows:
                if n["date"] and n["date"] < cutoff:
                    continue                       # 超出區間
                all_rows.append({**n, "co_id": co_id, "name": name})
                got += 1
            time.sleep(0.8)                        # 友善延遲
        print(f"    區間內 {got} 則")
    build_excel(all_rows, start_date, end_date, out_path)
    print(f"完成：{out_path}（共 {len(all_rows)} 則）")


if __name__ == "__main__":
    main()
