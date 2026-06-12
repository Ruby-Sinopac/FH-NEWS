#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""金控（含子公司）新聞每月彙整 — 用公司名稱搜尋，原始抓取後做篩選。
需求：pip install requests openpyxl"""

import datetime
import time
import urllib.parse
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 設定（要增減金控或子公司，改這裡的名稱清單即可）──────────────
GROUPS = {
    "2881 富邦金": ["富邦金控", "台北富邦銀行", "富邦人壽", "富邦產險", "富邦證券", "富邦投信"],
    "2882 國泰金": ["國泰金控", "國泰世華銀行", "國泰人壽", "國泰產險", "國泰證券", "國泰投信"],
    "2884 玉山金": ["玉山金控", "玉山銀行", "玉山證券", "玉山投信", "玉山創投", "三商美邦人壽"],
    "2885 元大金": ["元大金控", "元大證券", "元大銀行", "元大人壽", "元大投信", "元大期貨"],
    "2887 台新新光金": ["台新新光金控", "台新銀行", "新光銀行", "新光人壽", "台新證券", "元富證券", "台新投信"],
    "2890 永豐金": ["永豐金控", "永豐銀行", "永豐金證券", "永豐投信", "永豐期貨", "京城銀行"],
    "2891 中信金": ["中信金控", "中國信託銀行", "台灣人壽", "中國信託證券", "中國信託投信", "中信創投"],
}
DAYS_BACK = 30   # 往前抓幾天（以執行當下為基準）

# ── 篩選設定（要鬆綁或加嚴，調整下面三組即可）─────────────────────
# (1) 雜訊關鍵字：標題只要含任一詞就剔除（股價/盤中閒聊、廣告/業配/活動）
#     不想濾掉某詞，把它刪掉或在前面加 "#" 不適用，直接從清單移除即可。
NOISE_KEYWORDS = [
    # ── 股價／盤中閒聊 ──
    "盤中", "盤後", "盤前", "盤勢", "個股", "股價", "技術分析", "目標價",
    "外資買超", "外資賣超", "投信買超", "三大法人", "籌碼", "K線", "K棒",
    "漲停", "跌停", "周線", "月線", "均線", "殖利率排行", "存股",
    "台股開盤", "台股收盤", "熱門股", "強勢股", "飆股", "盤點", "看盤",
    # ── 廣告／業配／活動 ──
    "抽獎", "好康", "限時", "贈品", "折扣", "促銷", "體驗", "開箱",
    "贊助", "公益", "捐款", "路跑", "馬拉松", "籃球", "棒球", "球團", "球星",
    "記者會", "代言", "業配", "票選", "活動報名", "尾牙", "春酒", "抽刮刮樂",
]

# (2) 限定可信來源：留空 [] = 不限制；填了就「只保留」來源含清單內任一字串的新聞。
SOURCE_WHITELIST = [
    "經濟日報", "工商時報", "中央社", "鉅亨", "MoneyDJ", "財訊", "今周刊",
    "商業周刊", "自由財經", "自由時報", "聯合新聞網", "聯合報", "中時",
    "ETtoday", "Yahoo", "信傳媒", "風傳媒", "鏡週刊", "鏡報", "天下", "遠見",
    "三立", "TVBS", "東森", "NOWnews", "新頭殼", "上報", "民視", "公視",
    "數位時代", "金融",
]

# (3) 名稱誤判：標題必須真的出現公司名（或下方別名）才保留，避免同名誤判
#     （例：玉山＝山岳、國泰＝泛用詞，只在內文順帶提及的也會被剔除）
REQUIRE_NAME_IN_TITLE = True
# 別名（標題常用簡稱），會連同 GROUPS 內的全名一起比對
ALIASES = {
    "台北富邦銀行": ["北富銀", "富邦銀"],
    "中國信託銀行": ["中信銀", "中國信託", "中信"],
    "中國信託證券": ["中信證", "中國信託"],
    "中國信託投信": ["中信投信", "中國信託"],
    "中信金控": ["中信金", "中國信託金控"],
    "國泰世華銀行": ["國泰世華", "世華銀"],
    "玉山銀行": ["玉山銀"],
    "元大證券": ["元大證"],
    "永豐金證券": ["永豐金證", "大永豐"],
    "台新銀行": ["台新銀"],
    "新光人壽": ["新壽"],
    "台灣人壽": ["台壽"],
    "富邦人壽": ["富壽"],
    "國泰人壽": ["國壽"],
}
# ───────────────────────────────────────────────────────────

RSS = "https://news.google.com/rss/search?q={q}&hl=zh-TW&gl=TW&ceid=TW:zh-Hant"


def fetch_google_news(name):
    """用公司名稱搜尋 Google 新聞 RSS，回傳 list[dict]。"""
    q = urllib.parse.quote(f'"{name}" when:{DAYS_BACK}d')
    r = requests.get(RSS.format(q=q), timeout=30,
                     headers={"User-Agent": "Mozilla/5.0"})
    r.raise_for_status()
    root = ET.fromstring(r.content)
    out = []
    for item in root.findall("./channel/item"):
        title = item.findtext("title", "")
        link = item.findtext("link", "")
        pub = item.findtext("pubDate", "")
        src_el = item.find("{http://news.google.com}source") or item.find("source")
        source = src_el.text if src_el is not None else ""
        try:
            d = parsedate_to_datetime(pub).date()
        except Exception:
            d = None
        out.append({"entity": name, "date": d, "title": title, "source": source, "link": link})
    return out


# ── 篩選函式 ───────────────────────────────────────────────
def is_noise(title):
    """標題含雜訊關鍵字 → True（要剔除）。"""
    return any(kw in title for kw in NOISE_KEYWORDS)


def source_ok(source):
    """來源在白名單內（或白名單為空）→ True（保留）。"""
    if not SOURCE_WHITELIST:
        return True
    return any(s in source for s in SOURCE_WHITELIST)


def title_has_name(title, entity):
    """標題確實出現公司名或其別名 → True（保留）。"""
    if not REQUIRE_NAME_IN_TITLE:
        return True
    names = [entity] + ALIASES.get(entity, [])
    return any(n in title for n in names)


def keep_news(n):
    """綜合判斷一則新聞是否保留。"""
    if is_noise(n["title"]):
        return False
    if not source_ok(n["source"]):
        return False
    if not title_has_name(n["title"], n["entity"]):
        return False
    return True
# ───────────────────────────────────────────────────────────


def build_excel(all_news, start_date, end_date, out_path):
    hf = PatternFill("solid", fgColor="1F4E78")
    hfont = Font(color="FFFFFF", bold=True)
    tfont = Font(bold=True, size=14)
    lfont = Font(bold=True, color="555555")
    lkfont = Font(color="0563C1", underline="single")
    wrap = Alignment(wrap_text=True, vertical="top")
    thin = Side(style="thin", color="D9D9D9")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = Workbook()
    ov = wb.active
    ov.title = "總覽"
    ov["A1"] = "金控（含子公司）新聞彙整（已篩選）"
    ov["A1"].font = tfont
    ov["A2"] = f"資料區間：{start_date} ~ {end_date}"
    ov["A2"].font = lfont
    for c, h in enumerate(["金控", "新聞則數"], start=1):
        cell = ov.cell(row=4, column=c, value=h)
        cell.fill, cell.font, cell.border = hf, hfont, bd
    r = 5
    for grp, items in all_news.items():
        ov.cell(row=r, column=1, value=grp).border = bd
        ov.cell(row=r, column=2, value=len(items)).border = bd
        r += 1
    ov.column_dimensions["A"].width = 18
    ov.column_dimensions["B"].width = 10

    cols = [("日期", 12), ("公司", 15), ("標題", 60), ("來源", 18), ("連結", 46)]
    for grp, items in all_news.items():
        ws = wb.create_sheet(title=grp[:31])
        items = sorted(items, key=lambda n: (n["date"] or datetime.date.min), reverse=True)
        ws["A1"] = grp
        ws["A1"].font = tfont
        ws["A2"] = f"資料區間：{start_date} ~ {end_date}（共 {len(items)} 則，已篩選）"
        ws["A2"].font = lfont
        for c, (h, w) in enumerate(cols, start=1):
            cell = ws.cell(row=4, column=c, value=h)
            cell.fill, cell.font, cell.border = hf, hfont, bd
            ws.column_dimensions[get_column_letter(c)].width = w
        rr = 5
        if not items:
            ws.cell(row=rr, column=1, value="（本期間查無新聞）")
        for n in items:
            ws.cell(row=rr, column=1, value=n["date"].isoformat() if n["date"] else "").border = bd
            ws.cell(row=rr, column=2, value=n["entity"]).border = bd
            c3 = ws.cell(row=rr, column=3, value=n["title"])
            c3.alignment, c3.border = wrap, bd
            ws.cell(row=rr, column=4, value=n["source"]).border = bd
            c5 = ws.cell(row=rr, column=5, value=n["link"])
            if n["link"]:
                c5.hyperlink, c5.font = n["link"], lkfont
            c5.border = bd
            rr += 1
        ws.freeze_panes = "A5"
    wb.save(out_path)


def main():
    today = datetime.date.today()
    cutoff = today - datetime.timedelta(days=DAYS_BACK)
    start_date, end_date = cutoff.isoformat(), today.isoformat()
    out_path = f"金控新聞_{today.strftime('%Y-%m')}.xlsx"

    print(f"抓取區間：{start_date} ~ {end_date}")
    all_news = {}
    total_raw = total_kept = 0
    for grp, names in GROUPS.items():
        print(f"• {grp}")
        seen, items = set(), []
        raw_cnt = dropped = 0
        for name in names:
            try:
                rows = fetch_google_news(name)
            except Exception as e:
                print(f"    {name} 失敗：{e}")
                rows = []
            for n in rows:
                if n["date"] and n["date"] < cutoff:
                    continue                      # 超出區間
                raw_cnt += 1
                if not keep_news(n):
                    dropped += 1
                    continue                      # 篩選：剔除雜訊/非白名單/名稱誤判
                key = (n["title"][:40], n["link"])
                if key in seen:
                    continue                      # 去重
                seen.add(key)
                items.append(n)
            time.sleep(0.5)                       # 友善延遲
        all_news[grp] = items
        total_raw += raw_cnt
        total_kept += len(items)
        print(f"    原始 {raw_cnt} 則 → 篩掉 {dropped} 則 → 保留 {len(items)} 則")
    build_excel(all_news, start_date, end_date, out_path)
    print(f"完成：{out_path}（原始 {total_raw} 則 → 保留 {total_kept} 則）")


if __name__ == "__main__":
    main()
