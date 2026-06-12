#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""金控（含子公司）新聞 + 重大訊息每月彙整。

- 新聞：用公司名稱搜 Google News RSS，抓取後做篩選（雜訊/來源/名稱誤判）。
- 重訊：用公司代號查公開資訊觀測站(MOPS) t05st01 歷史重大訊息。
- 兩者放進「同一個金控分頁」，並一起依時間順序排列；被篩除的新聞另存分頁。

需求：pip install requests openpyxl pandas lxml
"""

import datetime
import io
import re
import time
import urllib.parse
import xml.etree.ElementTree as ET
from difflib import SequenceMatcher
from email.utils import parsedate_to_datetime

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 設定（要增減金控或子公司，改這裡的名稱清單即可）──────────────
# key 的開頭數字 = 金控股票代號，會用來查 MOPS 重訊。
GROUPS = {
    "2881 富邦金": ["富邦金控", "台北富邦銀行", "富邦人壽", "富邦產險", "富邦證券", "富邦投信"],
    "2882 國泰金": ["國泰金控", "國泰世華銀行", "國泰人壽", "國泰產險", "國泰證券", "國泰投信"],
    "2884 玉山金": ["玉山金控", "玉山銀行", "玉山證券", "玉山投信", "玉山創投", "三商美邦人壽"],
    "2885 元大金": ["元大金控", "元大證券", "元大銀行", "元大人壽", "元大投信", "元大期貨"],
    "2887 台新新光金": ["台新新光金控", "台新銀行", "新光銀行", "新光人壽", "台新證券", "元富證券", "台新投信"],
    "2890 永豐金": ["永豐金控", "永豐銀行", "永豐金證券", "永豐投信", "永豐期貨", "京城銀行"],
    "2891 中信金": ["中信金控", "中國信託銀行", "台灣人壽", "中國信託證券", "中國信託投信", "中信創投"],
}
# 旗下另有「單獨上市」、想單獨查重訊的子公司，填在這裡（金控key -> [代號,...]）：
EXTRA_MOPS_CODES = {
    # "2890 永豐金": ["2809"],   # 京城銀行
    # "2884 玉山金": ["2867"],   # 三商美邦人壽
}
DAYS_BACK = 30   # 往前抓幾天（新聞與重訊共用）
# 同一金控內，標題相似度 >= 此門檻視為「同一則內容」，只保留先抓到的一篇。
# 調高(接近1)=只併幾乎一樣的；調低=併得更兇（但可能誤併不同新聞）。
DEDUP_THRESHOLD = 0.55

# ── 新聞篩選設定（要鬆綁或加嚴，調整下面三組即可）────────────────
# (1) 雜訊關鍵字：標題只要含任一詞就剔除（股價/盤中閒聊、廣告/業配/活動）
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
    "商業周刊", "自由財經", "自由時報", "聯合新聞網", "聯合報", "udn", "UDN", "中時",
    "ETtoday", "Yahoo", "信傳媒", "風傳媒", "鏡週刊", "鏡報", "遠見",
    "三立", "TVBS", "東森", "NOWnews", "新頭殼", "上報", "民視", "公視",
    "數位時代", "金融",
]

# (3) 名稱誤判：標題必須真的出現公司名（或下方別名）才保留，避免同名誤判
REQUIRE_NAME_IN_TITLE = True
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

MOPS = "https://mopsov.twse.com.tw/mops/web/ajax_t05st01"
MOPS_PAGE = ("https://mopsov.twse.com.tw/mops/web/t05st01"
             "?co_id={co_id}&year={year}&step=1&firstin=ture&TYPEK=all")
MOPS_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Referer": "https://mopsov.twse.com.tw/mops/web/t05st01",
    "Origin": "https://mopsov.twse.com.tw",
    "Content-Type": "application/x-www-form-urlencoded",
}


# ── 新聞抓取 ───────────────────────────────────────────────
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
        out.append({"kind": "新聞", "entity": name, "date": d, "time": "",
                    "title": title, "source": source, "link": link})
    return out


# ── 新聞篩選 ───────────────────────────────────────────────
def is_noise(title):
    return any(kw in title for kw in NOISE_KEYWORDS)


def source_ok(source):
    if not SOURCE_WHITELIST:
        return True
    return any(s in source for s in SOURCE_WHITELIST)


def title_has_name(title, entity):
    if not REQUIRE_NAME_IN_TITLE:
        return True
    names = [entity] + ALIASES.get(entity, [])
    return any(n in title for n in names)


def norm_title(title):
    """正規化標題：去掉尾端「 - 來源」、移除標點與空白，方便比對是否同一則。"""
    t = title.rsplit(" - ", 1)[0]
    return re.sub(r"[\s\W_]+", "", t)


def is_similar(nt, kept_norms):
    """nt 與已保留的任一標題太像（同一則內容）→ True。"""
    for k in kept_norms:
        if nt == k or SequenceMatcher(None, nt, k).ratio() >= DEDUP_THRESHOLD:
            return True
    return False


def drop_reason(n):
    """回傳新聞篩除原因；若應保留則回傳空字串 ""。"""
    if is_noise(n["title"]):
        return "雜訊關鍵字"
    if not source_ok(n["source"]):
        return "非白名單來源"
    if not title_has_name(n["title"], n["entity"]):
        return "標題未出現公司名"
    return ""


# ── 重大訊息抓取（MOPS）─────────────────────────────────────
def roc_year(d):
    return d.year - 1911


def parse_roc_date(s):
    """民國日期字串（115/06/10）→ datetime.date；失敗回 None。"""
    try:
        y, m, d = (int(x) for x in str(s).strip().split("/"))
        return datetime.date(y + 1911, m, d)
    except Exception:
        return None


def fetch_material(co_id, name, year):
    """抓某公司某民國年的重大訊息，回傳 list[dict]（已轉成共用格式）。"""
    data = {"encodeURIComponent": "1", "step": "1", "firstin": "1", "off": "1",
            "TYPEK": "all", "co_id": co_id, "year": str(year)}
    r = requests.post(MOPS, data=data, headers=MOPS_HEADERS, timeout=30)
    r.raise_for_status()
    r.encoding = "utf-8"
    try:
        tables = pd.read_html(io.StringIO(r.text))
    except ValueError:
        return []

    target = None
    for t in tables:
        if any("主旨" in str(c) for c in t.columns):
            if target is None or len(t) > len(target):
                target = t
    if target is None:
        return []

    def pick(cols, *keys):
        for c in cols:
            if any(k in str(c) for k in keys):
                return c
        return None

    cols = list(target.columns)
    c_date = pick(cols, "發言日期", "日期")
    c_time = pick(cols, "發言時間", "時間")
    c_subj = pick(cols, "主旨")
    link = MOPS_PAGE.format(co_id=co_id, year=year)

    out = []
    for _, row in target.iterrows():
        subj = str(row.get(c_subj, "")).strip()
        if not subj or subj == "nan":
            continue
        out.append({"kind": "重訊", "entity": name,
                    "date": parse_roc_date(row.get(c_date, "")),
                    "time": str(row.get(c_time, "")).strip(),
                    "title": subj, "source": "公開資訊觀測站", "link": link})
    return out


# ── 排序鍵：新聞與重訊一起依時間排（新到舊）──
def sort_key(n):
    return (n["date"] or datetime.date.min, n["time"] or "")


# ── 輸出 Excel ─────────────────────────────────────────────
def build_excel(all_items, dropped_news, start_date, end_date, out_path):
    hf = PatternFill("solid", fgColor="1F4E78")
    sub_fill = PatternFill("solid", fgColor="FCE4D6")      # 重訊列底色
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
    ov["A1"] = "金控 新聞 + 重大訊息 彙整（已篩選）"
    ov["A1"].font = tfont
    ov["A2"] = f"資料區間：{start_date} ~ {end_date}"
    ov["A2"].font = lfont
    for c, h in enumerate(["金控", "新聞則數", "重訊則數", "合計"], start=1):
        cell = ov.cell(row=4, column=c, value=h)
        cell.fill, cell.font, cell.border = hf, hfont, bd
    r = 5
    for grp, items in all_items.items():
        n_news = sum(1 for n in items if n["kind"] == "新聞")
        n_mops = sum(1 for n in items if n["kind"] == "重訊")
        ov.cell(row=r, column=1, value=grp).border = bd
        ov.cell(row=r, column=2, value=n_news).border = bd
        ov.cell(row=r, column=3, value=n_mops).border = bd
        ov.cell(row=r, column=4, value=len(items)).border = bd
        r += 1
    for col, w in zip("ABCD", (18, 10, 10, 8)):
        ov.column_dimensions[col].width = w

    cols = [("日期", 12), ("類型", 7), ("公司", 15), ("標題／主旨", 60),
            ("來源", 18), ("連結", 46)]
    for grp, items in all_items.items():
        ws = wb.create_sheet(title=grp[:31])
        items = sorted(items, key=sort_key, reverse=True)
        ws["A1"] = grp
        ws["A1"].font = tfont
        n_news = sum(1 for n in items if n["kind"] == "新聞")
        n_mops = sum(1 for n in items if n["kind"] == "重訊")
        ws["A2"] = (f"資料區間：{start_date} ~ {end_date}"
                    f"（新聞 {n_news} 則、重訊 {n_mops} 則，依時間排序）")
        ws["A2"].font = lfont
        for c, (h, w) in enumerate(cols, start=1):
            cell = ws.cell(row=4, column=c, value=h)
            cell.fill, cell.font, cell.border = hf, hfont, bd
            ws.column_dimensions[get_column_letter(c)].width = w
        rr = 5
        if not items:
            ws.cell(row=rr, column=1, value="（本期間查無資料）")
        for n in items:
            is_mops = n["kind"] == "重訊"
            ws.cell(row=rr, column=1,
                    value=n["date"].isoformat() if n["date"] else "").border = bd
            ws.cell(row=rr, column=2, value=n["kind"]).border = bd
            ws.cell(row=rr, column=3, value=n["entity"]).border = bd
            c4 = ws.cell(row=rr, column=4, value=n["title"])
            c4.alignment, c4.border = wrap, bd
            ws.cell(row=rr, column=5, value=n["source"]).border = bd
            c6 = ws.cell(row=rr, column=6, value=n["link"])
            if n["link"]:
                c6.hyperlink, c6.font = n["link"], lkfont
            c6.border = bd
            if is_mops:                                   # 重訊整列上色，易辨識
                for c in range(1, 7):
                    if c != 6:
                        ws.cell(row=rr, column=c).fill = sub_fill
            rr += 1
        ws.freeze_panes = "A5"

    # ── 已篩除分頁（新聞，含篩除原因）──
    ws = wb.create_sheet(title="已篩除")
    dcols = [("日期", 12), ("金控", 16), ("公司", 15), ("標題", 56),
             ("來源", 18), ("篩除原因", 14), ("連結", 40)]
    ws["A1"] = "被篩除的新聞（自行檢查是否誤殺）"
    ws["A1"].font = tfont
    ws["A2"] = f"資料區間：{start_date} ~ {end_date}（共 {len(dropped_news)} 則）"
    ws["A2"].font = lfont
    for c, (h, w) in enumerate(dcols, start=1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.fill, cell.font, cell.border = hf, hfont, bd
        ws.column_dimensions[get_column_letter(c)].width = w
    rr = 5
    if not dropped_news:
        ws.cell(row=rr, column=1, value="（沒有任何新聞被篩除）")
    for n in sorted(dropped_news, key=sort_key, reverse=True):
        ws.cell(row=rr, column=1,
                value=n["date"].isoformat() if n["date"] else "").border = bd
        ws.cell(row=rr, column=2, value=n["group"]).border = bd
        ws.cell(row=rr, column=3, value=n["entity"]).border = bd
        c4 = ws.cell(row=rr, column=4, value=n["title"])
        c4.alignment, c4.border = wrap, bd
        ws.cell(row=rr, column=5, value=n["source"]).border = bd
        ws.cell(row=rr, column=6, value=n["reason"]).border = bd
        c7 = ws.cell(row=rr, column=7, value=n["link"])
        if n["link"]:
            c7.hyperlink, c7.font = n["link"], lkfont
        c7.border = bd
        rr += 1
    ws.freeze_panes = "A5"

    wb.save(out_path)


def main():
    today = datetime.date.today()
    cutoff = today - datetime.timedelta(days=DAYS_BACK)
    start_date, end_date = cutoff.isoformat(), today.isoformat()
    out_path = f"金控彙整_{today.strftime('%Y-%m')}.xlsx"
    years = sorted({roc_year(cutoff), roc_year(today)}, reverse=True)

    print(f"抓取區間：{start_date} ~ {end_date}（重訊民國年 {years}）")
    all_items = {}
    dropped_news = []
    for grp, names in GROUPS.items():
        print(f"• {grp}")
        seen, dseen, items, kept_norms = set(), set(), [], []
        grp_name = " ".join(grp.split()[1:]) or grp     # 去掉代號的金控名

        # 1) 新聞
        n_raw = n_drop = n_dup = 0
        for name in names:
            try:
                rows = fetch_google_news(name)
            except Exception as e:
                print(f"    新聞 {name} 失敗：{e}")
                rows = []
            for n in rows:
                if n["date"] and n["date"] < cutoff:
                    continue
                n_raw += 1
                key = (n["title"][:40], n["link"])
                reason = drop_reason(n)
                if reason:
                    n_drop += 1
                    if key not in dseen:
                        dseen.add(key)
                        dropped_news.append({**n, "group": grp, "reason": reason})
                    continue
                if key in seen:
                    continue
                nt = norm_title(n["title"])
                if is_similar(nt, kept_norms):      # 同內容多篇報導，只留一篇
                    n_dup += 1
                    continue
                seen.add(key)
                kept_norms.append(nt)
                items.append(n)
            time.sleep(0.5)
        print(f"    新聞：原始 {n_raw} → 篩掉 {n_drop} → 去重 {n_dup} → 保留 "
              f"{sum(1 for n in items if n['kind'] == '新聞')} 則")

        # 2) 重訊（金控代號 + 額外指定代號）
        codes = [grp.split()[0]] + EXTRA_MOPS_CODES.get(grp, [])
        m_cnt = 0
        for co_id in codes:
            for y in years:
                try:
                    rows = fetch_material(co_id, grp_name, y)
                except Exception as e:
                    print(f"    重訊 {co_id} 民國{y} 失敗：{e}")
                    rows = []
                for n in rows:
                    if n["date"] and n["date"] < cutoff:
                        continue
                    items.append(n)
                    m_cnt += 1
                time.sleep(0.8)
        print(f"    重訊：{m_cnt} 則")

        all_items[grp] = items

    build_excel(all_items, dropped_news, start_date, end_date, out_path)
    total = sum(len(v) for v in all_items.values())
    print(f"完成：{out_path}（合計 {total} 則，含已篩除 {len(dropped_news)} 則新聞另存分頁）")


if __name__ == "__main__":
    main()
