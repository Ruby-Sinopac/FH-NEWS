"""把 SQLite 裡的長表資料整理成「往右長」的寬表 Excel。

預設版面（--sheet-by 指定「類別」欄）：
  - 每個「類別」的值 = 一個工作表
  - 每列 = 一個機構（--label 指定的欄，或第 label_col 欄）
  - 各欄 = 指標 × 期間，期間由左到右遞增 → 月份往右長

若未指定 --sheet-by，則改以「原始分頁名稱」當工作表。
數值若多數可轉成數字，會以數字型態輸出，方便排序與畫圖。
"""

from __future__ import annotations

import re
import sqlite3
from collections import OrderedDict

import pandas as pd


def _to_number(v):
    if v is None:
        return None
    s = str(v).strip().replace(",", "").replace("%", "").replace("　", "")
    if s in ("", "-", "—", "N/A", "n/a", "na"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _safe_sheet_name(name: str, used: set[str]) -> str:
    name = re.sub(r"[\[\]:*?/\\]", " ", str(name)).strip() or "Sheet"
    name = name[:31]
    base, i = name, 1
    while name in used:
        suffix = f"_{i}"
        name = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(name)
    return name


def _logical(sheet: str) -> str:
    return sheet.split("::")[-1]


def _fetch_rows(db_path, period_from, period_to):
    conn = sqlite3.connect(db_path)
    rows = conn.execute(
        """
        SELECT f.period, c.sheet, c.row, c.col, c.header, c.value
        FROM cells c JOIN source_files f ON f.id = c.file_id
        WHERE f.period IS NOT NULL
        ORDER BY f.period, c.sheet, c.row, c.col
        """
    ).fetchall()
    conn.close()

    def keep(p):
        return (period_from is None or p >= period_from) and (
            period_to is None or p <= period_to
        )

    return [r for r in rows if keep(r[0])]


def list_columns(db_path: str, period_from=None, period_to=None) -> None:
    """列出各原始分頁有哪些欄位（header）與範例值，協助選 --sheet-by / --label。"""
    rows = _fetch_rows(db_path, period_from, period_to)
    if not rows:
        raise RuntimeError("資料庫沒有資料（請先 load 或 run）。")
    # 每個邏輯分頁 → 每個 header → 範例值（依欄序）
    sheets: "OrderedDict[str, OrderedDict[int, tuple]]" = OrderedDict()
    for period, sheet, r, col, header, val in rows:
        ls = _logical(sheet)
        cols = sheets.setdefault(ls, OrderedDict())
        if col not in cols:
            cols[col] = (header or f"第{col}欄", [])
        if val and len(cols[col][1]) < 3 and val not in cols[col][1]:
            cols[col][1].append(val)
    for ls, cols in sheets.items():
        print(f"\n■ 分頁「{ls}」的欄位：")
        for col, (header, samples) in sorted(cols.items()):
            ex = "、".join(str(s) for s in samples)
            print(f"   第{col}欄  {header}    例：{ex}")


def export(
    db_path: str,
    out_path: str,
    *,
    sheet_by: str | None = None,
    label: str | None = None,
    label_col: int = 0,
    period_from: str | None = None,
    period_to: str | None = None,
) -> dict:
    """輸出寬表 Excel。回傳統計摘要 dict。"""
    rows = _fetch_rows(db_path, period_from, period_to)
    if not rows:
        raise RuntimeError("篩選後沒有資料可匯出（先 load/run，或檢查期間範圍）。")

    # 把同一原始列的 cells 聚成一筆：by_header / by_col
    rowcells: "OrderedDict[tuple, dict]" = OrderedDict()
    for period, sheet, r, col, header, val in rows:
        d = rowcells.setdefault((period, sheet, r), {"h": {}, "c": {}})
        if header:
            d["h"][header] = val
        d["c"][col] = val

    records = []  # (category, entity, metric, period, value)
    for (period, sheet, r), d in rowcells.items():
        # 類別（工作表）
        if sheet_by:
            cat = d["h"].get(sheet_by)
        else:
            cat = _logical(sheet)
        cat = str(cat).strip() if cat not in (None, "") else "(未分類)"
        # 機構（列）
        if label and label in d["h"]:
            entity = d["h"].get(label)
        else:
            entity = d["c"].get(label_col)
        if entity is None or str(entity).strip() == "":
            continue
        entity = str(entity).strip()
        # 指標（其餘欄位）
        for header, val in d["h"].items():
            if header in (sheet_by, label):
                continue
            if val is None or str(val).strip() == "":
                continue
            records.append((cat, entity, header, period, val))

    if not records:
        raise RuntimeError(
            "找不到可整理的資料。請用 `python -m fsc columns` 確認欄位名稱，"
            "再以 --sheet-by / --label 指定正確的『類別』與『機構名稱』欄。"
        )

    df = pd.DataFrame(records, columns=["cat", "entity", "metric", "period", "value"])
    periods = sorted(df["period"].unique())
    metric_order = list(dict.fromkeys(df["metric"]))

    sheets_written = 0
    used: set[str] = set()
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        for cat in dict.fromkeys(df["cat"]):
            g = df[df["cat"] == cat]
            nums = g["value"].map(_to_number)
            use_num = nums.notna().mean() >= 0.6
            g = g.assign(v=nums if use_num else g["value"])
            wide = g.pivot_table(
                index="entity", columns=["metric", "period"], values="v",
                aggfunc="first",
            )
            # 欄序：指標（依出現順序）為外層、期間（遞增）為內層 → 月份往右長
            cols = [(m, p) for m in metric_order for p in periods
                    if (m, p) in wide.columns]
            wide = wide.reindex(columns=pd.MultiIndex.from_tuples(cols))
            wide = wide.reindex(list(dict.fromkeys(g["entity"])))  # 機構保留出現順序
            wide.index.name = label or "機構名稱"
            ws = _safe_sheet_name(cat, used)
            wide.to_excel(xw, sheet_name=ws)
            sheets_written += 1

    return {
        "out": out_path,
        "sheets": sheets_written,
        "periods": periods,
        "rows": len(records),
        "categories": list(dict.fromkeys(df["cat"])),
    }
