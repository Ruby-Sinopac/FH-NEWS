"""把 SQLite 裡的長表資料整理成「往右長」的寬表 Excel。

版面：
  - 每個「指標欄位（header）」= 一個工作表
  - 每列 = 一個機構（取每列的標籤欄，預設第 0 欄）
  - 各欄 = 各期間（YYYY-MM），由左到右遞增 → 月份往右長

數值若多數可轉成數字，會以數字型態輸出（去除千分位逗號），方便排序與畫圖。
"""

from __future__ import annotations

import re
import sqlite3

import pandas as pd


def _to_number(v):
    """嘗試把字串轉成數字（去逗號、百分號、空白）；失敗回 None。"""
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
    """Excel 工作表名稱限制：≤31 字、不可含 []:*?/\\，且不可重複。"""
    name = re.sub(r"[\[\]:*?/\\]", " ", str(name)).strip() or "Sheet"
    name = name[:31]
    base, i = name, 1
    while name in used:
        suffix = f"_{i}"
        name = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(name)
    return name


def export(
    db_path: str,
    out_path: str,
    *,
    label_col: int = 0,
    period_from: str | None = None,
    period_to: str | None = None,
) -> dict:
    """讀 db_path，輸出寬表 Excel 到 out_path。回傳統計摘要 dict。"""
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

    if not rows:
        raise RuntimeError("資料庫沒有可匯出的資料（請先用 load 或 run 入庫）。")

    # 期間篩選
    def in_range(p: str) -> bool:
        return (period_from is None or p >= period_from) and (
            period_to is None or p <= period_to
        )

    # 第一輪：建立每列的「機構標籤」對照（period, sheet, row）→ 標籤欄的值
    labels: dict[tuple, str] = {}
    for period, sheet, r, col, header, val in rows:
        if col == label_col:
            labels[(period, sheet, r)] = val

    # 是否有多個邏輯工作表（:: 後的真實分頁名）→ 有的話指標前面加分頁名避免混淆
    logical_sheets = {s.split("::")[-1] for _, s, _, _, _, _ in rows}
    multi = len(logical_sheets) > 1

    records = []  # (metric, entity, period, value)
    for period, sheet, r, col, header, val in rows:
        if col == label_col or not in_range(period):
            continue
        entity = labels.get((period, sheet, r))
        if entity is None or str(entity).strip() == "":
            continue
        metric = (header or f"第{col}欄").strip()
        if multi:
            metric = f"{sheet.split('::')[-1]}-{metric}"
        records.append((metric, str(entity).strip(), period, val))

    if not records:
        raise RuntimeError("篩選後沒有資料可匯出（檢查期間範圍或標籤欄設定）。")

    df = pd.DataFrame(records, columns=["metric", "entity", "period", "value"])
    periods = sorted(df["period"].unique())

    sheets_written = 0
    used: set[str] = set()
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        for metric in dict.fromkeys(df["metric"]):  # 保留出現順序
            g = df[df["metric"] == metric]
            nums = g["value"].map(_to_number)
            use_num = nums.notna().mean() >= 0.6  # 多數可轉數字 → 用數字
            gg = g.assign(v=nums if use_num else g["value"])
            wide = gg.pivot_table(
                index="entity", columns="period", values="v", aggfunc="first"
            )
            # 機構順序：保留第一次出現順序；期間：由小到大（往右遞增）
            wide = wide.reindex(list(dict.fromkeys(g["entity"])))
            wide = wide.reindex([p for p in periods if p in wide.columns], axis=1)
            wide.index.name = "機構名稱"
            ws = _safe_sheet_name(metric, used)
            wide.to_excel(xw, sheet_name=ws)
            sheets_written += 1

    return {
        "out": out_path,
        "sheets": sheets_written,
        "periods": periods,
        "rows": len(records),
    }
