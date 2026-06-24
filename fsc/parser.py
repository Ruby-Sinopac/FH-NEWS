"""解析 Excel 檔，並從檔名/標題推測「資料期間」（年月）。

金管會各月 Excel 的欄位、標題列、合併儲存格常隨年度變動，
直接套固定 schema 容易壞掉。因此預設採「無損長表」：
把每個工作表的每一格都展開成 (sheet, row, col, header, value)，
完整收進資料庫後再於 SQL 端做正規化。
"""

from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd


@dataclass
class Cell:
    sheet: str
    row: int       # 0-based 資料列索引（不含標題列）
    col: int       # 0-based 欄索引
    header: str    # 該欄的標題（取第一列）
    value: str


def guess_period(*texts: str) -> str | None:
    """從檔名或標題猜測資料期間，回傳正規化的 'YYYY-MM'。

    支援：民國 10805 / 108年5月 / 108-05、西元 201905 / 2019-05 等。
    """
    for t in texts:
        if not t:
            continue
        # 108年5月 / 108年05月
        m = re.search(r"(\d{2,3})\s*年\s*(\d{1,2})\s*月", t)
        if m:
            return _roc_to_ym(int(m.group(1)), int(m.group(2)))
        # 2019-05 / 2019/05 / 201905（西元）
        m = re.search(r"(20\d{2})[-/_]?(\d{2})(?!\d)", t)
        if m:
            return f"{m.group(1)}-{m.group(2)}"
        # 10805 / 108-05（民國，5~6 碼）
        m = re.search(r"(?<!\d)(\d{3})[-/_]?(\d{2})(?!\d)", t)
        if m:
            return _roc_to_ym(int(m.group(1)), int(m.group(2)))
    return None


def _roc_to_ym(roc_year: int, month: int) -> str | None:
    if not (1 <= month <= 12):
        return None
    return f"{roc_year + 1911}-{month:02d}"


def read_cells(path: str) -> list[Cell]:
    """讀一個 Excel 的所有工作表，展開成長表的格子清單。"""
    cells: list[Cell] = []
    try:
        sheets = pd.read_excel(path, sheet_name=None, header=0, dtype=str)
    except Exception as exc:  # 檔案毀損或非 Excel
        raise RuntimeError(f"無法讀取 Excel：{path}：{exc}") from exc

    for sheet_name, df in sheets.items():
        df = df.dropna(how="all").dropna(axis=1, how="all")
        if df.empty:
            continue
        headers = [str(c) for c in df.columns]
        for r, (_, series) in enumerate(df.iterrows()):
            for c, val in enumerate(series.tolist()):
                if pd.isna(val) or str(val).strip() == "":
                    continue
                cells.append(
                    Cell(
                        sheet=str(sheet_name),
                        row=r,
                        col=c,
                        header=headers[c] if c < len(headers) else "",
                        value=str(val).strip(),
                    )
                )
    return cells
