"""解析 Excel 檔，並從檔名/標題推測「資料期間」（年月）。

金管會各月 Excel 的欄位、標題列、合併儲存格常隨年度變動，
直接套固定 schema 容易壞掉。因此預設採「無損長表」：
把每個工作表的每一格都展開成 (sheet, row, col, header, value)，
完整收進資料庫後再於 SQL 端做正規化。
"""

from __future__ import annotations

import io
import re
import zipfile
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

    支援：108年5月、西元 2019-05/201905、民國代碼
    （5碼補零如 11401/10704、4碼不補零如 1151/1074）。
    """
    for t in texts:
        if not t:
            continue
        # 108年5月 / 108年05月
        m = re.search(r"(\d{2,3})\s*年\s*(\d{1,2})\s*月", t)
        if m:
            ym = _roc_to_ym(int(m.group(1)), int(m.group(2)))
            if ym:
                return ym
        # 2019-05 / 2019/05 / 201905（西元）
        m = re.search(r"(20\d{2})[-/_]?(\d{2})(?!\d)", t)
        if m:
            return f"{m.group(1)}-{m.group(2)}"
        # 民國 5 碼：年(1xx) + 月(2碼)，例 11401→114年1月、10710→107年10月
        for mm in re.finditer(r"(?<!\d)(1\d{2})(\d{2})(?!\d)", t):
            ym = _roc_to_ym(int(mm.group(1)), int(mm.group(2)))
            if ym:
                return ym
        # 民國 4 碼：年(1xx) + 月(1碼)，例 1151→115年1月、1074→107年4月
        for mm in re.finditer(r"(?<!\d)(1\d{2})([1-9])(?!\d)", t):
            ym = _roc_to_ym(int(mm.group(1)), int(mm.group(2)))
            if ym:
                return ym
    return None


def _roc_to_ym(roc_year: int, month: int) -> str | None:
    if not (1 <= month <= 12):
        return None
    return f"{roc_year + 1911}-{month:02d}"


_TABLE_EXTS = (".xls", ".xlsx", ".xlsm", ".csv")


def read_cells(path: str) -> list[Cell]:
    """讀一個檔案，展開成長表的格子清單。

    支援 .xls/.xlsx/.xlsm/.csv，以及內含上述檔案的 .zip
    （金管會 fsc.gov.tw 的揭露報表多為 ZIP 包裝）。
    """
    if path.lower().endswith(".zip"):
        return _read_cells_from_zip(path)
    with open(path, "rb") as f:
        return _read_table_bytes(f.read(), path, sheet_prefix="")


def _read_cells_from_zip(path: str) -> list[Cell]:
    """解開 ZIP，讀其中每個 Excel/CSV，合併成格子清單。"""
    cells: list[Cell] = []
    found = 0
    try:
        zf = zipfile.ZipFile(path)
    except Exception as exc:
        raise RuntimeError(f"無法開啟 ZIP：{path}：{exc}") from exc
    with zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = _zip_member_name(info)
            base = name.rsplit("/", 1)[-1]
            if base.startswith(("__MACOSX", ".", "~$")):
                continue
            if not base.lower().endswith(_TABLE_EXTS):
                continue
            found += 1
            data = zf.read(info)
            # 以 ZIP 內檔名當 sheet 前綴，避免多檔工作表同名衝突
            cells.extend(_read_table_bytes(data, base, sheet_prefix=f"{base}::"))
    if found == 0:
        raise RuntimeError(f"ZIP 內沒有 Excel/CSV：{path}")
    return cells


def _zip_member_name(info: zipfile.ZipInfo) -> str:
    """還原 ZIP 成員檔名。台灣壓縮檔常用 cp950 而非 UTF-8。"""
    if info.flag_bits & 0x800:  # 已標記為 UTF-8
        return info.filename
    try:
        return info.filename.encode("cp437").decode("cp950")
    except Exception:
        return info.filename


# 真正標題列的判斷關鍵字（FSC 報表上方常有標題/單位/資料月份等說明列）
_HEADER_HINTS = ("名稱", "機構")


def _read_table_bytes(data: bytes, label: str, *, sheet_prefix: str) -> list[Cell]:
    """從位元組讀 Excel 或 CSV，回傳格子清單（自動跳過上方說明列、抓真正標題列）。"""
    low = label.lower()
    try:
        if low.endswith(".csv"):
            grids = {"CSV": pd.read_csv(io.BytesIO(data), header=None, dtype=str)}
        else:
            grids = pd.read_excel(
                io.BytesIO(data), sheet_name=None, header=None, dtype=str
            )
    except Exception as exc:
        raise RuntimeError(f"無法讀取 {label}：{exc}") from exc

    cells: list[Cell] = []
    for sheet_name, grid in grids.items():
        cells.extend(_grid_to_cells(grid, sheet_prefix, str(sheet_name)))
    return cells


def _find_header_row(grid: pd.DataFrame) -> int:
    """找真正的標題列：前 20 列中第一個含「名稱/機構」關鍵字的列；找不到回 0。"""
    for i in range(min(len(grid), 20)):
        vals = [str(x) for x in grid.iloc[i].tolist() if pd.notna(x)]
        if any(any(h in v for h in _HEADER_HINTS) for v in vals):
            return i
    return 0


def _grid_to_cells(grid: pd.DataFrame, sheet_prefix: str, sheet_name: str) -> list[Cell]:
    grid = grid.dropna(how="all").reset_index(drop=True)
    if grid.empty:
        return []
    hr = _find_header_row(grid)
    headers = [
        str(x).strip() if pd.notna(x) and str(x).strip() else f"第{c}欄"
        for c, x in enumerate(grid.iloc[hr].tolist())
    ]
    cells: list[Cell] = []
    for r, (_, series) in enumerate(grid.iloc[hr + 1:].iterrows()):
        for c, val in enumerate(series.tolist()):
            if pd.isna(val) or str(val).strip() == "":
                continue
            cells.append(
                Cell(
                    sheet=f"{sheet_prefix}{sheet_name}",
                    row=r,
                    col=c,
                    header=headers[c] if c < len(headers) else f"第{c}欄",
                    value=str(val).strip(),
                )
            )
    return cells
