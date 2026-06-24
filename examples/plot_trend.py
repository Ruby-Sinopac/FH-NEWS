"""範例：從 SQLite 讀資料畫趨勢圖。

前置：
    pip install matplotlib
    python -m fsc run          # 先把資料抓進 fsc_news.sqlite

用法：
    python examples/plot_trend.py "存款" --db fsc_news.sqlite

說明：
    這支只是示範「SQLite → 圖」的最後一哩。它會找出 header 含關鍵字的欄位，
    把各月份的數值加總後畫折線圖。實務上等你看到真實 Excel、做完正規化寬表後，
    查詢會更精準（例如直接 SELECT period, bank, deposit）。
"""

from __future__ import annotations

import argparse
import sqlite3

import pandas as pd
import matplotlib

matplotlib.use("Agg")  # 無視窗環境也能輸出圖檔
import matplotlib.pyplot as plt
from matplotlib import font_manager


def _setup_cjk_font() -> None:
    """挑一個系統上有的中文字型，避免標題顯示成方框。"""
    candidates = [
        "Microsoft JhengHei", "PingFang TC", "Heiti TC", "Noto Sans CJK TC",
        "Noto Sans CJK SC", "Source Han Sans TW", "WenQuanYi Zen Hei", "SimHei",
    ]
    available = {f.name for f in font_manager.fontManager.ttflist}
    for name in candidates:
        if name in available:
            plt.rcParams["font.sans-serif"] = [name]
            plt.rcParams["axes.unicode_minus"] = False
            return
    # 找不到中文字型也沒關係，月份軸仍是 ASCII，可正常閱讀


def main() -> None:
    _setup_cjk_font()
    ap = argparse.ArgumentParser()
    ap.add_argument("keyword", help="要畫的欄位關鍵字，例如『存款』")
    ap.add_argument("--db", default="fsc_news.sqlite")
    ap.add_argument("--out", default="trend.png")
    args = ap.parse_args()

    conn = sqlite3.connect(args.db)
    df = pd.read_sql_query(
        """
        SELECT f.period AS period, c.value AS value
        FROM cells c JOIN source_files f ON f.id = c.file_id
        WHERE c.header LIKE ? AND f.period IS NOT NULL
        """,
        conn,
        params=(f"%{args.keyword}%",),
    )
    conn.close()

    if df.empty:
        print(f"查無 header 含「{args.keyword}」的資料。先確認已 run 過、且欄位名稱對。")
        return

    # 數值清洗（去逗號），加總後依月份排序
    df["value"] = pd.to_numeric(
        df["value"].str.replace(",", "", regex=False), errors="coerce"
    )
    series = df.dropna(subset=["value"]).groupby("period")["value"].sum().sort_index()

    ax = series.plot(marker="o", figsize=(10, 5))
    ax.set_title(f"金管會月報趨勢：{args.keyword}")
    ax.set_xlabel("月份")
    ax.set_ylabel("加總值")
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(args.out, dpi=120)
    print(f"已輸出圖檔：{args.out}（{len(series)} 個月份）")


if __name__ == "__main__":
    main()
