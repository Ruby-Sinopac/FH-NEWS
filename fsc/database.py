"""SQLite 入庫：檔案登記表 + 無損長表。"""

from __future__ import annotations

import datetime as _dt
import sqlite3
from contextlib import contextmanager

from .parser import Cell

_SCHEMA = """
CREATE TABLE IF NOT EXISTS source_files (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    url           TEXT,
    filename      TEXT NOT NULL,
    sha256        TEXT NOT NULL UNIQUE,
    period        TEXT,            -- 'YYYY-MM'，可能為 NULL
    size_bytes    INTEGER,
    downloaded_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS cells (
    file_id  INTEGER NOT NULL REFERENCES source_files(id) ON DELETE CASCADE,
    sheet    TEXT NOT NULL,
    row      INTEGER NOT NULL,
    col      INTEGER NOT NULL,
    header   TEXT,
    value    TEXT,
    PRIMARY KEY (file_id, sheet, row, col)
);

CREATE INDEX IF NOT EXISTS idx_cells_header ON cells(header);
CREATE INDEX IF NOT EXISTS idx_files_period ON source_files(period);
"""


@contextmanager
def connect(db_path: str):
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db(conn: sqlite3.Connection) -> None:
    conn.executescript(_SCHEMA)


def file_exists(conn: sqlite3.Connection, sha256: str) -> bool:
    cur = conn.execute("SELECT 1 FROM source_files WHERE sha256 = ?", (sha256,))
    return cur.fetchone() is not None


def insert_file(
    conn: sqlite3.Connection,
    *,
    url: str,
    filename: str,
    sha256: str,
    period: str | None,
    size_bytes: int,
    cells: list[Cell],
) -> int:
    """寫入一個檔案及其所有格子，回傳 file_id。"""
    now = _dt.datetime.now().isoformat(timespec="seconds")
    cur = conn.execute(
        """INSERT INTO source_files (url, filename, sha256, period, size_bytes, downloaded_at)
           VALUES (?, ?, ?, ?, ?, ?)""",
        (url, filename, sha256, period, size_bytes, now),
    )
    file_id = cur.lastrowid
    conn.executemany(
        """INSERT OR IGNORE INTO cells (file_id, sheet, row, col, header, value)
           VALUES (?, ?, ?, ?, ?, ?)""",
        [(file_id, c.sheet, c.row, c.col, c.header, c.value) for c in cells],
    )
    return file_id
