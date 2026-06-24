"""下載 Excel 附件並存到本機，重複檔案以 sha256 去重。"""

from __future__ import annotations

import hashlib
import os
import re
import time
from dataclasses import dataclass
from urllib.parse import urlparse, parse_qs, unquote

import requests


@dataclass
class Downloaded:
    url: str
    path: str
    filename: str
    sha256: str
    size: int
    skipped: bool  # 之前已下載過（內容相同）


def _guess_filename(url: str, resp: requests.Response) -> str:
    """從 Content-Disposition 或 URL 推測檔名。"""
    cd = resp.headers.get("Content-Disposition", "")
    m = re.search(r"filename\*?=(?:UTF-8'')?\"?([^\";]+)", cd)
    if m:
        return unquote(m.group(1).strip())
    path = urlparse(url).path
    base = os.path.basename(path)
    if base and "." in base:
        return unquote(base)
    # 退而求其次，用 query 參數組出名字
    q = parse_qs(urlparse(url).query)
    serno = (q.get("dataserno") or q.get("serno") or ["file"])[0]
    return f"{serno}.xlsx"


def _sanitize(name: str) -> str:
    name = name.replace("/", "_").replace("\\", "_")
    return re.sub(r"[^\w.\-()（）一-鿿]+", "_", name).strip("_") or "file"


def download(
    session: requests.Session,
    url: str,
    raw_dir: str,
    *,
    retries: int = 4,
    delay: float = 1.5,
) -> Downloaded:
    """下載單一檔案到 raw_dir。內容若與既有檔相同則跳過寫入。"""
    os.makedirs(raw_dir, exist_ok=True)
    backoff = 2.0
    last_exc: Exception | None = None
    for attempt in range(retries):
        try:
            resp = session.get(url, timeout=120)
            resp.raise_for_status()
            content = resp.content
            sha = hashlib.sha256(content).hexdigest()
            filename = _sanitize(_guess_filename(url, resp))
            path = os.path.join(raw_dir, filename)

            # 若同名檔已存在且內容相同 → 跳過
            if os.path.exists(path):
                with open(path, "rb") as f:
                    if hashlib.sha256(f.read()).hexdigest() == sha:
                        return Downloaded(url, path, filename, sha, len(content), True)
                # 同名但內容不同 → 加 hash 前綴避免覆蓋
                stem, ext = os.path.splitext(filename)
                filename = f"{stem}_{sha[:8]}{ext}"
                path = os.path.join(raw_dir, filename)

            with open(path, "wb") as f:
                f.write(content)
            time.sleep(delay)
            return Downloaded(url, path, filename, sha, len(content), False)
        except requests.RequestException as exc:  # noqa: PERF203
            last_exc = exc
            if attempt < retries - 1:
                time.sleep(backoff)
                backoff *= 2
    raise RuntimeError(f"下載失敗 {url}：{last_exc}")
