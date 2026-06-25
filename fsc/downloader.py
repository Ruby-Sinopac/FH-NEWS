"""下載 Excel 附件並存到本機，重複檔案以 sha256 去重。"""

from __future__ import annotations

import hashlib
import os
import re
import time
from dataclasses import dataclass
from urllib.parse import urlparse, parse_qs, unquote

import requests


class NotFound(Exception):
    """伺服器回應 404/410：該檔不存在（例如該月份尚未發布）。"""


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
    # 下載程式型網址（如 fckdowndoc?file=/xxx.zip&flag=doc）：真檔名在 file 參數
    q = parse_qs(urlparse(url).query)
    if q.get("file"):
        fb = os.path.basename(q["file"][0])
        if fb:
            return unquote(fb)
    # 退而求其次，用 query 參數組出名字
    serno = (q.get("dataserno") or q.get("serno") or ["file"])[0]
    return f"{serno}.xlsx"


def _sanitize(name: str) -> str:
    name = name.replace("/", "_").replace("\\", "_")
    return re.sub(r"[^\w.\-()（）一-鿿]+", "_", name).strip("_") or "file"


def sniff(content: bytes) -> str:
    """用開頭位元組判斷內容類型：zip / ole2(舊版xls) / html / pdf / unknown。"""
    head = content[:512]
    if head[:2] == b"PK":
        return "zip"          # .xlsx 也是 zip，這裡統稱 zip
    if head[:4] == b"\xd0\xcf\x11\xe0":
        return "ole2"         # 舊版 .xls
    if head[:4] == b"%PDF":
        return "pdf"
    low = head.lstrip().lower()
    if low.startswith((b"<!doctype", b"<html", b"<?xml", b"<head", b"<meta")):
        return "html"
    return "unknown"


def download(
    session: requests.Session,
    url: str,
    raw_dir: str,
    *,
    retries: int = 6,
    delay: float = 1.5,
    referer: str | None = None,
    html_retries: int = 3,
) -> Downloaded:
    """下載單一檔案到 raw_dir。內容若與既有檔相同則跳過寫入。

    若伺服器以 200 回傳 HTML（軟性 404，該檔其實不存在），視為 NotFound。
    """
    os.makedirs(raw_dir, exist_ok=True)
    base = f"{urlparse(url).scheme}://{urlparse(url).netloc}/"
    req_headers = {
        "Referer": referer or base,
        "Accept": (
            "application/zip,application/vnd.ms-excel,"
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,"
            "application/octet-stream,*/*;q=0.8"
        ),
    }  # 其餘 Sec-Fetch / sec-ch-ua 等沿用 session 預設（導覽情境）
    backoff = 2.0
    html_hits = 0
    last_exc: Exception | None = None
    for attempt in range(retries):
        try:
            resp = session.get(url, timeout=120, headers=req_headers)
            if resp.status_code in (404, 410):
                # 檔案不存在，不需重試
                raise NotFound(url)
            resp.raise_for_status()
            content = resp.content
            # 政府網站/WAF 可能對程式請求回傳 200 + HTML（擋檔或軟性 404）。
            # 多重試幾次、間隔逐步拉長（WAF 有時擋一兩次後會放行）；
            # 連續多次仍是 HTML 才視為不存在。
            if sniff(content) == "html":
                html_hits += 1
                if html_hits >= html_retries:
                    raise NotFound(url)
                time.sleep(delay * (html_hits + 1))
                continue
            sha = hashlib.sha256(content).hexdigest()
            filename = _sanitize(_guess_filename(url, resp))
            # 保險：副檔名與實際內容不符時補正（避免把 zip 當 xlsx 解析）
            if not filename.lower().endswith(
                (".zip", ".xlsx", ".xlsm", ".xls", ".csv", ".ods")
            ):
                kind = sniff(content)
                filename += {"zip": ".zip", "ole2": ".xls", "pdf": ".pdf"}.get(kind, ".zip")
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
        except requests.exceptions.SSLError as exc:
            last_exc = exc
            if session.verify:
                from .crawler import disable_ssl_verify

                print("    ⚠ 下載時憑證驗證失敗，自動改用『不驗證憑證』重試…")
                disable_ssl_verify(session)
                continue
            if attempt < retries - 1:
                time.sleep(backoff)
                backoff *= 2
        except requests.RequestException as exc:  # noqa: PERF203
            last_exc = exc
            if attempt < retries - 1:
                time.sleep(backoff)
                backoff *= 2
    raise RuntimeError(f"下載失敗 {url}：{last_exc}")
