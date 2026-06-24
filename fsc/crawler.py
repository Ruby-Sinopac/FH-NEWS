"""走訪金管會頁面，找出 Excel 附件連結與（可選的）其他月份公告連結。

金管會網站（banking.gov.tw）有基本的反爬蟲，且為 JSP + session 架構，
因此這裡：
  - 帶上接近真實瀏覽器的 headers
  - 使用同一個 requests.Session 以保留 cookie
  - 失敗時做指數退避重試
"""

from __future__ import annotations

import time
from dataclasses import dataclass
from urllib.parse import urljoin, urlparse, parse_qs

import requests
from bs4 import BeautifulSoup

# Excel 副檔名
_EXCEL_EXTS = (".xls", ".xlsx", ".xlsm", ".csv")

# 常見的瀏覽器標頭，降低被 403 擋掉的機率
_DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;q=0.9,"
        "image/avif,image/webp,*/*;q=0.8"
    ),
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}


@dataclass
class Link:
    """頁面上找到的一個連結。"""

    url: str
    text: str


def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(_DEFAULT_HEADERS)
    return s


def fetch_html(session: requests.Session, url: str, *, retries: int = 4) -> str:
    """抓網頁 HTML，附帶 Referer 與指數退避重試。"""
    headers = {"Referer": f"{urlparse(url).scheme}://{urlparse(url).netloc}/"}
    delay = 2.0
    last_exc: Exception | None = None
    for attempt in range(retries):
        try:
            resp = session.get(url, headers=headers, timeout=30)
            resp.raise_for_status()
            # 金管會頁面為 UTF-8，但偶有未宣告編碼的情況
            resp.encoding = resp.apparent_encoding or "utf-8"
            return resp.text
        except requests.RequestException as exc:  # noqa: PERF203
            last_exc = exc
            if attempt < retries - 1:
                time.sleep(delay)
                delay *= 2
    raise RuntimeError(f"無法取得頁面 {url}：{last_exc}")


def _looks_like_excel(href: str, text: str) -> bool:
    low = href.lower()
    if any(low.split("?")[0].endswith(ext) for ext in _EXCEL_EXTS):
        return True
    # 金管會附件常透過下載參數提供，連結本身不帶副檔名，
    # 改以連結文字判斷（檔名/標題多半含 xls 或「Excel」字樣）。
    tl = text.lower()
    if any(ext.lstrip(".") in tl for ext in _EXCEL_EXTS) or "excel" in tl:
        return True
    if "download" in low or "dl.jsp" in low or "/multiplehtml/" in low:
        return True
    return False


def find_excel_links(html: str, base_url: str) -> list[Link]:
    """從頁面 HTML 找出所有疑似 Excel 附件的連結（去重）。"""
    soup = BeautifulSoup(html, "lxml")
    seen: set[str] = set()
    out: list[Link] = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if href.startswith(("javascript:", "mailto:", "#")):
            continue
        text = a.get_text(strip=True)
        if not _looks_like_excel(href, text):
            continue
        full = urljoin(base_url, href)
        if full in seen:
            continue
        seen.add(full)
        out.append(Link(url=full, text=text or full))
    return out


def find_article_links(html: str, base_url: str) -> list[Link]:
    """找出連到「其他月份公告」的連結（multimessage_view.jsp 且帶 dataserno）。"""
    soup = BeautifulSoup(html, "lxml")
    seen: set[str] = set()
    out: list[Link] = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        full = urljoin(base_url, href)
        q = parse_qs(urlparse(full).query)
        if "dataserno" in q and "multimessage_view" in full:
            if full in seen:
                continue
            seen.add(full)
            out.append(Link(url=full, text=a.get_text(strip=True) or full))
    return out
