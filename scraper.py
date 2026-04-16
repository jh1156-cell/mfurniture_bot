from __future__ import annotations

import json
import re
from dataclasses import dataclass, asdict
from typing import Optional
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}


@dataclass
class ProductInfo:
    url: str
    name: str
    kind: str
    brand: str = "Magic-furniture"
    model: str = ""
    size: str = ""
    image_url: str = ""

    def to_dict(self) -> dict:
        return asdict(self)


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


def abs_image_url(url: str, base_url: str) -> str:
    url = (url or "").strip()
    if not url:
        return ""
    if url.startswith("//"):
        return "https:" + url
    return urljoin(base_url, url)


def classify_kind(name: str, url: str = "") -> str:
    s = f"{name} {url}".lower()
    if any(k in s for k in ["sofa", "쇼파", "소파"]):
        return "SOFA"
    if any(k in s for k in ["table", "테이블", "/12/"]):
        return "TABLE"
    if any(k in s for k in ["chair", "체어", "암체어", "의자", "/24/"]):
        return "CHAIR"
    if any(k in s for k in ["stool", "스툴"]):
        return "STOOL"
    if any(k in s for k in ["bench", "벤치"]):
        return "BENCH"
    return "PRODUCT"


def _extract_name(soup: BeautifulSoup) -> str:
    for selector in [
        "meta[property='og:title']",
        "meta[name='twitter:title']",
        "title",
    ]:
        node = soup.select_one(selector)
        if not node:
            continue
        value = node.get("content") if node.name == "meta" else node.get_text(" ", strip=True)
        value = normalize_space(value)
        if value:
            value = re.sub(r"\s*\|.*$", "", value)
            return value

    for node in soup.find_all(["h1", "h2", "h3", "strong", "span", "div"]):
        txt = normalize_space(node.get_text(" ", strip=True))
        if txt and len(txt) < 80 and not any(x in txt for x in ["원", "배송", "상품상세", "리뷰"]):
            return txt
    return ""


def _extract_image_url(soup: BeautifulSoup, url: str) -> str:
    meta = soup.select_one("meta[property='og:image']")
    if meta and meta.get("content"):
        return abs_image_url(meta["content"], url)

    for img in soup.select("img"):
        src = img.get("src") or img.get("data-src") or img.get("ec-data-src")
        src = abs_image_url(src or "", url)
        if src and any(k in src.lower() for k in ["/product/", "big/", ".jpg", ".png", ".webp"]):
            return src
    return ""


def _extract_size_from_text(text: str) -> str:
    patterns = [
        r"사이즈\s*[:：]?\s*([WwDdHh×xX0-9\-~ .,/A-Za-z가-힣]+)",
        r"SIZE\s*[:：]?\s*([WwDdHh×xX0-9\-~ .,/A-Za-z가-힣]+)",
        r"규격\s*[:：]?\s*([WwDdHh×xX0-9\-~ .,/A-Za-z가-힣]+)",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            candidate = normalize_space(m.group(1))
            candidate = re.split(r"(?:배송비|정기결제|옵션|COLOR|최소주문수량|총 상품금액)", candidate)[0].strip()
            if len(candidate) >= 4 and any(c in candidate for c in ["W", "D", "H", "x", "×"]):
                return candidate.replace(" x ", " × ").replace(" X ", " × ")
    return ""


def _extract_size_from_pairs(text: str) -> str:
    # W560 x D520 x H750 x SH440 x AH650 같은 형태
    m = re.search(
        r"((?:[A-Z]{1,3}\s*\d+(?:[.,]\d+)?\s*(?:[x×]\s*)?){2,8})",
        text,
        re.IGNORECASE,
    )
    if not m:
        return ""
    candidate = normalize_space(m.group(1))
    if all(token not in candidate.upper() for token in ["W", "D", "H"]):
        return ""
    candidate = re.sub(r"\s*[xX]\s*", " × ", candidate)
    candidate = re.sub(r"\s+", " ", candidate).strip()
    return candidate


def _extract_size(soup: BeautifulSoup) -> str:
    page_text = normalize_space(soup.get_text(" ", strip=True))
    size = _extract_size_from_text(page_text)
    if size:
        return size
    return _extract_size_from_pairs(page_text)


def scrape_product(url: str, timeout: int = 25) -> ProductInfo:
    response = requests.get(url, headers=HEADERS, timeout=timeout)
    response.raise_for_status()
    response.encoding = response.encoding or "utf-8"

    soup = BeautifulSoup(response.text, "html.parser")

    name = _extract_name(soup)
    image_url = _extract_image_url(soup, url)
    size = _extract_size(soup)
    model = name
    kind = classify_kind(name, url)

    return ProductInfo(
        url=url,
        name=name or "",
        kind=kind,
        model=model or "",
        size=size or "",
        image_url=image_url or "",
    )


def parse_links(raw_text: str) -> list[str]:
    links = []
    for line in (raw_text or "").splitlines():
        line = line.strip()
        if not line:
            continue
        m = re.search(r"https?://\S+", line)
        if m:
            links.append(m.group(0).rstrip(",)"))
    return links
