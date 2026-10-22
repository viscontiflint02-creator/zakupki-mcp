"""Ориентировочные цены из интернета: поиск по названию позиции и цены из выдачи.

Работает на ПК сотрудника (там есть интернет). Ищем через HTML-версию DuckDuckGo без
ключей и регистрации. Цены вытаскиваются из заголовков и сниппетов выдачи; это
подсказка для специалиста, а не коммерческое предложение: ссылку всегда показываем.
"""
from __future__ import annotations

import html
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import parse_qs, unquote, urlsplit

import httpx

SEARCH_URL = "https://html.duckduckgo.com/html/"
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")
RESULT_PATTERN = re.compile(r'<a[^>]+class="result__a"[^>]+href="(?P<href>[^"]+)"[^>]*>(?P<title>.*?)</a>', re.S)
SNIPPET_PATTERN = re.compile(r'class="result__snippet"[^>]*>(?P<snippet>.*?)</a>', re.S)
# «12 345 ₽», «12345,50 руб.», «от 990 р.» — пробелы-разделители тысяч, в том числе неразрывные.
PRICE_PATTERN = re.compile(r"(?<![\d.,])(\d{1,3}(?:[   ]\d{3})+|\d+)(?:[.,](\d{1,2}))?\s?(?:₽|руб\.?|р\.)",
                           re.I)


def _text(fragment: str) -> str:
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", "", fragment or ""))).strip()


def _target(href: str) -> str:
    """DuckDuckGo заворачивает ссылки в //duckduckgo.com/l/?uddg=<настоящий адрес>."""
    href = html.unescape(href)
    if href.startswith("//"):
        href = "https:" + href
    query = parse_qs(urlsplit(href).query)
    return unquote(query["uddg"][0]) if "uddg" in query else href


def extract_prices(text: str) -> list[float]:
    prices = []
    for whole, fraction in PRICE_PATTERN.findall(text):
        value = float(re.sub(r"\D", "", whole) + (f".{fraction}" if fraction else ""))
        if value > 0:
            prices.append(value)
    return prices


def parse_results(page: str, limit: int, pattern: re.Pattern = None, snippet_pattern: re.Pattern = None) -> list[dict]:
    pattern, snippet_pattern = pattern or RESULT_PATTERN, snippet_pattern or SNIPPET_PATTERN
    results = []
    matches = list(pattern.finditer(page))
    for index, match in enumerate(matches):
        # Сниппет ищем только до следующего результата, чтобы не взять чужой.
        end = matches[index + 1].start() if index + 1 < len(matches) else len(page)
        found = snippet_pattern.search(page, match.end(), end)
        title, snippet = _text(match["title"]), _text(found["snippet"] if found else "")
        url = _target(match["href"])
        if not url.startswith("http") or "duckduckgo.com/y.js" in url:
            continue  # рекламные блоки
        results.append({"title": title, "url": url, "snippet": snippet,
                        "prices_rub": extract_prices(f"{title} {snippet}")})
        if len(results) >= limit:
            break
    return results


STRUCTURED_PRICE_PATTERNS = [
    # JSON-LD / microdata schema.org Offer, OpenGraph product:price.
    re.compile(r'"price"\s*:\s*"?(\d+(?:[.,]\d{1,2})?)"?', re.I),
    re.compile(r'itemprop=["\']price["\'][^>]*content=["\'](\d+(?:[.,]\d{1,2})?)', re.I),
    re.compile(r'content=["\'](\d+(?:[.,]\d{1,2})?)["\'][^>]*itemprop=["\']price["\']', re.I),
    re.compile(r'property=["\'](?:product|og):price:amount["\'][^>]*content=["\'](\d+(?:[.,]\d{1,2})?)', re.I),
]
MAX_PAGE_BYTES = 2_000_000


def page_prices(page: str) -> tuple[list[float], str]:
    """Цены со страницы товара: сначала разметка schema.org, иначе первые цены из текста."""
    structured = []
    for pattern in STRUCTURED_PRICE_PATTERNS:
        structured += [float(value.replace(",", ".")) for value in pattern.findall(page)]
    structured = [value for value in structured if value > 0]
    if structured:
        return sorted(set(structured))[:5], "разметка товара"
    body = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", page)
    return extract_prices(_text(body))[:5], "текст страницы"


def _visit(client: httpx.Client, item: dict) -> None:
    try:
        with client.stream("GET", item["url"]) as response:
            if response.status_code != 200 or "html" not in response.headers.get("content-type", ""):
                item["page"] = f"не открылась (HTTP {response.status_code})"
                return
            chunks, size = [], 0
            for chunk in response.iter_bytes():
                chunks.append(chunk)
                size += len(chunk)
                if size > MAX_PAGE_BYTES:
                    break
            page = b"".join(chunks).decode(response.encoding or "utf-8", "replace")
    except httpx.HTTPError as error:
        item["page"] = f"не открылась ({error.__class__.__name__})"
        return
    prices, source = page_prices(page)
    if prices:
        item["prices_rub"] = prices
        item["page"] = f"цены со страницы: {source}"
    else:
        item["page"] = "цены на странице не найдены (часто это каталог или сайт, отдающий цены скриптом)"


LITE_URL = "https://lite.duckduckgo.com/lite/"
LITE_PATTERN = re.compile(r"<a[^>]+href=\"(?P<href>[^\"]+)\"[^>]*class=['\"]result-link['\"][^>]*>(?P<title>.*?)</a>", re.S)
LITE_SNIPPET_PATTERN = re.compile(r"class=['\"]result-snippet['\"][^>]*>(?P<snippet>.*?)</td>", re.S)


MIN_INTERVAL = 3.0  # DuckDuckGo ограничивает частые запросы с одного адреса
CACHE_SECONDS = 3600
_pace_lock = threading.Lock()
_last_request = 0.0
_cache: dict[tuple[str, int], tuple[float, list[dict]]] = {}


def _paced_post(client: httpx.Client, url: str, data: dict) -> httpx.Response:
    global _last_request
    with _pace_lock:
        wait = _last_request + MIN_INTERVAL - time.monotonic()
        if wait > 0:
            time.sleep(wait)
        try:
            return client.post(url, data=data)
        finally:
            _last_request = time.monotonic()


def _search(client: httpx.Client, query: str, limit: int) -> list[dict]:
    """Выдача DuckDuckGo. На частые запросы он отвечает 202 — ждём и пробуем lite-версию."""
    cached = _cache.get((query, limit))
    if cached and time.monotonic() - cached[0] < CACHE_SECONDS:
        return [dict(item) for item in cached[1]]
    response = None
    for url, pause in ((SEARCH_URL, 3), (SEARCH_URL, 6), (LITE_URL, 0)):
        response = _paced_post(client, url, {"q": query, "kl": "ru-ru"})
        if response.status_code == 200:
            lite = url == LITE_URL
            results = parse_results(response.text, limit, LITE_PATTERN if lite else RESULT_PATTERN,
                                    LITE_SNIPPET_PATTERN if lite else SNIPPET_PATTERN)
            _cache[(query, limit)] = (time.monotonic(), [dict(item) for item in results])
            return results
        time.sleep(pause)
    raise RuntimeError(f"Поисковик временно ограничил частые запросы (HTTP {response.status_code}). "
                       "Повторите поиск этой позиции через минуту.")


def search_prices(query: str, limit: int = 8, open_pages: int = 5,
                  transport: httpx.BaseTransport | None = None) -> dict:
    query = re.sub(r"\s+", " ", query).strip()
    if not query:
        raise ValueError("Укажите, что искать: наименование позиции")
    limit = max(1, min(int(limit), 20))
    try:
        with httpx.Client(timeout=12, transport=transport, headers={"User-Agent": USER_AGENT},
                          follow_redirects=True) as client:
            results = _search(client, f"{query} цена купить", limit)
            # Открываем страницы параллельно: последовательно это заняло бы до минуты.
            with ThreadPoolExecutor(max_workers=5) as pool:
                list(pool.map(lambda item: _visit(client, item),
                              [item for item in results if not item["prices_rub"]][:max(0, open_pages)]))
    except httpx.HTTPError as error:
        raise RuntimeError(f"Поиск цен недоступен: нет связи с интернетом ({error.__class__.__name__})") from None
    prices = sorted(price for item in results for price in item["prices_rub"])
    summary = None
    if prices:
        summary = {"min": prices[0], "median": prices[len(prices) // 2], "max": prices[-1], "count": len(prices)}
    return {"query": query, "results": results, "summary_rub": summary,
            "note": "Ориентировочные цены из поисковой выдачи (DuckDuckGo). Откройте ссылку и проверьте "
                    "цену, наличие и характеристики; это не коммерческое предложение."}
