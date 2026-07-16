"""Steam 社群市集抓取。

提供每個 market_hash_name 的：
  - 最低販賣價 (lowest sell)
  - 最高求購價 (highest buy order)
  - 交易量 (24h volume)

Steam 沒有「查詢過去某時間點掛單簿」的接口，只能取當下即時值。
priceoverview 提供 lowest_price 與 24h volume，但沒有最高求購；
itemordershistogram 才有 highest_buy_order / lowest_sell_order，
但需要 item_nameid（要先從商品頁 HTML 解析並快取）。

Steam 反爬很嚴，預設節流；遇 429 會退避重試。
"""

import re
import time
import logging
from urllib.parse import quote

import requests

log = logging.getLogger("csprice.steam")

# Steam 貨幣代碼：1=USD, 3=EUR, 23=CNY(¥)。與 BUFF(人民幣) 對齊預設用 23。
CURRENCY_CNY = 23
APPID_CS2 = 730

_NAMEID_RE = re.compile(r"Market_LoadOrderSpread\(\s*(\d+)\s*\)")
_PRICE_NUM_RE = re.compile(r"[\d.,]+")


def _parse_price(text):
    """'¥ 1,234.56' / '$12.30' -> 1234.56 / 12.30，失敗回 None。"""
    if text is None:
        return None
    m = _PRICE_NUM_RE.search(str(text))
    if not m:
        return None
    raw = m.group(0)
    # 移除千分位逗號；Steam 這些地區以 . 為小數點。
    raw = raw.replace(",", "")
    try:
        return float(raw)
    except ValueError:
        return None


class SteamClient:
    def __init__(self, currency=CURRENCY_CNY, country="CN", language="schinese",
                 delay=3.0, timeout=20, max_retries=4, session=None,
                 cookie=None, mode="full", cooldown=300):
        self.currency = currency
        self.country = country
        self.language = language
        self.delay = delay          # 每次請求後最少間隔秒數（避免 429）
        self.timeout = timeout
        self.max_retries = max_retries
        self.mode = mode            # "full"=含最高求購, "lite"=只 priceoverview
        self.cooldown = cooldown    # 持續 429 時的長冷卻秒數
        self.s = session or requests.Session()
        self.s.headers.update({
            "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                           "AppleWebKit/537.36 (KHTML, like Gecko) "
                           "Chrome/122.0 Safari/537.36"),
            "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
            "Referer": "https://steamcommunity.com/market/",
            "X-Requested-With": "XMLHttpRequest",
        })
        if cookie:
            # 登入後的 Steam session（steamLoginSecure=...）限流門檻高很多
            self.s.headers["Cookie"] = cookie
        self._nameid_cache = {}

    # -- 低階：帶退避的 GET --------------------------------------------------
    def _get(self, url, params=None, referer=None, allow_cooldown=True):
        headers = {"Referer": referer} if referer else None
        backoff = self.delay
        last_was_429 = False
        for attempt in range(1, self.max_retries + 1):
            try:
                r = self.s.get(url, params=params, headers=headers,
                               timeout=self.timeout)
            except requests.RequestException as e:
                log.warning("Steam 請求例外 (%s/%s): %s", attempt,
                            self.max_retries, e)
                time.sleep(backoff)
                backoff *= 2
                last_was_429 = False
                continue
            if r.status_code == 429:
                log.warning("Steam 429 限流，退避 %.1fs (%s/%s)", backoff * 2,
                            attempt, self.max_retries)
                time.sleep(backoff * 2)
                backoff *= 2
                last_was_429 = True
                continue
            return r
        # 短退避都用完仍被限流 → 長冷卻再重試一輪（Steam 通常需數分鐘解封）
        if last_was_429 and allow_cooldown and self.cooldown > 0:
            log.warning("Steam 持續限流，冷卻 %ss 後再試…（可 Ctrl+C 中止，"
                        "已抓部分會存檔）", self.cooldown)
            time.sleep(self.cooldown)
            return self._get(url, params, referer, allow_cooldown=False)
        return None

    # -- item_nameid：商品頁解析 + 快取 --------------------------------------
    def get_item_nameid(self, market_hash_name):
        if market_hash_name in self._nameid_cache:
            return self._nameid_cache[market_hash_name]
        url = (f"https://steamcommunity.com/market/listings/"
               f"{APPID_CS2}/{quote(market_hash_name)}")
        r = self._get(url)
        nameid = None
        if r is not None and r.status_code == 200:
            m = _NAMEID_RE.search(r.text)
            if m:
                nameid = m.group(1)
        if nameid is None:
            log.info("找不到 item_nameid：%s（可能無此磨損或被限流）",
                     market_hash_name)
        self._nameid_cache[market_hash_name] = nameid
        time.sleep(self.delay)
        return nameid

    # -- priceoverview：24h 交易量 + 最低賣價 ---------------------------------
    def price_overview(self, market_hash_name):
        url = "https://steamcommunity.com/market/priceoverview/"
        params = {
            "appid": APPID_CS2,
            "currency": self.currency,
            "market_hash_name": market_hash_name,
        }
        r = self._get(url, params=params)
        time.sleep(self.delay)
        if r is None or r.status_code != 200:
            return {}
        try:
            data = r.json()
        except ValueError:
            return {}
        if not data.get("success"):
            return {}
        return {
            "lowest_price": _parse_price(data.get("lowest_price")),
            "median_price": _parse_price(data.get("median_price")),
            "volume": _parse_int(data.get("volume")),
        }

    # -- itemordershistogram：最高求購 / 最低賣單 -----------------------------
    def order_histogram(self, market_hash_name, item_nameid=None):
        item_nameid = item_nameid or self.get_item_nameid(market_hash_name)
        if not item_nameid:
            return {}
        url = "https://steamcommunity.com/market/itemordershistogram"
        params = {
            "country": self.country,
            "language": self.language,
            "currency": self.currency,
            "item_nameid": item_nameid,
            "two_factor": 0,
        }
        referer = (f"https://steamcommunity.com/market/listings/"
                   f"{APPID_CS2}/{quote(market_hash_name)}")
        r = self._get(url, params=params, referer=referer)
        time.sleep(self.delay)
        if r is None or r.status_code != 200:
            return {}
        try:
            data = r.json()
        except ValueError:
            return {}
        if not data.get("success"):
            return {}
        return {
            "highest_buy_order": _parse_price(
                _cents(data.get("highest_buy_order"))),
            "lowest_sell_order": _parse_price(
                _cents(data.get("lowest_sell_order"))),
        }

    # -- 批次搜尋：一次抓多款（最抗限流）------------------------------------
    def search_ak(self, max_pages=20, page_size=100):
        """批次抓 AK-47 清單：/market/search/render 一次回最多 100 款，
        只需數次請求。回傳 dict[market_hash_name] -> {lowest_sell, listings}。
        來源僅有最低賣價與在售數量，無最高求購與 24h 交易量。"""
        url = "https://steamcommunity.com/market/search/render/"
        out = {}
        start = 0
        while start < max_pages * page_size:
            params = {
                "norender": 1, "appid": APPID_CS2, "currency": self.currency,
                "count": page_size, "start": start,
                "category_730_Weapon[]": "tag_weapon_ak47",
                "search_descriptions": 0,
            }
            r = self._get(url, params=params)
            time.sleep(self.delay)
            if r is None or r.status_code != 200:
                log.warning("Steam search 失敗（start=%s）", start)
                break
            try:
                data = r.json()
            except ValueError:
                break
            results = data.get("results") or []
            total = data.get("total_count") or 0
            for it in results:
                hn = it.get("hash_name")
                if not hn:
                    continue
                out[hn] = {
                    "lowest_sell": _cents(it.get("sell_price")),
                    "listings": it.get("sell_listings"),
                }
            log.info("Steam search：start=%s 本頁 %s 款，累計 %s / 共 %s",
                     start, len(results), len(out), total)
            # 依實際回傳筆數前進（Steam 常只回 10/頁，不可用 page_size 硬跳）
            start += len(results)
            if not results or start >= total:
                break
        return out

    # -- 對外：整合單一 market_hash_name 的所有欄位 --------------------------
    def fetch(self, market_hash_name):
        """回傳 dict：lowest_sell / highest_buy / volume（缺值為 None）。"""
        overview = self.price_overview(market_hash_name)
        # lite 模式只打 priceoverview（1 次/款）降低限流；代價是無最高求購。
        hist = {} if self.mode == "lite" else self.order_histogram(market_hash_name)
        # 最低賣價優先採 histogram（即時掛單），退回 priceoverview。
        lowest_sell = hist.get("lowest_sell_order")
        if lowest_sell is None:
            lowest_sell = overview.get("lowest_price")
        return {
            "lowest_sell": lowest_sell,
            "highest_buy": hist.get("highest_buy_order"),
            "volume": overview.get("volume"),
            "currency": self.currency,
        }


def _cents(v):
    """itemordershistogram 的價格為「分」的整數字串，轉成元。"""
    if v in (None, ""):
        return None
    try:
        return int(v) / 100.0
    except (ValueError, TypeError):
        return v


def _parse_int(v):
    if v is None:
        return None
    m = re.search(r"[\d,]+", str(v))
    if not m:
        return None
    try:
        return int(m.group(0).replace(",", ""))
    except ValueError:
        return None
