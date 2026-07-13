"""BUFF (buff.163.com) 抓取。

BUFF 上「一個 goods = 一個皮膚 + 一種磨損」，因此以分類 weapon_ak47 分頁抓取
即可拿到「所有 AK-47 皮膚 × 各磨損」的完整清單，每筆同時附：
  - sell_min_price 最低販賣價
  - buy_max_price  最高求購價
  - sell_num       在售數量（作為 BUFF 交易量/流通量的近似）
  - market_hash_name 用來與 Steam 對齊

BUFF 價格接口需登入 Cookie（session=...）。伺服器在中國，部分網路環境無法直連。
"""

import time
import logging

import requests

from . import wears

log = logging.getLogger("csprice.buff")

GAME = "csgo"
CATEGORY_AK47 = "weapon_ak47"


class BuffClient:
    def __init__(self, cookie, delay=1.5, timeout=20, max_retries=4,
                 session=None):
        if not cookie:
            raise ValueError("BUFF 需要登入 Cookie（session=...）")
        self.delay = delay
        self.timeout = timeout
        self.max_retries = max_retries
        self.s = session or requests.Session()
        self.s.headers.update({
            "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                           "AppleWebKit/537.36 (KHTML, like Gecko) "
                           "Chrome/122.0 Safari/537.36"),
            "Referer": "https://buff.163.com/market/csgo",
            "Accept": "application/json, text/plain, */*",
            "X-Requested-With": "XMLHttpRequest",
            "Cookie": cookie if "=" in cookie else f"session={cookie}",
        })

    # 依序試出第一個回 OK 的參數組；不同時期 BUFF 對 category/sort_by 接受度不一，
    # search 最穩。
    STRATEGIES = [
        {"category": CATEGORY_AK47},
        {"search": "AK-47"},
        {"search": "AK-47", "category": CATEGORY_AK47},
    ]

    def _get(self, url, params):
        backoff = self.delay
        for attempt in range(1, self.max_retries + 1):
            try:
                r = self.s.get(url, params=params, timeout=self.timeout)
            except requests.RequestException as e:
                log.warning("BUFF 請求例外 (%s/%s): %s", attempt,
                            self.max_retries, e)
                time.sleep(backoff)
                backoff *= 2
                continue
            if r.status_code == 429 or r.status_code >= 500:
                log.warning("BUFF %s，退避 %.1fs", r.status_code, backoff * 2)
                time.sleep(backoff * 2)
                backoff *= 2
                continue
            return r
        return None

    def _goods_page(self, page_num, extra):
        url = "https://buff.163.com/api/market/goods"
        params = {"game": GAME, "page_num": page_num, "page_size": 80}
        params.update(extra)
        r = self._get(url, params)
        time.sleep(self.delay)
        if r is None:
            return None, "no-response"
        try:
            data = r.json()
        except ValueError:
            snippet = (r.text or "")[:200].replace("\n", " ")
            log.error("BUFF 回傳非 JSON（HTTP %s），可能 Cookie 失效或被擋。開頭：%s",
                      r.status_code, snippet)
            return None, "not-json"
        if data.get("code") != "OK":
            log.error("BUFF code=%s msg=%s（page %s, params=%s）",
                      data.get("code"), data.get("msg"), page_num, extra)
            return None, data.get("code")
        return data.get("data", {}), "OK"

    def _pick_strategy(self):
        """用第 1 頁探測，回傳 (可用參數組, 第1頁資料)；都失敗回 (None, None)。"""
        for extra in self.STRATEGIES:
            data, status = self._goods_page(1, extra)
            if status == "OK" and data is not None:
                log.info("BUFF 採用參數：%s", extra)
                return extra, data
            if status in ("not-json", "no-response"):
                break  # Cookie/網路問題，換參數也沒用
        return None, None

    def discover_ak_goods(self, max_pages=50):
        """抓所有 AK-47 goods，回傳原始 item dict 清單。"""
        extra, first = self._pick_strategy()
        if extra is None:
            log.error("BUFF 所有參數組都失敗——多半是 Cookie 失效或被風控，"
                      "請重新登入 buff.163.com 取新的 session。")
            return []
        items = list(first.get("items", []))
        total_page = first.get("total_page")
        log.info("BUFF AK-47 第 1/%s 頁，本頁 %s 筆", total_page, len(items))
        page = 2
        while page <= max_pages:
            if total_page is not None and page > total_page:
                break
            data, status = self._goods_page(page, extra)
            if status != "OK" or data is None:
                break
            page_items = data.get("items", [])
            if not page_items:
                break
            items.extend(page_items)
            total_page = data.get("total_page", total_page)
            log.info("BUFF AK-47 第 %s/%s 頁，本頁 %s 筆，累計 %s",
                     page, total_page, len(page_items), len(items))
            page += 1
        return items

    @staticmethod
    def parse_item(item):
        """把 BUFF goods item 正規化成內部欄位。

        回傳 dict 含：market_hash_name / skin / wear_key / buff 三價量。
        無法判定磨損者 wear_key 為 None（呼叫端會略過）。
        """
        mhn = item.get("market_hash_name") or ""
        wear_key = wears.key_from_market_hash_name(mhn)
        if wear_key is None:
            # 後備：從 goods_info 標籤的 exterior 中文判定
            tags = (((item.get("goods_info") or {}).get("info") or {})
                    .get("tags") or {})
            ext = (tags.get("exterior") or {}).get("localized_name")
            wear_key = wears.key_from_chinese(ext)
        skin = _skin_from_mhn(mhn)
        return {
            "goods_id": item.get("id"),
            "name_cn": item.get("name"),
            "market_hash_name": mhn,
            "skin": skin,
            "wear_key": wear_key,
            "lowest_sell": _to_float(item.get("sell_min_price")),
            "highest_buy": _to_float(item.get("buy_max_price")),
            "sell_num": _to_int(item.get("sell_num")),
            "buy_num": _to_int(item.get("buy_num")),
        }


def _skin_from_mhn(mhn):
    """'AK-47 | Redline (Field-Tested)' -> 'Redline'。"""
    if not mhn:
        return None
    body = mhn
    if "|" in body:
        body = body.split("|", 1)[1]
    if "(" in body:
        body = body.rsplit("(", 1)[0]
    return body.strip()


def _to_float(v):
    if v in (None, ""):
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _to_int(v):
    if v in (None, ""):
        return None
    try:
        return int(v)
    except (ValueError, TypeError):
        return None
