#!/usr/bin/env python3
"""CS2 AK-47 皮膚 Steam vs BUFF 價格比較 → Excel（單檔自包含版）。

這支是把整個 csprice 套件合併成的「單一檔案」版本，方便你下載一個檔就能跑，
不需要 git clone 整個專案。功能與 ak_price_compare.py 完全相同。

只依賴兩個第三方套件：
    pip install requests openpyxl

爬取每款 AK-47 皮膚在五種磨損（全新出場 / 略有磨損 / 戰場實測 /
戰痕累累 / 重度磨損）下，Steam 與 BUFF 的：最低販賣價、最高求購價、交易量，
並輸出成可比較的 Excel。

⚠ 時間點：Steam / BUFF 皆無「查詢過去某時間掛單」的接口，本工具抓的是
「執行當下」的即時值，Excel 會記錄實際抓取時間。

用法：
    # 完整（含 BUFF，皮膚清單由 BUFF 動態發現，最完整）
    python ak_price_compare_standalone.py --buff-cookie "session=xxxx" -o ak.xlsx

    # 只抓 Steam（皮膚用內建後備清單，不需 Cookie）
    python ak_price_compare_standalone.py --steam-only -o ak.xlsx

    # 自我測試：用假資料驗證 Excel 產出流程（不連網）
    python ak_price_compare_standalone.py --self-test -o demo.xlsx
"""

import os
import re
import sys
import time
import argparse
import logging
from datetime import datetime
from urllib.parse import quote

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

log = logging.getLogger("csprice")

# ======================================================================
# 磨損定義
# ======================================================================
# 依使用者需求的五種磨損，順序即 Excel 輸出順序。
WEARS = [
    {"key": "FN", "zh_tw": "全新出場", "zh_cn": "崭新出厂", "en": "Factory New"},
    {"key": "MW", "zh_tw": "略有磨損", "zh_cn": "略有磨损", "en": "Minimal Wear"},
    {"key": "FT", "zh_tw": "戰場實測", "zh_cn": "久经沙场", "en": "Field-Tested"},
    {"key": "WW", "zh_tw": "戰痕累累", "zh_cn": "破损不堪", "en": "Well-Worn"},
    {"key": "BS", "zh_tw": "重度磨損", "zh_cn": "战痕累累", "en": "Battle-Scarred"},
]
_EN_TO_KEY = {w["en"].lower(): w["key"] for w in WEARS}
_ZH_TO_KEY = {}
for _w in WEARS:
    _ZH_TO_KEY[_w["zh_tw"]] = _w["key"]
    _ZH_TO_KEY[_w["zh_cn"]] = _w["key"]
WEAR_ORDER = {w["key"]: i for i, w in enumerate(WEARS)}
WEAR_BY_KEY = {w["key"]: w for w in WEARS}
WANTED_KEYS = [w["key"] for w in WEARS]


def key_from_english(exterior_en):
    if not exterior_en:
        return None
    return _EN_TO_KEY.get(exterior_en.strip().lower())


def key_from_market_hash_name(mhn):
    if not mhn or "(" not in mhn:
        return None
    exterior = mhn.rsplit("(", 1)[-1].rstrip(")").strip()
    return key_from_english(exterior)


def key_from_chinese(exterior_zh):
    if not exterior_zh:
        return None
    return _ZH_TO_KEY.get(exterior_zh.strip())


def wear_zh(wear_key):
    w = WEAR_BY_KEY.get(wear_key)
    return w["zh_tw"] if w else wear_key


# ======================================================================
# AK-47 後備皮膚清單（無 BUFF 時用）
# ======================================================================
AK47_SKINS_FALLBACK = [
    "Aquamarine Revenge", "Asiimov", "Baroque Purple", "B the Monster",
    "Black Laminate", "Blue Laminate", "Bloodsport", "Cartel", "Case Hardened",
    "Elite Build", "Emerald Pinstripe", "The Empress", "Fire Serpent",
    "First Class", "Frontside Misty", "Fuel Injector", "Gold Arabesque",
    "Green Laminate", "Head Shot", "Hydroponic", "Ice Coaled", "Inheritance",
    "Jaguar", "Jet Set", "Legion of Anubis", "Leet Museo", "Nautilus",
    "Neon Revolution", "Neon Rider", "Nightwish", "Orbit Mk01", "Panthera onca",
    "Phantom Disruptor", "Point Disarray", "Predator", "Rat Rod",
    "Red Laminate", "Redline", "Safari Mesh", "Safety Net", "Searing Rage",
    "Slate", "Steel Delta", "Uncharted", "Vulcan", "Wasteland Rebel",
    "Wild Lotus", "X-Ray",
]


def mhn_of(skin, wear_en):
    return f"AK-47 | {skin} ({wear_en})"


# ======================================================================
# Steam 抓取
# ======================================================================
CURRENCY_CNY = 23
APPID_CS2 = 730
_NAMEID_RE = re.compile(r"Market_LoadOrderSpread\(\s*(\d+)\s*\)")
_PRICE_NUM_RE = re.compile(r"[\d.,]+")


def _parse_price(text):
    if text is None:
        return None
    m = _PRICE_NUM_RE.search(str(text))
    if not m:
        return None
    raw = m.group(0).replace(",", "")
    try:
        return float(raw)
    except ValueError:
        return None


def _cents(v):
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


class SteamClient:
    def __init__(self, currency=CURRENCY_CNY, country="CN", language="schinese",
                 delay=3.0, timeout=20, max_retries=4, session=None):
        self.currency = currency
        self.country = country
        self.language = language
        self.delay = delay
        self.timeout = timeout
        self.max_retries = max_retries
        self.s = session or requests.Session()
        self.s.headers.update({
            "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                           "AppleWebKit/537.36 (KHTML, like Gecko) "
                           "Chrome/122.0 Safari/537.36"),
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        })
        self._nameid_cache = {}

    def _get(self, url, params=None, referer=None):
        headers = {"Referer": referer} if referer else None
        backoff = self.delay
        for attempt in range(1, self.max_retries + 1):
            try:
                r = self.s.get(url, params=params, headers=headers,
                               timeout=self.timeout)
            except requests.RequestException as e:
                log.warning("Steam 請求例外 (%s/%s): %s", attempt,
                            self.max_retries, e)
                time.sleep(backoff)
                backoff *= 2
                continue
            if r.status_code == 429:
                log.warning("Steam 429 限流，退避 %.1fs (%s/%s)", backoff * 2,
                            attempt, self.max_retries)
                time.sleep(backoff * 2)
                backoff *= 2
                continue
            return r
        return None

    def get_item_nameid(self, mhn):
        if mhn in self._nameid_cache:
            return self._nameid_cache[mhn]
        url = (f"https://steamcommunity.com/market/listings/"
               f"{APPID_CS2}/{quote(mhn)}")
        r = self._get(url)
        nameid = None
        if r is not None and r.status_code == 200:
            m = _NAMEID_RE.search(r.text)
            if m:
                nameid = m.group(1)
        if nameid is None:
            log.info("找不到 item_nameid：%s（可能無此磨損或被限流）", mhn)
        self._nameid_cache[mhn] = nameid
        time.sleep(self.delay)
        return nameid

    def price_overview(self, mhn):
        url = "https://steamcommunity.com/market/priceoverview/"
        params = {"appid": APPID_CS2, "currency": self.currency,
                  "market_hash_name": mhn}
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

    def order_histogram(self, mhn, item_nameid=None):
        item_nameid = item_nameid or self.get_item_nameid(mhn)
        if not item_nameid:
            return {}
        url = "https://steamcommunity.com/market/itemordershistogram"
        params = {"country": self.country, "language": self.language,
                  "currency": self.currency, "item_nameid": item_nameid,
                  "two_factor": 0}
        referer = (f"https://steamcommunity.com/market/listings/"
                   f"{APPID_CS2}/{quote(mhn)}")
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
            "highest_buy_order": _parse_price(_cents(data.get("highest_buy_order"))),
            "lowest_sell_order": _parse_price(_cents(data.get("lowest_sell_order"))),
        }

    def fetch(self, mhn):
        overview = self.price_overview(mhn)
        hist = self.order_histogram(mhn)
        lowest_sell = hist.get("lowest_sell_order")
        if lowest_sell is None:
            lowest_sell = overview.get("lowest_price")
        return {
            "lowest_sell": lowest_sell,
            "highest_buy": hist.get("highest_buy_order"),
            "volume": overview.get("volume"),
            "currency": self.currency,
        }


# ======================================================================
# BUFF 抓取
# ======================================================================
GAME = "csgo"
CATEGORY_AK47 = "weapon_ak47"


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


def _skin_from_mhn(mhn):
    if not mhn:
        return None
    body = mhn
    if "|" in body:
        body = body.split("|", 1)[1]
    if "(" in body:
        body = body.rsplit("(", 1)[0]
    return body.strip()


class BuffClient:
    def __init__(self, cookie, delay=1.5, timeout=20, max_retries=4, session=None):
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
            "Cookie": cookie if "=" in cookie else f"session={cookie}",
        })

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

    def _goods_page(self, page_num):
        url = "https://buff.163.com/api/market/goods"
        params = {"game": GAME, "page_num": page_num,
                  "category": CATEGORY_AK47, "sort_by": "default"}
        r = self._get(url, params)
        time.sleep(self.delay)
        if r is None:
            return None
        try:
            data = r.json()
        except ValueError:
            log.error("BUFF 回傳非 JSON（page %s），可能 Cookie 失效或被擋", page_num)
            return None
        if data.get("code") != "OK":
            log.error("BUFF 回應 code=%s msg=%s（page %s）",
                      data.get("code"), data.get("msg"), page_num)
            return None
        return data.get("data", {})

    def discover_ak_goods(self, max_pages=50):
        items = []
        page = 1
        total_page = None
        while page <= max_pages:
            data = self._goods_page(page)
            if not data:
                break
            page_items = data.get("items", [])
            items.extend(page_items)
            total_page = data.get("total_page", total_page)
            log.info("BUFF AK-47 第 %s/%s 頁，本頁 %s 筆，累計 %s",
                     page, total_page, len(page_items), len(items))
            if total_page is not None and page >= total_page:
                break
            if not page_items:
                break
            page += 1
        return items

    @staticmethod
    def parse_item(item):
        mhn = item.get("market_hash_name") or ""
        wear_key = key_from_market_hash_name(mhn)
        if wear_key is None:
            tags = (((item.get("goods_info") or {}).get("info") or {})
                    .get("tags") or {})
            ext = (tags.get("exterior") or {}).get("localized_name")
            wear_key = key_from_chinese(ext)
        return {
            "goods_id": item.get("id"),
            "name_cn": item.get("name"),
            "market_hash_name": mhn,
            "skin": _skin_from_mhn(mhn),
            "wear_key": wear_key,
            "lowest_sell": _to_float(item.get("sell_min_price")),
            "highest_buy": _to_float(item.get("buy_max_price")),
            "sell_num": _to_int(item.get("sell_num")),
            "buy_num": _to_int(item.get("buy_num")),
        }


# ======================================================================
# 合併流程
# ======================================================================
def build_universe_from_buff(buff_items):
    universe = {}
    for raw in buff_items:
        parsed = BuffClient.parse_item(raw)
        if parsed["wear_key"] not in WANTED_KEYS:
            continue
        if not parsed["market_hash_name"]:
            continue
        universe[parsed["market_hash_name"]] = parsed
    return universe


def build_universe_from_fallback():
    universe = {}
    for skin in AK47_SKINS_FALLBACK:
        for w in WEARS:
            mhn = mhn_of(skin, w["en"])
            universe[mhn] = {
                "market_hash_name": mhn, "skin": skin, "wear_key": w["key"],
                "lowest_sell": None, "highest_buy": None, "sell_num": None,
            }
    return universe


def merge_rows(universe, steam_by_mhn, fetched_at):
    rows = []
    for mhn, buff in universe.items():
        steam = steam_by_mhn.get(mhn, {})
        s_sell = steam.get("lowest_sell")
        b_sell = buff.get("lowest_sell")
        sell_diff = (s_sell - b_sell) if (s_sell is not None
                                          and b_sell is not None) else None
        sell_ratio = (s_sell / b_sell) if (s_sell is not None
                                           and b_sell not in (None, 0)) else None
        rows.append({
            "skin": buff.get("skin"),
            "wear_key": buff.get("wear_key"),
            "wear_zh": wear_zh(buff.get("wear_key")),
            "market_hash_name": mhn,
            "steam_lowest_sell": s_sell,
            "steam_highest_buy": steam.get("highest_buy"),
            "steam_volume": steam.get("volume"),
            "buff_lowest_sell": b_sell,
            "buff_highest_buy": buff.get("highest_buy"),
            "buff_sell_num": buff.get("sell_num"),
            "sell_diff": round(sell_diff, 2) if sell_diff is not None else None,
            "sell_ratio": round(sell_ratio, 4) if sell_ratio is not None else None,
            "fetched_at": fetched_at,
        })
    rows.sort(key=lambda r: (r["skin"] or "",
                             WEAR_ORDER.get(r["wear_key"], 99)))
    return rows


# ======================================================================
# Excel 輸出
# ======================================================================
COLUMNS = [
    ("皮膚", "skin", 22, None),
    ("磨損", "wear_zh", 10, None),
    ("market_hash_name", "market_hash_name", 34, None),
    ("Steam最低賣價", "steam_lowest_sell", 14, "#,##0.00"),
    ("Steam最高求購", "steam_highest_buy", 14, "#,##0.00"),
    ("Steam交易量(24h)", "steam_volume", 14, "#,##0"),
    ("BUFF最低賣價", "buff_lowest_sell", 14, "#,##0.00"),
    ("BUFF最高求購", "buff_highest_buy", 14, "#,##0.00"),
    ("BUFF在售量", "buff_sell_num", 12, "#,##0"),
    ("賣價差(Steam-BUFF)", "sell_diff", 16, "#,##0.00"),
    ("賣價比(Steam/BUFF)", "sell_ratio", 16, "0.00"),
    ("抓取時間", "fetched_at", 20, None),
]
_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_THIN = Side(style="thin", color="D9D9D9")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_ALT_FILL = PatternFill("solid", fgColor="F2F6FB")


def write_workbook(rows, path, currency_label="¥ (人民幣)"):
    wb = Workbook()
    ws = wb.active
    ws.title = "AK-47 價格比較"
    ws.cell(row=1, column=1,
            value=f"CS2 AK-47 皮膚 Steam vs BUFF 價格比較　幣別：{currency_label}")
    ws.cell(row=1, column=1).font = Font(bold=True, size=12)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLUMNS))
    header_row = 2
    for c, (title, _key, width, _fmt) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=header_row, column=c, value=title)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = _BORDER
        ws.column_dimensions[get_column_letter(c)].width = width
    for i, row in enumerate(rows):
        r = header_row + 1 + i
        for c, (_title, key, _width, fmt) in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=r, column=c, value=row.get(key))
            cell.border = _BORDER
            if fmt:
                cell.number_format = fmt
            if i % 2 == 1:
                cell.fill = _ALT_FILL
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(COLUMNS))}{header_row}"
    wb.save(path)
    return path


# ======================================================================
# CLI
# ======================================================================
def parse_args(argv=None):
    p = argparse.ArgumentParser(description="CS2 AK-47 Steam/BUFF 價格比較")
    p.add_argument("-o", "--output", default="ak_price_compare.xlsx")
    p.add_argument("--buff-cookie", default=os.environ.get("BUFF_COOKIE"))
    p.add_argument("--steam-only", action="store_true")
    p.add_argument("--buff-only", action="store_true")
    p.add_argument("--currency", type=int, default=23)
    p.add_argument("--steam-delay", type=float, default=3.0)
    p.add_argument("--limit", type=int, default=0)
    p.add_argument("--self-test", action="store_true")
    p.add_argument("-v", "--verbose", action="store_true")
    return p.parse_args(argv)


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def run_self_test(output):
    fake_buff_items = [
        {"id": 1, "name": "AK-47 | 红线 (久经沙场)",
         "market_hash_name": "AK-47 | Redline (Field-Tested)",
         "sell_min_price": "82.50", "buy_max_price": "78.00",
         "sell_num": 1200, "buy_num": 340},
        {"id": 2, "name": "AK-47 | 红线 (略有磨损)",
         "market_hash_name": "AK-47 | Redline (Minimal Wear)",
         "sell_min_price": "150.00", "buy_max_price": "142.00",
         "sell_num": 300, "buy_num": 80},
        {"id": 3, "name": "AK-47 | 二西莫夫 (崭新出厂)",
         "market_hash_name": "AK-47 | Asiimov (Factory New)",
         "sell_min_price": "420.00", "buy_max_price": "405.00",
         "sell_num": 90, "buy_num": 45},
        {"id": 4, "name": "AK-47 | 表面淬火 (战痕累累)",
         "market_hash_name": "AK-47 | Case Hardened (Battle-Scarred)",
         "sell_min_price": "260.00", "buy_max_price": "240.00",
         "sell_num": 55, "buy_num": 12},
        {"id": 5, "name": "AK-47 | 皇后 (破损不堪)",
         "market_hash_name": "AK-47 | The Empress (Well-Worn)",
         "sell_min_price": "180.00", "buy_max_price": "170.00",
         "sell_num": 20, "buy_num": 6},
    ]
    universe = build_universe_from_buff(fake_buff_items)
    fake_steam = {
        "AK-47 | Redline (Field-Tested)":
            {"lowest_sell": 90.10, "highest_buy": 85.00, "volume": 5400},
        "AK-47 | Redline (Minimal Wear)":
            {"lowest_sell": 168.00, "highest_buy": 160.00, "volume": 900},
        "AK-47 | Asiimov (Factory New)":
            {"lowest_sell": 455.00, "highest_buy": 430.00, "volume": 210},
        "AK-47 | The Empress (Well-Worn)":
            {"lowest_sell": 205.00, "highest_buy": 190.00, "volume": 60},
    }
    rows = merge_rows(universe, fake_steam, now_str())
    write_workbook(rows, output)
    print(f"[self-test] 產出 {len(rows)} 列 -> {output}")
    for r in rows:
        print(f"  {r['skin']:<14} {r['wear_zh']:<6} "
              f"Steam賣={r['steam_lowest_sell']} BUFF賣={r['buff_lowest_sell']} "
              f"差={r['sell_diff']}")
    return 0


def main(argv=None):
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.INFO if args.verbose else logging.WARNING,
        format="%(levelname)s %(name)s: %(message)s")

    if args.self_test:
        return run_self_test(args.output)

    use_buff = not args.steam_only
    use_steam = not args.buff_only
    fetched_at = now_str()

    if use_buff:
        if not args.buff_cookie:
            print("錯誤：未提供 BUFF Cookie。請用 --buff-cookie 或設 BUFF_COOKIE，"
                  "或改用 --steam-only。", file=sys.stderr)
            return 2
        print("→ 從 BUFF 發現所有 AK-47 皮膚 …")
        buff = BuffClient(args.buff_cookie)
        items = buff.discover_ak_goods()
        universe = build_universe_from_buff(items)
        print(f"  BUFF 取得 {len(universe)} 筆 (皮膚×指定磨損)")
    else:
        universe = build_universe_from_fallback()
        print(f"→ 使用內建後備皮膚清單，共 {len(universe)} 筆 (皮膚×磨損)")

    keys = list(universe.keys())
    if args.limit:
        keys = keys[:args.limit]

    steam_by_mhn = {}
    if use_steam:
        steam = SteamClient(currency=args.currency, delay=args.steam_delay)
        print(f"→ 抓 Steam（{len(keys)} 筆，間隔 {args.steam_delay}s，"
              "數量多時很慢請耐心）…")
        for i, mhn in enumerate(keys, 1):
            steam_by_mhn[mhn] = steam.fetch(mhn)
            if i % 10 == 0 or i == len(keys):
                print(f"  Steam 進度 {i}/{len(keys)}")

    sub_universe = {k: universe[k] for k in keys}
    rows = merge_rows(sub_universe, steam_by_mhn, fetched_at)
    cur = {1: "$ (美元)", 3: "€ (歐元)", 23: "¥ (人民幣)"}.get(
        args.currency, str(args.currency))
    write_workbook(rows, args.output, currency_label=cur)
    print(f"✔ 完成：{len(rows)} 列 -> {args.output}（抓取時間 {fetched_at}）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
