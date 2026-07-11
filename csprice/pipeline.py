"""整合流程：決定要抓哪些 (皮膚 × 磨損)，抓 Steam / BUFF，合併成列。"""

import logging

from . import wears
from .skins import AK47_SKINS_FALLBACK, market_hash_name
from .excel import wear_zh

log = logging.getLogger("csprice.pipeline")


def build_universe_from_buff(buff_items):
    """由 BUFF goods 建立 (market_hash_name -> 正規化資料) 對照，
    只保留使用者要的五種磨損。回傳 dict[str, dict]。"""
    from .buff import BuffClient
    universe = {}
    for raw in buff_items:
        parsed = BuffClient.parse_item(raw)
        if parsed["wear_key"] not in wears.WANTED_KEYS:
            continue
        if not parsed["market_hash_name"]:
            continue
        universe[parsed["market_hash_name"]] = parsed
    return universe


def build_universe_from_fallback():
    """無 BUFF 時，用硬編皮膚清單 × 五磨損組出 market_hash_name。"""
    universe = {}
    for skin in AK47_SKINS_FALLBACK:
        for w in wears.WEARS:
            mhn = market_hash_name(skin, w["en"])
            universe[mhn] = {
                "market_hash_name": mhn,
                "skin": skin,
                "wear_key": w["key"],
                "lowest_sell": None,   # BUFF 欄留空
                "highest_buy": None,
                "sell_num": None,
            }
    return universe


def merge_rows(universe, steam_by_mhn, fetched_at):
    """把 BUFF(universe 內含) 與 Steam 資料合併成 Excel 列，依皮膚→磨損排序。"""
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
                             wears.WEAR_ORDER.get(r["wear_key"], 99)))
    return rows
