#!/usr/bin/env python3
"""CS2 AK-47 皮膚 Steam vs BUFF 價格比較 → Excel。

爬取每款 AK-47 皮膚在五種磨損（全新出場 / 略有磨損 / 戰場實測 /
戰痕累累 / 重度磨損）下，Steam 與 BUFF 的：
    最低販賣價、最高求購價、交易量
並輸出成可比較的 Excel。

⚠ 時間點：Steam / BUFF 皆無「查詢過去某時間掛單」的接口，本工具抓的是
「執行當下」的即時值，Excel 會記錄實際抓取時間。

用法：
    # 完整（含 BUFF，皮膚清單由 BUFF 動態發現，最完整）
    python ak_price_compare.py --buff-cookie "session=xxxx" -o ak.xlsx

    # 只抓 Steam（皮膚用內建後備清單）
    python ak_price_compare.py --steam-only -o ak.xlsx

    # 自我測試：用假資料驗證 Excel 產出流程（不連網）
    python ak_price_compare.py --self-test -o demo.xlsx
"""

import os
import sys
import argparse
import logging
from datetime import datetime

from csprice import wears
from csprice.excel import write_workbook
from csprice import pipeline


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="CS2 AK-47 Steam/BUFF 價格比較")
    p.add_argument("-o", "--output", default="ak_price_compare.xlsx",
                   help="輸出 Excel 路徑")
    p.add_argument("--buff-cookie", default=os.environ.get("BUFF_COOKIE"),
                   help="BUFF 登入 Cookie（或設環境變數 BUFF_COOKIE）")
    p.add_argument("--steam-only", action="store_true",
                   help="不抓 BUFF，皮膚用內建後備清單")
    p.add_argument("--buff-only", action="store_true",
                   help="只抓 BUFF，不連 Steam")
    p.add_argument("--currency", type=int, default=30,
                   help="Steam 貨幣代碼：23=CNY、30=TWD(新台幣,預設)、1=USD、3=EUR")
    p.add_argument("--twd-rate", type=float, default=4.77,
                   help="BUFF 人民幣換算表格幣別的匯率（預設 4.77，即 1 RMB=4.77 TWD）")
    p.add_argument("--steam-delay", type=float, default=3.0,
                   help="Steam 每請求間隔秒數，被限流就調大（如 8~15）")
    p.add_argument("--steam-mode", choices=["search", "lite", "full"],
                   default="search",
                   help="search=批次抓最低賣價+在售量(數次請求,最抗限流,預設); "
                        "lite=逐款最低賣價+24h量(1請求/款); "
                        "full=逐款再加最高求購(3請求/款,最易被限流)")
    p.add_argument("--steam-cookie", default=os.environ.get("STEAM_COOKIE"),
                   help="Steam 登入 Cookie(steamLoginSecure=...)，限流門檻高很多")
    p.add_argument("--steam-cooldown", type=int, default=300,
                   help="Steam 持續被限流時的長冷卻秒數，0=關閉")
    p.add_argument("--limit", type=int, default=0,
                   help="只處理前 N 個 (皮膚×磨損) 用於測試，0=全部")
    p.add_argument("--self-test", action="store_true",
                   help="用假資料驗證 Excel 產出（不連網）")
    p.add_argument("-v", "--verbose", action="store_true")
    return p.parse_args(argv)


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def run_self_test(output):
    """不連網，塞入假資料跑完 merge + Excel，驗證整條產出流程。"""
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
        {"id": 4, "name": "AK-47 | 表面淬火 (战痕累累)",  # BS
         "market_hash_name": "AK-47 | Case Hardened (Battle-Scarred)",
         "sell_min_price": "260.00", "buy_max_price": "240.00",
         "sell_num": 55, "buy_num": 12},
        {"id": 5, "name": "AK-47 | 皇后 (破损不堪)",       # WW
         "market_hash_name": "AK-47 | The Empress (Well-Worn)",
         "sell_min_price": "180.00", "buy_max_price": "170.00",
         "sell_num": 20, "buy_num": 6},
    ]
    universe = pipeline.build_universe_from_buff(fake_buff_items)
    fake_steam = {
        "AK-47 | Redline (Field-Tested)":
            {"lowest_sell": 90.10, "highest_buy": 85.00, "volume": 5400},
        "AK-47 | Redline (Minimal Wear)":
            {"lowest_sell": 168.00, "highest_buy": 160.00, "volume": 900},
        "AK-47 | Asiimov (Factory New)":
            {"lowest_sell": 455.00, "highest_buy": 430.00, "volume": 210},
        # Case Hardened 故意缺 Steam 資料 -> 測試留空與差價 None
        "AK-47 | The Empress (Well-Worn)":
            {"lowest_sell": 205.00, "highest_buy": 190.00, "volume": 60},
    }
    rows = pipeline.merge_rows(universe, fake_steam, now_str())
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

    # 1) 建立要處理的 (皮膚 × 磨損) 宇宙
    if use_buff:
        if not args.buff_cookie:
            print("錯誤：未提供 BUFF Cookie。請用 --buff-cookie 或設 BUFF_COOKIE，"
                  "或改用 --steam-only。", file=sys.stderr)
            return 2
        from csprice.buff import BuffClient
        print("→ 從 BUFF 發現所有 AK-47 皮膚 …")
        buff = BuffClient(args.buff_cookie)
        items = buff.discover_ak_goods()
        universe = pipeline.build_universe_from_buff(items)
        print(f"  BUFF 取得 {len(universe)} 筆 (皮膚×指定磨損)")
    else:
        universe = pipeline.build_universe_from_fallback()
        print(f"→ 使用內建後備皮膚清單，共 {len(universe)} 筆 (皮膚×磨損)")

    keys = list(universe.keys())
    if args.limit:
        keys = keys[:args.limit]

    cur = {1: "$ (美元)", 3: "€ (歐元)", 23: "¥ (人民幣)",
           30: "NT$ (新台幣)"}.get(args.currency, str(args.currency))
    if args.twd_rate != 1.0:
        cur += f"（BUFF 以 1 RMB={args.twd_rate} 換算）"
    sub_universe = {k: universe[k] for k in keys}

    def save(steam_map):
        rows = pipeline.merge_rows(sub_universe, steam_map, fetched_at,
                                   buff_rate=args.twd_rate)
        write_workbook(rows, args.output, currency_label=cur)
        return len(rows)

    # 2) 抓 Steam（邊抓邊存）
    steam_by_mhn = {}
    if use_steam:
        from csprice.steam import SteamClient
        steam = SteamClient(currency=args.currency, delay=args.steam_delay,
                            cookie=args.steam_cookie, mode=args.steam_mode,
                            cooldown=args.steam_cooldown)
        note = "" if args.steam_cookie else "（未帶登入 Cookie）"
        print(f"→ 抓 Steam（mode={args.steam_mode}，間隔 {args.steam_delay}s{note}）…")
        if args.steam_mode == "search":
            found = steam.search_ak()
            steam_by_mhn = {mhn: {"lowest_sell": v.get("lowest_sell"),
                                  "highest_buy": None, "volume": None,
                                  "listings": v.get("listings")}
                            for mhn, v in found.items()}
            hit = sum(1 for k in keys if k in steam_by_mhn)
            print(f"  Steam search 完成：對到 {hit}/{len(keys)} 款")
        else:
            try:
                for i, mhn in enumerate(keys, 1):
                    steam_by_mhn[mhn] = steam.fetch(mhn)
                    if i % 10 == 0 or i == len(keys):
                        save(steam_by_mhn)
                        print(f"  Steam 進度 {i}/{len(keys)}（已存檔 {args.output}）")
            except KeyboardInterrupt:
                n = save(steam_by_mhn)
                print(f"\n⚠ 已中止，將已抓的 {n} 列存到 {args.output}")
                return 0

    # 3) 合併 + 輸出
    n = save(steam_by_mhn)
    print(f"✔ 完成：{n} 列 -> {args.output}（抓取時間 {fetched_at}）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
