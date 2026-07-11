"""AK-47 皮膚清單（後備用）。

主要取得「所有 AK-47 皮膚」的方式是從 BUFF 依分類 weapon_ak47 動態抓取
（見 buff.py 的 discover_ak_goods），那份清單最權威、且會自動涵蓋新皮膚。

這裡的硬編清單只在「無法連 BUFF / 未提供 Cookie」時作為 Steam-only 後備。
內容為 market_hash_name 中 "AK-47 | " 之後的部分，未必涵蓋全部最新皮膚。
"""

# "AK-47 | XXX" 的 XXX 部分
AK47_SKINS_FALLBACK = [
    "Aquamarine Revenge",
    "Asiimov",
    "Baroque Purple",
    "B the Monster",
    "Black Laminate",
    "Blue Laminate",
    "Bloodsport",
    "Cartel",
    "Case Hardened",
    "Elite Build",
    "Emerald Pinstripe",
    "Empress",  # 完整名稱為 "The Empress"，見下方特例
    "The Empress",
    "Fire Serpent",
    "First Class",
    "Frontside Misty",
    "Fuel Injector",
    "Gold Arabesque",
    "Green Laminate",
    "Head Shot",
    "Hydroponic",
    "Ice Coaled",
    "Inheritance",
    "Jaguar",
    "Jet Set",
    "Legion of Anubis",
    "Leet Museo",
    "Nautilus",
    "Neon Revolution",
    "Neon Rider",
    "Nightwish",
    "Orbit Mk01",
    "Panthera onca",
    "Phantom Disruptor",
    "Point Disarray",
    "Predator",
    "Rat Rod",
    "Red Laminate",
    "Redline",
    "Safari Mesh",
    "Safety Net",
    "Searing Rage",
    "Slate",
    "Steel Delta",
    "Uncharted",
    "Vulcan",
    "Wasteland Rebel",
    "Wild Lotus",
    "X-Ray",
]


def base_market_hash_name(skin: str) -> str:
    """由皮膚名組出 'AK-47 | XXX' 前綴。"""
    return f"AK-47 | {skin}"


def market_hash_name(skin: str, wear_en: str) -> str:
    """組出完整 market_hash_name，例：'AK-47 | Redline (Field-Tested)'。"""
    return f"AK-47 | {skin} ({wear_en})"
