"""磨損等級定義，含繁中 / 簡中 / 英文對照。"""

# 依使用者需求的五種磨損，順序即 Excel 輸出順序。
# key: 內部代碼；zh_tw: 繁中；zh_cn: 簡中(BUFF 用)；en: Steam market_hash_name 用
WEARS = [
    {"key": "FN", "zh_tw": "全新出場", "zh_cn": "崭新出厂", "en": "Factory New"},
    {"key": "MW", "zh_tw": "略有磨損", "zh_cn": "略有磨损", "en": "Minimal Wear"},
    {"key": "FT", "zh_tw": "戰場實測", "zh_cn": "久经沙场", "en": "Field-Tested"},
    {"key": "WW", "zh_tw": "戰痕累累", "zh_cn": "破损不堪", "en": "Well-Worn"},
    {"key": "BS", "zh_tw": "重度磨損", "zh_cn": "战痕累累", "en": "Battle-Scarred"},
]

# 各種寫法 -> 內部 key 的對照，方便從 BUFF / Steam 名稱反查。
_EN_TO_KEY = {w["en"].lower(): w["key"] for w in WEARS}
_ZH_TO_KEY = {}
for _w in WEARS:
    _ZH_TO_KEY[_w["zh_tw"]] = _w["key"]
    _ZH_TO_KEY[_w["zh_cn"]] = _w["key"]

WEAR_ORDER = {w["key"]: i for i, w in enumerate(WEARS)}
WEAR_BY_KEY = {w["key"]: w for w in WEARS}
WANTED_KEYS = [w["key"] for w in WEARS]


def key_from_english(exterior_en: str):
    """由英文磨損名(如 'Field-Tested')取得內部 key，找不到回 None。"""
    if not exterior_en:
        return None
    return _EN_TO_KEY.get(exterior_en.strip().lower())


def key_from_market_hash_name(market_hash_name: str):
    """由完整 market_hash_name 尾端括號取磨損 key，例：
    'AK-47 | Redline (Field-Tested)' -> 'FT'。"""
    if not market_hash_name or "(" not in market_hash_name:
        return None
    exterior = market_hash_name.rsplit("(", 1)[-1].rstrip(")").strip()
    return key_from_english(exterior)


def key_from_chinese(exterior_zh: str):
    """由中文磨損名(繁或簡)取得 key。"""
    if not exterior_zh:
        return None
    return _ZH_TO_KEY.get(exterior_zh.strip())
