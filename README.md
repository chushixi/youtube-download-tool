# CS2 AK-47 Steam / BUFF 價格比較工具

爬取 CS2 遊戲中**所有 AK-47 皮膚**在五種磨損下，Steam 社群市集與 BUFF（网易 buff）
的 **最低販賣價、最高求購價、交易量**，輸出成一份可比較的 Excel。

五種磨損（依你的用語）：

| 內部碼 | 繁中 | BUFF 簡中 | Steam 英文 |
|---|---|---|---|
| FN | 全新出場 | 崭新出厂 | Factory New |
| MW | 略有磨損 | 略有磨损 | Minimal Wear |
| FT | 戰場實測 | 久经沙场 | Field-Tested |
| WW | 戰痕累累 | 破损不堪 | Well-Worn |
| BS | 重度磨損 | 战痕累累 | Battle-Scarred |

> ⚠️ 注意 BUFF 簡體「战痕累累」其實對應 **Battle-Scarred（重度磨損）**，
> 與繁中直覺相反。本工具一律以 `market_hash_name` 的英文磨損為準，避免誤判。

---

## ⚠️ 兩個重要限制（務必先讀）

1. **無法回溯歷史時間點。** Steam 與 BUFF 都**沒有**「查詢過去某個時間點的
   掛單簿（最低賣價 / 最高求購 / 在售量）」的接口。Steam 只有歷史「成交價曲線」，
   沒有歷史掛單快照。因此本工具抓的是**執行當下**的即時值，Excel 會記錄
   實際抓取時間（`抓取時間` 欄）。**你原先要的「6/10 18:00」那個時間點的資料，
   現在已無法取得。** 若要累積歷史資料，請用排程（cron / 工作排程器）每天固定
   時間跑一次並存檔（見下方「定時排程」）。

2. **本工具需在有網路的環境執行。** 產生它的雲端環境其網路政策封鎖了
   `steamcommunity.com` 與 `buff.163.com`（回 403），無法在該環境實際抓取。
   請在**你自己的電腦**上執行。

---

## 安裝

```bash
pip install -r requirements.txt
```

## 使用

### 1. 完整比較（Steam + BUFF，最推薦）

皮膚清單由 BUFF 依分類 `weapon_ak47` **動態發現**，自動涵蓋所有 AK-47 皮膚
（含日後新出的），並直接取得 BUFF 三項數據；再逐一去 Steam 抓對應數據。

```bash
python ak_price_compare.py --buff-cookie "session=你的BUFFsession值" -o ak.xlsx
# 或用環境變數
export BUFF_COOKIE="session=你的BUFFsession值"
python ak_price_compare.py -o ak.xlsx
```

**如何取得 BUFF Cookie**：瀏覽器登入 buff.163.com → F12 開發者工具 →
Application/应用 → Cookies → 複製 `session` 的值（形如 `session=1_xxxxxxxx`）。
Cookie 會過期，過期後重新複製即可。

### 2. 只抓 Steam（不需 Cookie）

皮膚改用內建後備清單（`csprice/skins.py`，未必涵蓋最新皮膚），BUFF 欄留空。

```bash
python ak_price_compare.py --steam-only -o ak.xlsx
```

### 3. 只抓 BUFF

```bash
python ak_price_compare.py --buff-only --buff-cookie "session=..." -o ak.xlsx
```

### 4. 自我測試（不連網，驗證 Excel 產出）

```bash
python ak_price_compare.py --self-test -o demo.xlsx
```

### 常用參數

| 參數 | 說明 |
|---|---|
| `-o, --output` | 輸出 Excel 路徑（預設 `ak_price_compare.xlsx`） |
| `--buff-cookie` | BUFF 登入 Cookie（或設 `BUFF_COOKIE`） |
| `--steam-only` / `--buff-only` | 只抓其中一邊 |
| `--currency` | Steam 幣別碼：`1`=USD、`3`=EUR、`23`=CNY（預設 23，與 BUFF 對齊） |
| `--steam-delay` | Steam 每請求間隔秒數（預設 3.0，太小會被 429 限流） |
| `--limit N` | 只處理前 N 筆，測試用 |
| `-v` | 顯示詳細進度 |

> **Steam 很慢且會限流**：Steam 對每個 IP 約每分鐘 20 次請求就開始 429。
> 每款皮膚每種磨損要 2~3 次請求，全部 AK-47（約 40 款 × 5 磨損 = 200 筆）
> 在預設 3 秒間隔下可能要跑 20~40 分鐘。被限流時工具會自動退避重試。

## 輸出欄位

| 欄位 | 來源 |
|---|---|
| 皮膚 / 磨損 / market_hash_name | 識別 |
| Steam最低賣價 | itemordershistogram（退回 priceoverview） |
| Steam最高求購 | itemordershistogram |
| Steam交易量(24h) | priceoverview volume |
| BUFF最低賣價 | goods `sell_min_price` |
| BUFF最高求購 | goods `buy_max_price` |
| BUFF在售量 | goods `sell_num`（BUFF 未提供公開「成交量」，以在售量近似） |
| 賣價差(Steam-BUFF) | 計算 |
| 賣價比(Steam/BUFF) | 計算 |
| 抓取時間 | 執行當下 |

## 定時排程（累積歷史資料）

因為抓的是即時值，若要重建「每天固定時間」的比較，請用排程每天存一份帶日期的檔：

```bash
# Linux crontab：每天 18:00 抓一次
0 18 * * * cd /path/to/tool && BUFF_COOKIE="session=..." \
  python ak_price_compare.py -o "ak_$(date +\%Y\%m\%d).xlsx"
```

## 專案結構

```
ak_price_compare.py     # CLI 進入點
csprice/
  wears.py              # 五種磨損對照與判定
  skins.py              # AK-47 後備皮膚清單
  steam.py              # Steam 抓取（含 item_nameid 解析、限流退避）
  buff.py               # BUFF 抓取（分類發現 + 解析）
  pipeline.py           # 合併 Steam/BUFF 成列
  excel.py              # 輸出 Excel
requirements.txt
```
