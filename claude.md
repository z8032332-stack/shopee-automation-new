# 蝦皮自動化工具 - 專案說明

## 核心裝置
- **Samsung Galaxy A52s 5G**（Android 11，1080x2400）
- ADB 連線方式：WiFi（TCP/IP），port 5555
- **IP 會隨手機重開機變動**，每次上傳前必須確認 `.env` 的 `PHONE_IP`

---

## 環境變數（.env）— 公司電腦版
所有路徑集中在 `.env`，不寫死在程式碼。

```
EXCEL_PATH=D:\Users\user\Desktop\蝦皮影片專案\蝦皮關鍵字選品_2026年3-4月new.xlsx
OUTPUT_DIR=D:\Users\user\Desktop\蝦皮影片專案\output_videos
FINAL_DIR=D:\Users\user\Desktop\蝦皮影片專案\output_final
BGM_DIR=D:\Users\user\Desktop\蝦皮影片專案\music
COOKIES_FILE=D:\Users\user\Desktop\蝦皮影片專案\shopee_cookies.json
GEMINI_KEY=AIzaSyDch3A7cNovcocVUBeSBIzngiKDGN9A6UE
START_ROW=22
COL_NAME=2, COL_LINK=3, COL_COPY=8, COL_TITLE=9, COL_STATUS=10
TTS_VOICE=zh-TW-HsiaoChenNeural

# 關鍵字選品（shopee_keyword_scraper_home.py）
KEYWORD_OUTPUT=D:\Users\user\Desktop\蝦皮影片專案\蝦皮關鍵字選品_2026年4-5月.xlsx
KEYWORD_TARGET=50
KEYWORD_MIN_SALES=300
```

`.env` 已加入 `.gitignore`，不會被 commit。

---

## 主要程式（目前版本）

| 檔案 | 用途 |
|------|------|
| `shopee_keyword_scraper_home.py` | **關鍵字選品主程式**（Skill B），Playwright CDP，自動輪換關鍵字，輸出 Excel |
| `shopee_video_maker_home3.py` | **抓影片主程式**，undetected_chromedriver，抓評論影片存到 `_clips_XXX/` |
| `shopee_video_producer.py` | **後製主程式 v2**，讀 clips → resize → 合併 → TTS旁白 + 逐句字幕 + BGM |
| `clear_status.py` | 清除 Excel 狀態欄（「影片完成」→「clips_ok(3)」），重跑用 |
| `shopee_upload_a52s.py` | 蝦皮短影音批次上傳（A52s 專用，ADB WiFi） |
| `dump_gallery.py` | Debug 用，傾印 UI hierarchy 找座標 |

**廢棄不用（勿刪，留著備查）：**
`shopee_video_maker_home.py` / `shopee_video_maker_home2.py`

---

## 兩階段工作流程

```
Step 1 抓影片
  python shopee_video_maker_home3.py
  → Chrome 彈出 → 登入蝦皮 → 按 Enter 繼續
  → 結果存到 output_videos/_clips_XXX/clip_00~02.mp4
  → Excel 狀態寫 clips_ok(3)

Step 2 後製
  python shopee_video_producer.py
  → 讀 _clips_XXX/ → FFmpeg resize 1080x1920
  → edge-tts 逐句生成旁白 MP3
  → 字幕貼在畫面垂直正中央（逐句同步）
  → BGM 25% + TTS 100% 混音
  → 影片循環補足 TTS 時長 + 1.5s
  → 輸出 output_final/XXX_品名.mp4
  → Excel 狀態寫「影片完成」
```

---

## 後製程式重點設定（shopee_video_producer.py v2）

| 項目 | 設定 |
|------|------|
| TTS 聲音 | `zh-TW-HsiaoChenNeural`（.env 的 TTS_VOICE 可換） |
| 字幕位置 | 畫面垂直正中央 `(VIDEO_H - box_h) // 2` |
| 字幕顯示 | 一句一句，與 TTS 同步（start_time + duration） |
| 影片無標題文字 | ✅ 標題只用於上傳平台，不疊在影片上 |
| BGM 音量 | 25%（TTS 100%） |
| 影片尺寸 | 1080 × 1920（直式） |
| rate-limit 防護 | 每句 TTS 間隔 0.8s，失敗最多重試 3 次 |
| 斷句規則 | 以 `。！？` 切分，純標點/無中文的句子過濾掉 |

---

## Excel 欄位設定

| 欄位 | COL 編號 | 內容 |
|------|----------|------|
| 品名 | 2 | 商品名稱 |
| 蝦皮連結 | 3 | 商品短網址（affiliate） |
| 文案 | 8 | 旁白內容（TTS + 字幕來源） |
| 標題 | 9 | 上傳平台用標題（不出現在影片） |
| 狀態 | 10 | clips_ok(3) / 影片完成 / 後製失敗 |

---

## 已知問題與解法

| 問題 | 解法 |
|------|------|
| TTS 奇偶句失敗（rate-limit） | 每句間隔 0.8s + 重試 3 次 |
| TTS 純標點句失敗 | split_sentences 過濾無中文句 |
| PermissionError WinError 32 | `ignore_cleanup_errors=True` + `gc.collect()` |
| ChromeDriver 版本不符 | 下載 ChromeDriver 147 放 `C:\ffmpeg\bin\` |
| 短網址無法取得 shopid/itemid | driver.get() 先導航，再從 current_url 取 ID |
| clear_status 重跑 | `python clear_status.py` 清狀態後再跑 producer |

---

## 網站專案（smileladypicks.com）

**路徑：** `D:\Users\user\Desktop\蝦皮影片專案\sites\smileladypicks\`
**部署：** Cloudflare Pages（GitHub repo: z8032332-stack/smileladypicks）
**Hugo 版本：** v0.159.1 extended
**主題：** PaperMod（git submodule）

### 網站結構

```
content/
├── about.md                    關於我：微笑小姐是誰
├── posts/
│   ├── shopee-automation.md    蝦皮分潤30天：從手動到自動化
│   ├── shopee-is-it-for-you.md 這個副業適不適合你
│   └── shopee-30-days.md       日入$100真實真相（主打文）
├── beauty/
│   ├── hair-care.md            護髮好物（舊文恢復）
│   └── eye-makeup.md           眼部彩妝（舊文恢復）
├── home/
│   └── sleep-goods.md          睡眠改善好物（舊文恢復）
└── daily/
    ├── baby-shampoo.md         兒童洗髮精推薦
    ├── bath-mat.md             兒童浴室防滑墊
    └── bath-toys.md            浴室洗澡玩具
```

### 首頁版型（layouts/index.html）
4 區塊：Hero → 副業入口🔥 → 選物🛍️ → 輕教學💡

### 待辦

| 項目 | 狀態 |
|------|------|
| 域名換為 smileladypicks.com | ✅ 完成 |
| 舊文章 3 篇恢復 | ✅ 完成 |
| 新文章 7 篇上線 | ✅ 完成 |
| H1/H2/H3 + TOC + keywords | ✅ 完成 |
| Cloudflare 自訂域名綁定 | ✅ 用戶自行完成 |
| Google Analytics GA4 接入 | ⏳ 待辦（hugo.toml 第8行填 G-XXXXXXXXXX） |

### GA4 接入步驟（下次做）
1. 前往 analytics.google.com → 建立資源 → 取得 `G-XXXXXXXXXX`
2. 編輯 `hugo.toml` 第 8 行：`ID = "G-XXXXXXXXXX"`
3. `hugo --minify` → `git add -A` → `git commit` → `git push`

---

## 關鍵字選品工作流程（shopee_keyword_scraper_home.py）

```
前置：
  1. 關閉所有 Chrome 視窗
  2. 用 CDP 模式重開 Chrome（或請 Claude 幫開）：
     powershell -Command "Stop-Process -Name chrome -Force; Start-Sleep 2;
     Start-Process 'C:\Program Files\Google\Chrome\Application\chrome.exe'
     -ArgumentList '--remote-debugging-port=9222','--user-data-dir=C:\Users\user\AppData\Local\Google\Chrome\User Data'"
  3. 登入 affiliate.shopee.tw

執行：
  python shopee_keyword_scraper_home.py
  → 自動連 CDP port 9222
  → 關鍵字隨機順序，每個關鍵字最多 MAX_PER_KW 筆（預設 5）
  → 排除 product_history.json 中已選過的商品
  → 過濾有影片 + 有分潤連結
  → 輸出 Excel（KEYWORD_OUTPUT）
  → 更新 product_history.json
```

**Excel 欄位（10欄）：** 編號、品名、分潤連結、價格、分潤率、銷量、對應關鍵字、文案、標題、狀態

**商品不重複機制：**
- `product_history.json` 記錄每次選到的商品 ID + 日期
- 每次自動排除已選過的商品
- 關鍵字隨機順序，每關鍵字上限 `MAX_PER_KW=5` 筆，避免單一品類霸版

**Append 模式（補充用）：**
- `.env` 設 `KEYWORD_APPEND=1` → 接續寫入現有 Excel，編號自動接續
- 補完後記得改回 `KEYWORD_APPEND=0`（或移除）

**filter_excel_videos.py（影片數篩選）：**
- 查每筆商品的影片數，保留 >= MIN_VIDEO_COUNT 的前 KEEP_TOP 筆
- ⚠️ 目前 api/v4/item/get 只回傳 seller demo 影片（最多 1 支），非評論 clips
- 0 筆符合時自動保護，不覆寫 Excel

---

## 2026-04-18 今日進度（家裡電腦）

### 環境設定
- 建立 `.env.home` / `.env.company` 範本，路徑全用環境變數，換電腦只需 `copy .env.home .env`
- 家裡電腦 Excel：`蝦皮素材\蝦皮選品_2026年 (1).xlsx`
- 家裡電腦輸出：`蝦皮素材\影片輸出\`（clips）、`蝦皮素材\影片完成\`（成品）

### 抓影片（home3.py）
- 修正欄位：B品名/C連結/H文案/I標題/J狀態
- 加 http 跳過非連結列
- 跑出 13 組 clips（`_clips_001~013`），每組 2-3 支

### 文案生成（gen_copy.py）新增
- 用 Groq `llama-3.3-70b-versatile` 生成文案+標題
- 標題：無 emoji、限 140 字
- 文案：10 句，每句約 10 字（短句版，影片約 15-30 秒）
- GROQ_KEY 存 `.env`，不上傳 GitHub
- 成功生成 13 筆，已寫入 Excel H/I 欄

### 後製（shopee_video_producer.py）調整
- `MIN_CLIPS=1`（最少 1 支即可後製）
- `fps=24` + `-preset ultrafast`（加速渲染）
- FFmpeg timeout 120→300 秒（家裡電腦較慢）
- 寫到 `.tmp.mp4` 再改名，避免毀損
- `-movflags +faststart`（moov atom 前置，開頭不毀損）
- **新邏輯**：clip 平均分配時長 = TTS總時長 ÷ clip數（3支→各10秒、2支→各15秒）
- 今日完成：`002_clear淨洗髮.mp4`（30MB）

### 已知家裡電腦速度
- 每支影片後製約 30-40 分鐘（公司約 5-8 分鐘）
- 家裡電腦 CPU 慢，moviepy 渲染很吃重

---

## 2026-04-08 今日進度

### 影片上傳
- **A52s 批次上傳完成：成功 41 / 失敗 0**
- Excel：`蝦皮關鍵字選品_2026年3-4月new.xlsx`，第10~50筆
- 上傳腳本：`shopee_upload_a52s.py --phone a52s`

### 關鍵字選品
- 跑出 50 筆，存到 `蝦皮關鍵字選品_2026年4-5月.xlsx`
- `product_history.json` 已記錄 50 筆（下次自動排除）
- 問題：單一關鍵字（存錢筒/藍芽耳機）抓太多 → 已加 `MAX_PER_KW=5` 修正
- 明天重跑，清掉 product_history 或用 append 補足

### 待辦（明天）
- [ ] 重新跑選品（MAX_PER_KW=5 確保多樣性）
- [ ] 清掉 `product_history.json` 重跑，或 append 補齊 50 筆

---

## 注意事項
- 手機解鎖 PIN：`0000`
- S25 FE 使用前確認手機 IP（DHCP 每次可能不同）
- `.env` 不進 git
- 影片無標題文字，標題只用於上傳蝦皮時填入
