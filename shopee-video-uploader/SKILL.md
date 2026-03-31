---
name: shopee-video-uploader
description: 蝦皮短影音自動上傳技能。當使用者想要用 ADB 把後製完成的影片自動上傳到蝦皮短影音 APP（含自動輸入文案、搜尋加入分潤商品、關閉合拍拼接、發布、失敗重試）時，一定要使用這個 skill。觸發情境包括：「幫我上傳蝦皮短影音」、「把影片上傳到蝦皮 APP」、「自動上傳蝦皮影片」、「蝦皮短影音批次上傳」、「shopee_upload」。此 skill 涵蓋：ADB 推影片到手機、重啟蝦皮 APP、導航到短影音拍攝頁、從媒體庫選影片、跳過編輯頁、輸入關鍵字文案、關閉合拍/拼接開關、搜尋並加入分潤商品、發布、以及上傳失敗自動重試。
---

# 蝦皮短影音自動上傳 (shopee-video-uploader)

使用 `uiautomator2` + ADB 自動操作手機蝦皮 APP，把後製完成的影片批次上傳為短影音，並自動加入分潤商品。

## 環境需求

```
pip install openpyxl uiautomator2
```
- ADB：路徑 `D:\platform-tools\adb.exe`
- 手機：Samsung S25 FE，解析度 1080×2340
- 手機 IP：`192.168.0.26:43467`（ADB over WiFi，需先在開發者模式開啟）

## 檔案路徑設定

| 變數 | 預設值 |
|---|---|
| EXCEL_PATH | `D:\Users\user\Downloads\Annie\Claud\蝦皮專案\蝦皮分潤前100_整理版.xlsx` |
| VIDEO_DIR | `D:\Users\user\Desktop\蝦皮影片專案\output_final` |
| ADB_PATH | `D:\platform-tools\adb.exe` |
| SCREENSHOT_DIR | `D:\Users\user\Desktop\蝦皮影片專案\screenshots` |
| DEVICE | `192.168.0.26:43467` |
| PHONE_VIDEO_DIR | `/sdcard/DCIM/ShopeeUpload` |

## Excel 欄位說明

| 欄位名稱 | 用途 |
|---|---|
| `編號` | 流水號（用來顯示第幾部） |
| `品名` | 商品名稱（用於搜尋分潤商品） |
| `關鍵字文案` | 影片說明文字（上傳時填入，限 150 字） |

影片檔案命名規則：`row{excel_row:03d}_final.mp4`（如 `row002_final.mp4`）

## 使用方式

```bash
cd D:\Users\user\Desktop\蝦皮影片專案

# 批次上傳（預設最多 50 部）
python shopee_upload.py

# 只上傳指定 Excel 行號（測試用）
python shopee_upload.py --row 3

# 跑到發佈前停下（不真的發布，DRY-RUN）
python shopee_upload.py --dry-run

# 從第 5 筆開始，上傳 10 部
python shopee_upload.py --start 5 --count 10

# 指定手機裝置 IP:PORT
python shopee_upload.py --device 192.168.0.26:12345
```

## 上傳流程（每部影片）

```
1. 推影片到手機 DCIM/ShopeeUpload/（先清空舊檔）
2. 強制重啟蝦皮 APP
3. 點底部「短影音 tab」
4. 點右上角「+」按鈕
5. 點「媒體庫」→ 篩選「短影音」→ 點第一支影片
6. 下一步（影片預覽頁）
7. 下一步（跳過編輯頁）
8. 輸入關鍵字文案（限 150 字，用 uiautomator2 貼上）
9. 關閉「允許他人合拍」「允許他人拼接」開關
10. 搜尋並加入分潤商品（智慧縮短搜尋詞，自動過濾通用詞）
11. 發布
12. 偵測上傳失敗 → 自動重試（最多 3 次）
13. 清除手機上的影片
```

## 座標地圖（Samsung S25 FE, 1080×2340）

```python
COORD = {
    "shortvideo_tab":  (540, 2150),   # 底部直播短影音 tab
    "plus_btn":        (1000, 153),   # 右上角 + 按鈕
    "media_lib":       (862, 1849),   # 媒體庫按鈕
    "video_tab":       (540, 307),    # 短影音篩選 tab
    "first_video":     (133, 506),    # 第一個影片
}
```

> 如果換手機或 APP 版本改版，座標可能需要更新。執行 `shopee_explore.py` 可輔助探索新座標。

## 核心程式碼（shopee_upload.py）

```python
# -*- coding: utf-8 -*-
"""
蝦皮短影音批次上傳腳本
用法：
  python shopee_upload.py                        # 上傳全部待上傳的影片
  python shopee_upload.py --row 2                # 只上傳第2筆
  python shopee_upload.py --dry-run              # 跑到發佈前停下
  python shopee_upload.py --row 2 --dry-run      # 第2筆 dry-run
"""

import os, sys, time, re, argparse, subprocess

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

import openpyxl
import uiautomator2 as u2

EXCEL_PATH = r"D:\Users\user\Downloads\Annie\Claud\蝦皮專案\蝦皮分潤前100_整理版.xlsx"
VIDEO_DIR  = r"D:\Users\user\Desktop\蝦皮影片專案\output_final"
ADB_PATH   = r"D:\platform-tools\adb.exe"
SCREENSHOT_DIR = r"D:\Users\user\Desktop\蝦皮影片專案\screenshots"
DEVICE     = "192.168.0.26:43467"
PHONE_VIDEO_DIR = "/sdcard/DCIM/ShopeeUpload"

COORD = {
    "shortvideo_tab":  (540, 2150),
    "plus_btn":        (1000, 153),
    "media_lib":       (862, 1849),
    "video_tab":       (540, 307),
    "first_video":     (133, 506),
}


def adb(cmd):
    full = f'"{ADB_PATH}" -s {DEVICE} {cmd}'
    r = subprocess.run(full, shell=True, capture_output=True, timeout=30)
    out = r.stdout.decode('utf-8', errors='replace') if r.stdout else ''
    err = r.stderr.decode('utf-8', errors='replace') if r.stderr else ''
    return out + err


def screenshot(name):
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)
    path = os.path.join(SCREENSHOT_DIR, f"{name}.png")
    subprocess.run(
        f'"{ADB_PATH}" -s {DEVICE} exec-out screencap -p',
        shell=True, stdout=open(path, 'wb'), timeout=15
    )
    return path


def tap(x, y, wait=1.5):
    adb(f"shell input tap {x} {y}")
    time.sleep(wait)


def swipe(x1, y1, x2, y2, duration=300, wait=1):
    adb(f"shell input swipe {x1} {y1} {x2} {y2} {duration}")
    time.sleep(wait)


def connect_device():
    result = adb(f"connect {DEVICE}")
    print(f"  ADB: {result.strip()}")
    time.sleep(1)


def push_video(local_path):
    """推送影片到手機（先清空舊檔確保只有一個影片）"""
    adb('shell content delete --uri content://media/external/video/media --where "_data LIKE \'%ShopeeUpload%\'"')
    time.sleep(1)
    adb(f"shell rm -rf {PHONE_VIDEO_DIR}")
    time.sleep(1)
    adb(f"shell mkdir -p {PHONE_VIDEO_DIR}")
    filename = os.path.basename(local_path)
    remote = f"{PHONE_VIDEO_DIR}/{filename}"
    print(f"  推送影片: {filename}")
    adb(f"push \"{local_path}\" \"{remote}\"")
    adb(f'shell am broadcast -a android.intent.action.MEDIA_SCANNER_SCAN_FILE -d "file://{remote}"')
    adb(f'shell am broadcast -a android.intent.action.MEDIA_SCANNER_SCAN_FILE -d "file:///sdcard/DCIM"')
    time.sleep(5)
    check = adb(f'shell ls -la "{remote}"')
    print(f"  檔案確認: {check.strip()}")
    return remote


def cleanup_phone_video(remote_path):
    adb(f'shell rm -f "{remote_path}"')
    adb(f'shell am broadcast -a android.intent.action.MEDIA_SCANNER_SCAN_FILE -d "file://{remote_path}"')


def open_shopee():
    print("  開啟蝦皮...")
    adb("shell monkey -p com.shopee.tw -c android.intent.category.LAUNCHER 1")
    time.sleep(5)


def close_all_apps():
    print("  關閉所有分頁...")
    adb("shell input keyevent KEYCODE_APP_SWITCH")
    time.sleep(2)
    d = u2.connect(DEVICE)
    btn = d(text='關閉全部')
    if btn.exists(timeout=3):
        btn.click()
        time.sleep(2)
    else:
        adb("shell input keyevent KEYCODE_HOME")
        time.sleep(1)


def find_and_click(d, texts, timeout=5, label="元素"):
    if isinstance(texts, str):
        texts = [texts]
    for txt in texts:
        el = d(textContains=txt)
        if el.exists(timeout=timeout):
            el.click()
            return True
    print(f"  ⚠ 找不到{label}: {texts}")
    return False


def close_popup(d):
    time.sleep(2)
    for _ in range(3):
        if d(text='短影音').exists(timeout=1) or d(text='推薦').exists(timeout=1):
            break
        tap(540, 1900, wait=1)
    time.sleep(1)


def navigate_to_create(d):
    print("  進入短影音...")
    tap(*COORD["shortvideo_tab"], wait=3)
    close_popup(d)
    print("  點 + 按鈕...")
    tap(*COORD["plus_btn"], wait=3)
    screenshot("after_plus")


def select_video_from_gallery(d):
    print("  點媒體庫...")
    tap(*COORD["media_lib"], wait=3)
    screenshot("gallery_page")
    print("  篩選短影音...")
    tap(*COORD["video_tab"], wait=2)
    print("  選第一個影片...")
    tap(*COORD["first_video"], wait=3)
    print("  下一步（影片預覽）...")
    if not find_and_click(d, ['下一步'], timeout=5, label="影片預覽下一步"):
        tap(956, 2120, wait=3)
    time.sleep(3)


def skip_editor(d):
    print("  跳過編輯頁...")
    time.sleep(3)
    screenshot("editor_page")
    if not find_and_click(d, ['下一步'], timeout=8, label="編輯頁下一步"):
        tap(900, 2060, wait=3)
    time.sleep(5)
    screenshot("after_editor")


def enter_caption(d, caption_text):
    print("  輸入文案...")
    found = False
    for txt in ['為您的短影音撰寫內文', '撰寫內文']:
        el = d(textContains=txt)
        if el.exists(timeout=3):
            el.click()
            time.sleep(1)
            found = True
            break
    if not found:
        tap(540, 300, wait=1)
    if len(caption_text) > 150:
        caption_text = caption_text[:150]
    d.send_keys(caption_text)
    time.sleep(1)
    ok = d(text='OK')
    if ok.exists(timeout=3):
        ok.click()
        time.sleep(2)
    else:
        tap(540, 600, wait=1)
    screenshot("after_caption")
    print(f"  ✓ 文案已輸入（{len(caption_text)}字）")


def toggle_off_switches(d):
    print("  關閉合拍/拼接...")
    for label in ['允許他人合拍', '允許他人拼接']:
        el = d(text=label)
        if el.exists(timeout=3):
            bounds = el.info['bounds']
            toggle_y = (bounds['top'] + bounds['bottom']) // 2
            tap(980, toggle_y, wait=1)
            print(f"    ✓ 已關閉 {label}")
    screenshot("after_toggle")


def _build_search_terms(product_name):
    """根據品名建立由長到短的搜尋詞，過濾通用詞"""
    clean = re.sub(r'[^\w\s]', ' ', product_name)
    clean = re.sub(r'\s+', ' ', clean).strip()
    words = clean.split()
    skip_words = {'現貨', '免運', '台灣', '出貨', '限時', '特價', '熱賣', '新款',
                  '隔日達', '當日', '預購', '批發', '包郵', '直送', '即日',
                  '近日', '到貨', '工廠', '直營', '正品', '爆款'}
    filtered = [w for w in words if w not in skip_words]
    terms = []
    if len(filtered) >= 5: terms.append(' '.join(filtered[:5]))
    if len(filtered) >= 3: terms.append(' '.join(filtered[:3]))
    if len(filtered) >= 2: terms.append(' '.join(filtered[:2]))
    if len(filtered) >= 1: terms.append(filtered[0])
    if not terms and words:
        terms = [' '.join(words[:3]), words[0]]
    return terms


def _navigate_back_to_publish(d):
    for _ in range(5):
        if d(textContains='撰寫內文').exists(timeout=1): return
        if d(textContains='發佈').exists(timeout=1): return
        cancel = d(text='取消')
        if cancel.exists(timeout=1):
            cancel.click()
            time.sleep(1)
            return
        d.press('back')
        time.sleep(2)


def add_product(d, product_name):
    print(f"  新增商品: {product_name[:30]}...")
    for _ in range(3):
        btn = d(textContains='新增商品')
        if btn.exists(timeout=2):
            btn.click()
            time.sleep(5)
            break
        swipe(540, 1500, 540, 800, 300, wait=1)
    else:
        print("  ⚠ 找不到新增商品按鈕")
        return False
    screenshot("add_product_page")
    tab = d(textContains='推廣分潤')
    if tab.exists(timeout=3):
        tab.click()
        time.sleep(2)
    search_terms = _build_search_terms(product_name)
    for term in search_terms:
        print(f"    搜尋: {term}")
        el = d(className='android.widget.EditText')
        if el.exists(timeout=3):
            el.click()
            time.sleep(0.5)
            el.clear_text()
            time.sleep(0.5)
            el.set_text(term)
            time.sleep(1)
            d.press('enter')
            time.sleep(3)
            if d(textContains='沒有搜尋結果').exists(timeout=2):
                print(f"    ✗ 沒結果，縮短再試")
                continue
            add_btn = d(text='加入')
            if add_btn.exists(timeout=3):
                add_btn[0].click()
                print("    ✓ 已點加入")
                time.sleep(3)
                done = d(text='完成')
                if done.exists(timeout=5):
                    done.click()
                    time.sleep(2)
                    screenshot("after_product_done")
                d.press('back')
                time.sleep(2)
                if d(textContains='新增商品').exists(timeout=1):
                    d.press('back')
                    time.sleep(2)
                print("  ✓ 商品已加入")
                return True
            else:
                continue
    print("  ⚠ 找不到商品，返回發布頁")
    _navigate_back_to_publish(d)
    return False


def publish(d, dry_run=False):
    screenshot("before_publish")
    if dry_run:
        print("  🔸 DRY-RUN: 到此為止，不真的發佈")
        d.press('back')
        time.sleep(1)
        discard = d(text='Discard')
        if not discard.exists(timeout=2):
            discard = d(textContains='捨棄')
        if discard.exists(timeout=2):
            discard.click()
            time.sleep(2)
        return True
    print("  發佈中...")
    for txt in ['發佈', '發布', 'Publish']:
        btn = d(text=txt)
        if btn.exists(timeout=3):
            btn.click()
            time.sleep(5)
            return True
    print("  ⚠ 找不到發佈按鈕")
    screenshot("publish_fail")
    return False


def handle_upload_failure(d, max_retries=3):
    for attempt in range(max_retries):
        print(f"  檢查上傳狀態（嘗試 {attempt+1}/{max_retries}）...")
        time.sleep(5)
        screenshot(f"upload_check_{attempt}")
        fail = d(textContains='上傳失敗')
        if fail.exists(timeout=3):
            print("  ⚠ 上傳失敗，關閉全部分頁重試...")
            close_all_apps()
            open_shopee()
            time.sleep(3)
            tap(*COORD["shortvideo_tab"], wait=5)
            for i in range(30):
                time.sleep(2)
                if d(textContains='上傳中').exists: continue
                if d(textContains='上傳失敗').exists: break
                print("  ✓ 上傳完成！")
                return True
        else:
            uploading = d(textContains='上傳中')
            if uploading.exists:
                print("  上傳中，等待完成...")
                for i in range(30):
                    time.sleep(2)
                    if not d(textContains='上傳中').exists(timeout=2):
                        if not d(textContains='上傳失敗').exists(timeout=2):
                            print("  ✓ 上傳完成！")
                            return True
                        break
            else:
                print("  ✓ 上傳完成！")
                return True
    print("  ✗ 多次重試仍失敗")
    return False


def upload_one(row_data, dry_run=False):
    idx = row_data['編號']
    excel_row = row_data['excel_row']
    name = row_data['品名']
    caption = row_data.get('關鍵字文案', '') or ''
    video_path = os.path.join(VIDEO_DIR, f"row{excel_row:03d}_final.mp4")

    print(f"\n{'='*50}")
    print(f"📹 上傳第 {idx} 部: {name[:30]}...")

    if not os.path.exists(video_path):
        print(f"  ✗ 影片不存在: {video_path}")
        return False
    if not caption:
        print(f"  ✗ 沒有文案，跳過")
        return False

    remote_path = push_video(video_path)
    d = u2.connect(DEVICE)
    open_shopee()
    navigate_to_create(d)
    select_video_from_gallery(d)
    skip_editor(d)
    enter_caption(d, caption)
    toggle_off_switches(d)
    add_product(d, name)
    _navigate_back_to_publish(d)
    result = publish(d, dry_run=dry_run)

    if not dry_run and result:
        success = handle_upload_failure(d)
        if success:
            cleanup_phone_video(remote_path)
        return success

    cleanup_phone_video(remote_path)
    return result


def read_excel():
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    rows = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        data = dict(zip(headers, row))
        if data.get('編號') and data.get('品名'):
            data['excel_row'] = row_idx
            rows.append(data)
    return rows


def main():
    global DEVICE
    parser = argparse.ArgumentParser(description='蝦皮短影音批次上傳')
    parser.add_argument('--row', type=int, help='只上傳指定 Excel 行號')
    parser.add_argument('--dry-run', action='store_true', help='跑到發佈前停下')
    parser.add_argument('--start', type=int, default=1, help='從第幾筆開始')
    parser.add_argument('--count', type=int, default=50, help='上傳幾部')
    parser.add_argument('--device', type=str, default=DEVICE, help='手機 IP:PORT')
    args = parser.parse_args()
    DEVICE = args.device

    print("🚀 蝦皮短影音批次上傳")
    print(f"   裝置: {DEVICE}")
    print(f"   Dry-run: {args.dry_run}")

    connect_device()
    rows = read_excel()
    print(f"   Excel 共 {len(rows)} 筆資料")

    if args.row:
        targets = [r for r in rows if r['excel_row'] == args.row]
    else:
        targets = []
        for r in rows:
            if r['excel_row'] < args.start: continue
            video_path = os.path.join(VIDEO_DIR, f"row{r['excel_row']:03d}_final.mp4")
            if os.path.exists(video_path) and r.get('關鍵字文案'):
                targets.append(r)
            if len(targets) >= args.count: break

    print(f"   本次上傳: {len(targets)} 部\n")

    success_count = fail_count = 0
    for i, row_data in enumerate(targets):
        print(f"\n[{i+1}/{len(targets)}]", end="")
        try:
            ok = upload_one(row_data, dry_run=args.dry_run)
            if ok:
                success_count += 1
                print(f"  ✅ 第 {row_data['編號']} 部完成")
            else:
                fail_count += 1
                print(f"  ❌ 第 {row_data['編號']} 部失敗")
        except Exception as e:
            fail_count += 1
            print(f"  ❌ 第 {row_data['編號']} 部異常: {e}")
            try:
                close_all_apps()
            except:
                pass

    print(f"\n{'='*50}")
    print(f"📊 結果: 成功 {success_count} / 失敗 {fail_count} / 總共 {len(targets)}")
    print(f"{'='*50}")


if __name__ == '__main__':
    main()
```

## 與其他 skill 的關係

```
shopee-video-scraper   →  output_videos/row*.mp4（素材）
       ↓
shopee-video-producer  →  output_final/row*_final.mp4（後製品）
       ↓
shopee-video-uploader  →  蝦皮 APP 短影音（本 skill）
```

## 常見問題排查

| 問題 | 原因 | 解法 |
|---|---|---|
| ADB 連線失敗 | 手機 IP 或 Port 不對 | `adb connect 192.168.0.26:43467` 手動確認 |
| 點擊位置偏移 | APP 版本或手機解析度不同 | 更新 `COORD` 字典 |
| 找不到「新增商品」 | 頁面未完全載入或需捲動 | 腳本會自動向上捲動 3 次後放棄 |
| 搜尋不到商品 | 品名關鍵字太長或含特殊符號 | `_build_search_terms` 會自動縮短重試 |
| 上傳失敗 | 網路或 APP 問題 | 腳本會自動重啟 APP 重試最多 3 次 |
| Excel `關鍵字文案` 欄空白 | 尚未填寫文案 | 先用 `shopee-copywriter` skill 生成文案 |
