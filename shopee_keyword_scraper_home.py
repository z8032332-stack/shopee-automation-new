# [HOME COMPUTER VERSION] 家用電腦版本 - 2026年3-4月
# 與公司版差異：直接 CDP 連線至 Chrome port 9222，不使用 profile 登入
# OUTPUT: C:\Users\user\Desktop\蝦皮關鍵字選品_2026年3-4月.xlsx
import sys, io, asyncio, json, re, os, random
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from playwright.async_api import async_playwright
from urllib.parse import quote
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

KEYWORDS = [
    '雨鞋','雨傘','吸入式補蚊燈','switch 2','蘋果快充','藍芽喇叭',
    '馬丁鞋','日系薄外套','皮克敏髮夾','存錢筒','行李箱','點讀筆',
    '手機殼','行動電源','雨衣','水壺','藍芽耳機','除濕機','延長線',
    '拖鞋','外套','後背包','耳環','衛生紙','安全帽',
    '嬰幼兒玩具','背包','春夏裝','運動鞋','穿戴甲','adidas',
]
BLACKLIST = ['藥','酒','菸','棉花棒','化妝棉']
TARGET = 50; BATCH = 120; MIN_SALES = 300
OUTPUT = r'C:\Users\user\Desktop\蝦皮關鍵字選品_2026年3-4月.xlsx'

def is_bl(name): return any(k in name for k in BLACKLIST)
def pp(raw):
    try:
        v = int(raw)
        return str(v // 100000) if v > 100000 else str(v)
    except: return str(raw)

async def safe_eval(page, js, retries=3):
    for i in range(retries):
        try: return await page.evaluate(js)
        except:
            if i < retries - 1: await asyncio.sleep(1.5)
    return None

async def search_kw(page, kw, limit=20):
    enc = quote(kw)
    js = """
    (async () => {
      try {
        const r = await fetch('https://affiliate.shopee.tw/api/v3/offer/product/list?list_type=0&sort_type=1&page_offset=0&page_limit=LIMIT&keyword=KW', {credentials:'include'});
        const d = await r.json();
        return d;
      } catch(e) { return {error: e.toString()}; }
    })()
    """.replace('LIMIT', str(limit)).replace('KW', enc)
    d = await safe_eval(page, js) or {}
    items = []
    for it in (d.get('data') or {}).get('list') or []:
        info = it.get('batch_item_for_item_card_full') or {}
        name = info.get('name', '')
        if is_bl(name): continue
        sales = info.get('sold') or info.get('historical_sold') or 0
        if sales < MIN_SALES: continue
        price = pp(info.get('price') or info.get('price_min', 0))
        shop_id = str(info.get('shopid', ''))
        item_id = str(info.get('itemid', '') or it.get('item_id', ''))
        items.append({
            'name': name, 'price': price, 'sales': sales, 'keyword': kw,
            'comm_rate': it.get('default_commission_rate', '') or it.get('seller_commission_rate', ''),
            'affiliate_link': it.get('long_link', ''),
            'product_url': it.get('product_link', ''),
            'shop_id': shop_id, 'item_id': item_id,
            'has_video': bool((info.get('video_info_list') or [])),
        })
    return items

async def chk_video(page, sid, iid):
    if not sid or not iid: return False
    js = """
    (async () => {
      try {
        const r = await fetch('https://shopee.tw/api/v4/item/get?itemid=IID&shopid=SID', {credentials:'include'});
        const d = await r.json();
        return (d && d.data && d.data.video_info_list && d.data.video_info_list.length > 0);
      } catch(e) { return false; }
    })()
    """.replace('IID', str(iid)).replace('SID', str(sid))
    return bool(await safe_eval(page, js))

async def get_link(page, url):
    if not url: return ''
    enc = quote(url, safe='')
    js = """
    (async () => {
      try {
        const r = await fetch('https://affiliate.shopee.tw/api/v3/product/get_affiliate_link?product_url=URL', {credentials:'include'});
        const d = await r.json();
        return (d && d.data && (d.data.short_link || d.data.link)) || '';
      } catch(e) { return ''; }
    })()
    """.replace('URL', enc)
    return await safe_eval(page, js) or ''

def build_excel(products):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = '蝦皮關鍵字選品'
    thin = Side(style='thin', color='DDDDDD')
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdrs = ['品名','分潤連結','價格','分潤率','銷量','對應關鍵字','痛點','狀態']
    wds  = [55, 48, 10, 10, 10, 16, 30, 12]
    for ci, (h, w) in enumerate(zip(hdrs, wds), 1):
        c = ws.cell(1, ci, h)
        c.font = Font(name='微軟正黑體', bold=True, color='FFFFFF', size=11)
        c.fill = PatternFill('solid', fgColor='C0392B')
        c.border = bd
        c.alignment = Alignment(horizontal='center', vertical='center')
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 28
    for ri, p in enumerate(products, 2):
        fill = PatternFill('solid', fgColor='FFF5F5' if ri % 2 == 0 else 'FFFFFF')
        vals = [p.get('name',''), p.get('affiliate_link',''), f"${p.get('price','0')}",
                p.get('comm_rate',''), p.get('sales',0), p.get('keyword',''), '', '']
        for ci, val in enumerate(vals, 1):
            c = ws.cell(ri, ci, val); c.fill = fill; c.border = bd
            c.font = Font(name='Arial', size=10, color='0563C1', underline='single') if ci == 2 \
                     else Font(name='微軟正黑體', size=10)
            c.alignment = Alignment(horizontal='left' if ci <= 2 else 'center',
                                    vertical='center', wrap_text=(ci == 1))
        ws.row_dimensions[ri].height = 32
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:H{len(products)+1}'
    wb.save(OUTPUT)
    print(f'Excel 已儲存：{OUTPUT}')

async def main():
    print('=' * 50)
    print('蝦皮關鍵字選品 Skill B - 2026年3-4月')
    print('=' * 50)

    async with async_playwright() as p:
        print('\n連接 Chrome (port 9222)...')
        browser = await p.chromium.connect_over_cdp('http://127.0.0.1:9222')
        ctx = browser.contexts[0]

        # 找 affiliate 頁面
        pg_aff = next((pg for pg in ctx.pages if 'affiliate.shopee.tw' in pg.url), None)
        if not pg_aff:
            pg_aff = await ctx.new_page()
            await pg_aff.goto('https://affiliate.shopee.tw/offer/product_offer',
                              wait_until='domcontentloaded', timeout=20000)
            await asyncio.sleep(3)
        print(f'affiliate 頁面：{pg_aff.url}')

        # 找或開 shopee.tw 頁面（用來查影片）
        pg_main = next((pg for pg in ctx.pages
                        if pg.url.startswith('https://shopee.tw') and 'affiliate' not in pg.url), None)
        if not pg_main:
            pg_main = await ctx.new_page()
            await pg_main.goto('https://shopee.tw', wait_until='domcontentloaded', timeout=20000)
            await asyncio.sleep(2)
        print(f'shopee 頁面：{pg_main.url[:50]}')

        # API 測試
        test = await search_kw(pg_aff, '手機殼', limit=2)
        print(f'API 測試：{"OK - " + str(len(test)) + "筆" if test else "FAIL"}')

        # [2] 搜尋
        print(f'\n[2/4] 搜尋關鍵字（目標 {BATCH} 筆）...')
        all_p = []; seen = set()
        for kw in KEYWORDS:
            if len(all_p) >= BATCH: break
            items = await search_kw(pg_aff, kw, limit=20)
            added = 0
            for item in items:
                uid = f"{item['shop_id']}_{item['item_id']}"
                if uid in seen or uid == '_': continue
                seen.add(uid); all_p.append(item); added += 1
            print(f'  [{kw}] {added}筆 累計{len(all_p)}')
            await asyncio.sleep(random.uniform(0.4, 0.8))
        print(f'共收集 {len(all_p)} 筆')

        # [3] 過濾影片 + 補連結
        print('\n[3/4] 過濾影片 + 補充分潤連結...')
        valid = []; nv = 0; nl = 0
        for i, item in enumerate(all_p):
            if len(valid) >= TARGET: break
            print(f'  [{i+1:3}/{len(all_p)}] {item["name"][:22]}...', end=' ', flush=True)
            hv = item.get('has_video', False)
            if not hv:
                hv = await chk_video(pg_main, item['shop_id'], item['item_id'])
            if not hv:
                print('X 無影片'); nv += 1; continue
            print('V', end=' ', flush=True)
            lnk = item.get('affiliate_link', '')
            if not lnk:
                lnk = await get_link(pg_aff, item['product_url'])
                item['affiliate_link'] = lnk
            if not lnk:
                print('-> 無連結'); nl += 1; continue
            print(f'OK ({len(valid)+1}/{TARGET})')
            valid.append(item)
            await asyncio.sleep(0.2)
        print(f'\n有效 {len(valid)} | 無影片 {nv} | 無連結 {nl}')

        pass  # browser stays open

    # [4] Excel
    print('\n[4/4] 輸出 Excel...')
    build_excel(valid[:TARGET])
    print('\n完成！')

asyncio.run(main())
