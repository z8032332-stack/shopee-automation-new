# 一次性工具：用已登入的 Chrome Profile 存蝦皮 Cookie（Playwright 版）
import json, asyncio
from playwright.async_api import async_playwright

COOKIES_FILE = r'D:\Users\user\Desktop\蝦皮影片專案\shopee_cookies.json'
USER_DATA    = r'C:\Users\user\AppData\Local\Google\Chrome\User Data'

async def main():
    async with async_playwright() as p:
        print('開啟 Chrome（用現有 Profile）...')
        context = await p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA,
            channel='chrome',
            headless=False,
            args=['--no-sandbox'],
        )
        page = context.pages[0] if context.pages else await context.new_page()
        await page.goto('https://shopee.tw')
        await page.wait_for_timeout(5000)

        if 'buyer/login' in page.url:
            print('\n尚未登入，請手動登入後按 Enter...')
            input()

        cookies = await context.cookies()
        with open(COOKIES_FILE, 'w', encoding='utf-8') as f:
            json.dump(cookies, f, ensure_ascii=False)

        print(f'\n成功存了 {len(cookies)} 個 Cookie → {COOKIES_FILE}')
        spc = [c for c in cookies if c['name'] in ('SPC_EC', 'SPC_U', 'SPC_F')]
        for c in spc:
            print(f'  {c["name"]}: OK')

        await context.close()

asyncio.run(main())
