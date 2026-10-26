import asyncio
from playwright.async_api import async_playwright
async def main():
    async with async_playwright() as p:
        b = await p.chromium.launch(executable_path='/opt/pw-browsers/chromium', args=['--no-sandbox'])
        page = await (await b.new_context(viewport={'width':1366,'height':900})).new_page()
        errs=[]; page.on('pageerror', lambda e: errs.append(str(e)))
        await page.goto('file:///home/claude/qhr-import/dist/index.html')
        await page.locator('input[type=file]').first.set_input_files('/home/claude/qhr-import/test/template_test.xlsx')
        await page.wait_for_selector('.filebox')
        await page.locator('input[type=file]').first.set_input_files('/home/claude/qhr-import/test/client_messy_sample.xlsx')
        await page.wait_for_function("document.querySelectorAll('.filebox').length===2")
        await page.get_by_role('button', name='التالي: مطابقة الأعمدة').click()
        await page.wait_for_selector('.maprow')
        await page.locator('.actions .btn.primary').click()
        await page.wait_for_selector('table.dt')
        st = lambda: page.locator('.stats').inner_text()
        print('before', (await st()).replace('\n',' '))
        # املأ بلد الإقامة
        await page.locator('.mapper').first.locator('select').first.select_option(label='السعودية')
        await page.get_by_role('button', name='تعبئة الفاضي').first.click()
        # صفّ الأخطاء
        await page.locator('.stat.error').click()
        print('errors filter rows:', await page.locator('table.dt tbody tr').count())
        # عدّل أول خلية خطأ نصية: الاسم الإنجليزي الفاضي → اكتب
        cell = page.locator('.cell.error').first
        print('first error cell tip:', await cell.get_attribute('title'))
        # افتح درج الصف
        await page.locator('.rownum').first.click()
        print('drawer:', (await page.locator('.drawer').inner_text()).replace('\n',' | ')[:300])
        await page.screenshot(path='/tmp/shots/7_drawer.png')
        await page.locator('.drawer-bg').click(position={'x':1300,'y':400})
        # استبعد صف عبر الدرج
        await page.locator('.rownum').first.click()
        await page.locator('.drawer .switch input').check()
        await page.locator('.drawer-bg').click(position={'x':1300,'y':400})
        print('after exclude', (await st()).replace('\n',' '))
        print('errors', errs)
        await b.close()
asyncio.run(main())
