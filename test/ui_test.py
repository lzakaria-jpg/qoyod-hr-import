import asyncio, sys
from playwright.async_api import async_playwright
async def main():
    async with async_playwright() as p:
        b = await p.chromium.launch(executable_path='/opt/pw-browsers/chromium', args=['--no-sandbox'])
        ctx = await b.new_context(viewport={'width':1366,'height':900}, accept_downloads=True)
        page = await ctx.new_page()
        errors=[]
        page.on('pageerror', lambda e: errors.append(str(e)))
        page.on('console', lambda m: errors.append('console:'+m.text) if m.type=='error' else None)
        await page.goto('file:///home/claude/qhr-import/dist/index.html')
        await page.wait_for_selector('.drop')
        await page.screenshot(path='/tmp/shots/1_empty.png')
        inputs = page.locator('input[type=file]')
        await inputs.nth(0).set_input_files('/home/claude/qhr-import/test/template_test.xlsx')
        await page.wait_for_selector('.filebox')
        await page.locator('input[type=file]').first.set_input_files('/home/claude/qhr-import/test/client_messy_sample.xlsx')
        await page.wait_for_function("document.querySelectorAll('.filebox').length===2")
        await page.screenshot(path='/tmp/shots/2_files.png', full_page=True)
        await page.get_by_role('button', name='التالي: مطابقة الأعمدة').click()
        await page.wait_for_selector('.maprow')
        await page.screenshot(path='/tmp/shots/3_map.png', full_page=True)
        # defaults button
        btn = page.get_by_role('button', name='تطبيق الافتراضيات الشائعة')
        if await btn.count(): await btn.click()
        await page.get_by_role('button', name=lambda n: False) if False else None
        await page.locator('.actions .btn.primary').click()
        await page.wait_for_selector('table.dt')
        await page.screenshot(path='/tmp/shots/4_review.png', full_page=False)
        print('stats:', (await page.locator('.stats').inner_text()).replace('\n',' '))
        # تعبئة بلد الإقامة للفاضي
        sel0 = page.locator('.mapper').first.locator('select').first
        await sel0.select_option(label='السعودية')
        await page.get_by_role('button', name='تعبئة الفاضي').first.click()
        await page.wait_for_timeout(300)
        print('stats mid:', (await page.locator('.stats').inner_text()).replace('\n',' '))
        # mapper: apply suggestion for first unmatched
        sel = page.locator('.mapper select')
        await page.screenshot(path='/tmp/shots/4b.png')
        n = await sel.count(); print('mapper rows', n)
        for i in range(n):
            s = page.locator('.mapper select').nth(0)
            opts = await s.locator('option').all_inner_texts()
            if len(opts)>1:
                await s.select_option(index=1)
            await page.wait_for_timeout(150)
        await page.screenshot(path='/tmp/shots/5_review_after.png')
        print('stats after:', (await page.locator('.stats').inner_text()).replace('\n',' '))
        await page.get_by_role('button', name='التالي: التصدير').click()
        await page.wait_for_selector('.card h2')
        await page.screenshot(path='/tmp/shots/6_export.png', full_page=True)
        async with page.expect_download() as d:
            await page.get_by_role('button', name=lambda x: False) if False else None
            await page.locator('button.btn.primary', has_text='تنزيل ملف الرفع').click()
        dl = await d.value
        await dl.save_as('/tmp/ui_out.xlsx'); print('download', dl.suggested_filename)
        async with page.expect_download() as d2:
            await page.get_by_text('تنزيل تقرير الملاحظات').click()
        dl2 = await d2.value; await dl2.save_as('/tmp/ui_report.xlsx'); print('download', dl2.suggested_filename)
        print('errors:', errors)
        await b.close()
asyncio.run(main())
