from pathlib import Path
from playwright.sync_api import sync_playwright
from auto import FISCAL_PROFILE_DIR

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(str(FISCAL_PROFILE_DIR), headless=True, viewport={"width": 1800, "height": 2200})
    page = context.new_page()
    for name, url in [
        ('hermes_est', 'ENXTPA-RMS/estimates/free-cash-flow/'),
        ('spgi_cf', 'NYSE-SPGI/financials/cash-flow-statement/annual/'),
        ('hermes_overview', 'ENXTPA-RMS/'),
        ('spgi_income', 'NYSE-SPGI/financials/income-statement/annual/'),
    ]:
        page.goto('https://fiscal.ai/company/' + url, wait_until='domcontentloaded', timeout=120000)
        page.wait_for_timeout(5000)
        if name == 'hermes_est':
            print(page.get_by_text('Annual', exact=True).evaluate_all('(els)=>els.map(e=>e.outerHTML)'), flush=True)
            page.get_by_text('Annual', exact=True).first.click()
            page.wait_for_timeout(3000)
        Path('generated/' + name + '.txt').write_text(page.locator('body').inner_text(), encoding='utf-8')
        Path('generated/' + name + '_tables.json').write_text(__import__('json').dumps(page.locator('table').evaluate_all('(els)=>els.map(e=>e.outerHTML)')), encoding='utf-8')
        print(name + ' saved', flush=True)
    context.close()
