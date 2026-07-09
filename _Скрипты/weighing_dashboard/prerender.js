const { chromium } = require('playwright');
const fs = require('fs');
(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  const page = await browser.newPage({ viewport: { width: 1200, height: 1400 } });
  const errs = []; page.on('pageerror', e => errs.push(e.message));
  await page.goto('file://' + process.cwd() + '/dashboard_dynamic.html');
  await page.waitForFunction(() => {
    const k = document.getElementById('k-cycles');
    const t = document.querySelector('#tkph-grid .tkph-cell .v');
    return k && k.textContent.indexOf('—') === -1 && t && t.textContent.length > 0;
  }, { timeout: 8000 });
  await page.waitForTimeout(300);
  let html = await page.content();
  if (!html.toLowerCase().startsWith('<!doctype')) html = '<!DOCTYPE html>\n' + html;
  fs.writeFileSync('Дашборд весового контроля NTE200.html', html);
  console.log('prerendered', html.length, 'bytes, errors:', JSON.stringify(errs));
  await browser.close();
})();
