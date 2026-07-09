const { chromium } = require('playwright');
const fs = require('fs');
(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  const page = await browser.newPage({ viewport: { width: 1100, height: 1400 } });
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.goto('file://' + process.cwd() + '/Дашборд весового контроля NTE200.html');
  await page.waitForFunction(() => {
    const k = document.getElementById('k-cycles');
    return k && k.textContent && k.textContent.indexOf('—') === -1;
  }, { timeout: 8000 });
  await page.waitForTimeout(300);
  // mark that this DOM is prerendered, so we could detect, and grab full HTML
  const html = await page.content();
  fs.writeFileSync('prerendered.html', html);
  console.log('prerendered bytes:', html.length, 'errors:', JSON.stringify(errors));
  await browser.close();
})();
