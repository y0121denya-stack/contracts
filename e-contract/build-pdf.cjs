// important-checks.html から A4 の PDF を作る。
// 使い方：node e-contract/build-pdf.cjs
const path = require('path');
const { chromium } = require('playwright');

(async () => {
  const src = path.join(__dirname, 'important-checks.html');
  const out = path.join(__dirname, 'important-checks.pdf');
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.goto('file://' + src);
  await page.pdf({ path: out, format: 'A4', preferCSSPageSize: true, printBackground: true });
  await browser.close();
  console.log('作成しました: ' + out);
})();
