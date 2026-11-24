// Export hi-res images of specific elements for the PPT deck
const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1280, height: 720 }, deviceScaleFactor: 2.5 });
  await p.goto('file://' + path.resolve('supp1.html'));
  await p.evaluate(() => document.fonts.ready);
  const pages = await p.$$('.page');
  const cols = await pages[1].$$('.pcol .phone');
  for (let i = 0; i < cols.length; i++) await cols[i].screenshot({ path: `img/phone${i + 1}.png`, omitBackground: true });
  await b.close();
})();
