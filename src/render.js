// usage: node render.js <html> <pdf> [pngPrefix]
const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const [, , html, pdf, pngPrefix] = process.argv;
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1280, height: 720 } });
  await p.goto('file://' + path.resolve(html));
  await p.evaluate(() => document.fonts.ready);
  await p.pdf({ path: pdf, width: '1280px', height: '720px', printBackground: true });
  if (pngPrefix) {
    const els = await p.$$('.page');
    for (let i = 0; i < els.length; i++) await els[i].screenshot({ path: `${pngPrefix}-${i + 1}.png` });
  }
  await b.close();
})();
