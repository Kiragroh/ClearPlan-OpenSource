// Deterministic, offline figure export; layout adapted from diagram-design's inline-SVG template.
const fs = require('node:fs');
const path = require('node:path');
const { chromium } = require('playwright');

async function main() {
  const root = path.resolve(__dirname, '../..');
  const out = path.resolve(process.argv[2] || path.join(root, 'paper/figures'));
  const svg = fs.readFileSync(path.join(__dirname, 'architecture.svg'), 'utf8');
  if (/https?:\/\//.test(svg.replace('http://www.w3.org/2000/svg', ''))) throw Error('External figure dependency');
  fs.mkdirSync(out, { recursive: true });
  const stem = path.join(out, 'Figure_1_ClearPlan_architecture');
  const html = `<!doctype html><html lang="en"><meta charset="utf-8"><title>ClearPlan architecture</title><style>*{box-sizing:border-box}html,body{margin:0;background:white}svg{display:block;width:1200px;height:820px}@page{size:1200px 820px;margin:0}</style>${svg}</html>`;
  fs.writeFileSync(stem + '.svg', svg, 'utf8');
  fs.writeFileSync(stem + '.html', html, 'utf8');
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1200, height: 820 }, deviceScaleFactor: 2.5 });
    await page.route('**/*', route => route.abort());
    await page.setContent(html);
    await page.evaluate(() => document.fonts.ready);
    const overflow = await page.locator('svg text').evaluateAll(nodes => nodes.filter(n => {
      const b = n.getBBox(); return b.x < 0 || b.y < 0 || b.x + b.width > 1200 || b.y + b.height > 820;
    }).map(n => n.textContent));
    if (overflow.length) throw Error('Figure overflow: ' + overflow.join(', '));
    await page.locator('svg').screenshot({ path: stem + '.png' });
    await page.pdf({ path: stem + '.pdf', preferCSSPageSize: true, printBackground: true });
    console.log('Architecture SVG/HTML/PNG/PDF exported; 9 nodes, 8 directed data links.');
  } finally { await browser.close(); }
}
main().catch(e => { console.error(e.message); process.exit(1); });
