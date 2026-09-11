'use strict';

// Figure composition uses existing synthetic GUI pixels, never a recreated UI.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');

const USAGE = 'Usage: node compose_workspace.cjs --capture-dir <synthetic-captures> --output-dir <figure-dir>';
const CANVAS_WIDTH = 1080;
const CONTENT_WIDTH = 1032;
const PRINT_WIDTH_INCHES = 6.5;
const CROPS = {
  parameters: { x: 220, y: 584, width: 660, height: 364 },
  bev: { x: 388, y: 232, width: 1020, height: 684 }
};

function parseArgs(args) {
  const parsed = {};
  for (let i = 0; i < args.length; i += 2) {
    const key = { '--capture-dir': 'captureDir', '--output-dir': 'outputDir' }[args[i]];
    if (!key || parsed[key] || !args[i + 1] || args[i + 1].startsWith('--')) throw Error(USAGE);
    parsed[key] = path.resolve(args[i + 1]);
  }
  if (!parsed.captureDir || !parsed.outputDir) throw Error(USAGE);
  return parsed;
}

function pngDimensions(bytes) {
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  if (bytes.length < 33 || !bytes.subarray(0, 8).equals(signature) ||
      bytes.readUInt32BE(8) !== 13 || bytes.toString('ascii', 12, 16) !== 'IHDR')
    throw Error('Capture is not a PNG with an IHDR header.');
  return { width: bytes.readUInt32BE(16), height: bytes.readUInt32BE(20) };
}

function validateCaptureSize(dimensions) {
  if (dimensions.width !== 1600 || dimensions.height !== 1000)
    throw Error(`Capture is ${dimensions.width}x${dimensions.height}; reviewed CSS crops require exactly 1600x1000. Compact 1164x861 captures clip the plot axes/BEV footer and are not accepted.`);
  return dimensions;
}

function loadCapture(directory, name) {
  const candidates = [name + '.png', 'publication-' + name + '.png'];
  const existing = candidates.filter(candidate => fs.existsSync(path.join(directory, candidate)));
  if (existing.length !== 1) throw Error(`Expected exactly one ${name} capture (${candidates.join(' or ')}).`);
  const file = path.join(directory, existing[0]);
  if (!fs.statSync(file).isFile() || fs.statSync(file).size > 50 * 1024 * 1024)
    throw Error(`${name} capture is not a bounded PNG file.`);
  const bytes = fs.readFileSync(file);
  return { filename: existing[0], bytes, sha256: hash(bytes), ...validateCaptureSize(pngDimensions(bytes)) };
}

function hash(bytes) { return crypto.createHash('sha256').update(bytes).digest('hex'); }

function cropHtml(capture, crop, alt) {
  const scale = CONTENT_WIDTH / crop.width;
  return `<div class="capture" style="height:${crop.height * scale}px" data-crop="${crop.x},${crop.y},${crop.width},${crop.height}"><img alt="${alt}" src="data:image/png;base64,${capture.bytes.toString('base64')}" style="width:${capture.width * scale}px;height:${capture.height * scale}px;left:${-crop.x * scale}px;top:${-crop.y * scale}px"></div>`;
}

function composeHtml(captures) {
  const printScale = PRINT_WIDTH_INCHES * 96 / CANVAS_WIDTH;
  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline';">
<meta name="description" content="Two CSS-cropped captures of the publication-dual-layer synthetic ClearPlan fixture. No UI pixels have been repainted.">
<title>Figure 2 - ClearPlan synthetic workspace</title>
<style>
*{box-sizing:border-box}html,body{margin:0;background:#fff;color:#202B38;font-family:'Segoe UI',Arial,sans-serif}
#figure{width:${CANVAS_WIDTH}px;padding:24px;background:#fff}
header{display:flex;align-items:center;justify-content:space-between;gap:24px;margin-bottom:20px;height:40px}
header strong{font-size:24px;line-height:32px;font-weight:600;color:#16324A}
.mode{font-size:16px;line-height:20px;text-align:right;color:#77520F;border-left:4px solid #D99B36;padding-left:12px;font-weight:600}
section+section{margin-top:24px}h2{display:flex;align-items:center;gap:12px;font-size:24px;line-height:32px;font-weight:600;margin:0 0 12px;color:#16324A}
h2 b{display:inline-flex;width:32px;height:32px;align-items:center;justify-content:center;background:#0F766E;color:#fff;font-size:24px;font-weight:600}
.capture{position:relative;width:${CONTENT_WIDTH}px;overflow:hidden;background:#F3F5F7}
.capture img{position:absolute;display:block;max-width:none}
footer{font-size:16px;line-height:24px;color:#667585;margin-top:20px;padding-top:12px;border-top:1px solid #D6DDE4}
@page{size:${PRINT_WIDTH_INCHES}in __PRINT_HEIGHT_IN__in;margin:0}
@media print{html,body{width:${PRINT_WIDTH_INCHES * 96}px}#figure{zoom:${printScale};break-inside:avoid}}
</style></head><body><main id="figure">
<header><strong>ClearPlan | Synthetic workspace</strong><div class="mode">SIMULATED DATA<br>NOT FOR CLINICAL USE</div></header>
<section aria-labelledby="panel-a"><h2 id="panel-a"><b>A</b>Estimated plan rate across control points</h2>
${cropHtml(captures.parameters, CROPS.parameters, 'Actual synthetic GUI crop: nominal 800 MU/min, estimated segment rate, and the displayed 12-degree-per-second profile assumption.')}</section>
<section aria-labelledby="panel-b"><h2 id="panel-b"><b>B</b>Field-start BEV with two jawless MLC layers</h2>
${cropHtml(captures.bev, CROPS.bev, 'Actual synthetic GUI crop: first-control-point beam-eye view, both staggered MLC layers, isocenter coordinates, CT proxy, and field metadata.')}</section>
<footer>Actual application captures, cropped only | publication-dual-layer | Synthetic demonstration</footer>
</main></body></html>`;
}

function selfTest() {
  const tests = [
    ['requires both explicit directories', () => assert.throws(() => parseArgs([]))],
    ['rejects an unknown option', () => assert.throws(() => parseArgs(['--capture-dir', 'in', '--other', 'out']))],
    ['rejects a missing option value', () => assert.throws(() => parseArgs(['--capture-dir', '--output-dir', 'out']))],
    ['rejects duplicate options', () => assert.throws(() => parseArgs(['--capture-dir', 'in', '--capture-dir', 'out']))],
    ['resolves the two supplied directories', () => assert.deepEqual(parseArgs(['--capture-dir', 'in', '--output-dir', 'out']), { captureDir: path.resolve('in'), outputDir: path.resolve('out') })],
    ['rejects a non-PNG input', () => assert.throws(() => pngDimensions(Buffer.from('not a capture')))],
    ['rejects clipped compact captures', () => assert.throws(() => validateCaptureSize({ width: 1164, height: 861 }))],
    ['rejects an unexpected large viewport', () => assert.throws(() => validateCaptureSize({ width: 1920, height: 1080 }))],
    ['accepts the reviewed viewport', () => assert.deepEqual(validateCaptureSize({ width: 1600, height: 1000 }), { width: 1600, height: 1000 })]
  ];
  let failed = 0;
  for (const [name, test] of tests) {
    try { test(); console.log(`PASS ${name}`); }
    catch (error) { failed++; console.error(`FAIL ${name}: ${error.message}`); }
  }
  console.log(`compose_workspace: ${tests.length - failed} passed; ${failed} failed.`);
  process.exitCode = failed ? 1 : 0;
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  if (process.cwd().startsWith('\\\\')) throw Error('Run the Node/Playwright composer from a local working directory.');
  const captures = {
    parameters: loadCapture(options.captureDir, 'parameters'),
    bev: loadCapture(options.captureDir, 'bev')
  };
  const { chromium } = require('playwright');
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: CANVAS_WIDTH, height: 1600 }, deviceScaleFactor: 2 });
    await page.route('**/*', route => route.abort());
    let html = composeHtml(captures);
    await page.setContent(html);
    await page.evaluate(async () => {
      await document.fonts.ready;
      await Promise.all(Array.from(document.images, image => image.decode()));
      if (Array.from(document.images).some(image => image.naturalWidth !== 1600 || image.naturalHeight !== 1000))
        throw Error('Decoded capture dimensions do not match the reviewed viewport.');
    });
    const bounds = await page.locator('#figure').boundingBox();
    // CSS zoom affects print layout; transform alone can fragment a tall figure.
    const physicalHeight = (Math.ceil(bounds.height / CANVAS_WIDTH * PRINT_WIDTH_INCHES * 96) + 1) / 96;
    html = html.replaceAll('__PRINT_HEIGHT_IN__', physicalHeight.toFixed(6));
    await page.setContent(html);
    await page.evaluate(async () => { await document.fonts.ready; await Promise.all(Array.from(document.images, image => image.decode())); });
    const overflow = await page.locator('h2,header,footer').evaluateAll(nodes => nodes.some(node => node.scrollWidth > node.clientWidth + 1));
    if (overflow) throw Error('English panel caption overflows the composition.');
    fs.mkdirSync(options.outputDir, { recursive: true });
    const stem = path.join(options.outputDir, 'Figure_2_ClearPlan_workspace');
    fs.writeFileSync(stem + '.html', html, 'utf8');
    await page.locator('#figure').screenshot({ path: stem + '.png' });
    await page.pdf({ path: stem + '.pdf', preferCSSPageSize: true, printBackground: true });
    const manifest = {
      status: 'generated-requires-separate-source-and-visual-verification', fixture: 'publication-dual-layer',
      composition: 'two-vertically-stacked-CSS-crops', canvasWidthPx: CANVAS_WIDTH,
      canvasHeightPx: bounds.height, rasterScale: 2,
      printWidthInches: PRINT_WIDTH_INCHES, printHeightInches: physicalHeight,
      sourcePixelTreatment: 'original embedded PNGs; CSS crop and uniform scale only; no repainting, translated pixels, or generated UI',
      sources: Object.fromEntries(Object.entries(captures).map(([name, capture]) => [name,
        { filename: capture.filename, sha256: capture.sha256, width: capture.width, height: capture.height, crop: CROPS[name] }])),
      outputs: Object.fromEntries(['html', 'png', 'pdf'].map(extension => [extension,
        { filename: path.basename(stem) + '.' + extension, sha256: hash(fs.readFileSync(stem + '.' + extension)) }]))
    };
    fs.writeFileSync(stem + '.manifest.json', JSON.stringify(manifest, null, 2) + '\n', 'utf8');
    console.log(`Figure 2 exported: ${CANVAS_WIDTH * 2}px wide; PDF ${PRINT_WIDTH_INCHES} x ${physicalHeight.toFixed(2)} inches. Original GUI text remains unchanged. Source provenance and visual QA remain separate.`);
  } finally { await browser.close(); }
}

if (process.argv.slice(2).length === 1 && process.argv[2] === '--self-test') selfTest();
else main().catch(error => { console.error(error.message); process.exitCode = 1; });
