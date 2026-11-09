// The React, Vue and Angular wrappers in a browser (run with the e2e tests): each test page (test/wrappers/) is bundled
// and, through its window.__wr, the wrapper is mounted (a blank document), given a document, made read-only and
// editable again, saved, and unmounted.
import { chromium } from 'playwright';
import { build } from 'esbuild';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';

const out = path.resolve('test/out/wrappers');
fs.mkdirSync(out, { recursive: true });
fs.copyFileSync('src/editor/editor.css', path.join(out, 'editor.css'));
const APPS = { react: 'test/wrappers/react-app.js', vue: 'test/wrappers/vue-app.js', angular: 'test/wrappers/angular-app.ts' };
for (const [name, entry] of Object.entries(APPS)) {
  await build({
    entryPoints: [entry], bundle: true, format: 'esm', outfile: path.join(out, `${name}.js`), logLevel: 'error',
    define: { 'process.env.NODE_ENV': '"production"', __VUE_OPTIONS_API__: 'true', __VUE_PROD_DEVTOOLS__: 'false', __VUE_PROD_HYDRATION_MISMATCH_DETAILS__: 'false' },
    tsconfigRaw: { compilerOptions: { experimentalDecorators: true, useDefineForClassFields: false } },
  });
  fs.writeFileSync(path.join(out, `${name}.html`), `<!doctype html><meta charset="utf-8"><link rel="icon" href="data:,"><link rel="stylesheet" href="editor.css"><odx-test-app></odx-test-app><div id="app"></div><script type="module" src="${name}.js"></script>`);
}

const server = http.createServer((req, res) => {
  const p = path.join(out, decodeURIComponent(req.url.split('?')[0]));
  if (!p.startsWith(out) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'content-type': { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' }[path.extname(p)] || 'application/octet-stream' });
  fs.createReadStream(p).pipe(res);
}).listen(0);
const base = `http://localhost:${server.address().port}/`;
const launch = process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : process.env.BROWSER_CHANNEL ? { channel: process.env.BROWSER_CHANNEL } : {};
const browser = await chromium.launch(launch);
const fixture = [...fs.readFileSync('test/fixtures/form.docx')];
let failed = 0;
const check = (ok, msg) => { if (!ok) { failed++; console.log('FAIL', msg); } };

for (const name of Object.keys(APPS)) {
  const page = await browser.newPage({ viewport: { width: 1300, height: 900 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()));
  await page.goto(`${base}${name}.html`);
  try {
    await page.waitForFunction(() => window.__wr && document.querySelector('.odx-content p'), null, { timeout: 15000 });
  } catch { check(false, `${name}: no editor mounted (${errors.join(' | ')})`); await page.close(); continue; }
  const log = () => page.evaluate(() => window.__wr.log.join(','));
  check((await log()).includes('ready'), `${name}: no ready event`);
  // A document given after mounting opens.
  await page.evaluate((bytes) => window.__wr.open(new Uint8Array(bytes)), fixture);
  await page.waitForFunction(() => window.__wr.log.some((l) => l.startsWith('opened')), null, { timeout: 15000 }).catch(() => {});
  check((await log()).includes('opened form.docx'), `${name}: document not opened (${await log()})`);
  // Read-only hides the ribbon and stops editing; off again restores both.
  const mode = () => page.evaluate(() => ({ bar: getComputedStyle(document.querySelector('.odx-toolbar')).display, editable: document.querySelector('.odx-content').getAttribute('contenteditable') }));
  check((await mode()).editable === 'true', `${name}: not editable at first`);
  await page.evaluate(() => window.__wr.readOnly(true));
  await page.waitForTimeout(200);
  const ro = await mode();
  check(ro.bar === 'none' && ro.editable === 'false', `${name}: read-only not applied (${JSON.stringify(ro)})`);
  await page.evaluate(() => window.__wr.readOnly(false));
  await page.waitForTimeout(200);
  check((await mode()).editable === 'true', `${name}: editing not restored`);
  // Saving through the wrapper's API gives a .docx.
  const saved = await page.evaluate(() => { const b = window.__wr.save(); return [b.length, b[0], b[1]]; });
  check(saved[0] > 1000 && saved[1] === 0x50 && saved[2] === 0x4b, `${name}: save gave ${saved}`);
  // Unmounting removes the editor.
  await page.evaluate(() => window.__wr.unmount());
  await page.waitForTimeout(200);
  check(!(await page.$('.odx-content')), `${name}: editor left after unmount`);
  check(!errors.length, `${name}: page errors: ${errors.join(' | ')}`);
  console.log(`${name}: ${await log()}`);
  await page.close();
}
await browser.close();
server.close();
console.log(failed ? `${failed} failures` : 'wrappers: all passed');
process.exit(failed ? 1 : 0);
