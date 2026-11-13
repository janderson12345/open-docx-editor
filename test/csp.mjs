// Browser check of the editor self-hosted under a strict Content-Security-Policy (the one README.md gives): its stand-in
// fonts from the page's own copy (fonts.base), documents with pictures, shapes, a chart, text boxes and floats opened,
// every ribbon tab drawn, a PDF exported and the document saved, with no policy violation and no request to any other
// origin.
import { chromium, firefox, webkit } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';

const CSP = "default-src 'self'; img-src 'self' data: blob:; style-src 'self' 'unsafe-inline'; font-src 'self'; connect-src 'self' data: blob:; object-src 'none'; base-uri 'self'; frame-ancestors 'self'";
const root = path.resolve('demo');
const server = http.createServer((req, res) => {
  const p = path.join(root, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
  if (!p.startsWith(root) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
  const type = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.woff2': 'font/woff2' }[path.extname(p)] || 'application/octet-stream';
  res.writeHead(200, { 'content-type': type, 'content-security-policy': CSP });
  fs.createReadStream(p).pipe(res);
}).listen(0);
const origin = `http://localhost:${server.address().port}`;
// (BROWSER=firefox or webkit: in those engines, Playwright's own builds.)
const engine = { firefox, webkit }[process.env.BROWSER];
const browser = engine ? await engine.launch() : await chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : process.env.BROWSER_CHANNEL ? { channel: process.env.BROWSER_CHANNEL } : {});
const page = await browser.newPage({ viewport: { width: 1300, height: 950 } });
const outside = new Set();
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
page.on('request', (r) => { const u = r.url(); if (!u.startsWith(origin) && !/^(data|blob):/.test(u)) outside.add(u.slice(0, 120)); });
await page.addInitScript(() => {
  window.__violations = [];
  document.addEventListener('securitypolicyviolation', (e) => window.__violations.push(`${e.violatedDirective}: ${e.blockedURI || e.sample || ''} ${e.target?.outerHTML?.slice(0, 80) || ''}`));
});
let failed = 0;
const check = (ok, what) => { console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}`); if (!ok) failed++; };

await page.goto(origin + '/');
await page.waitForSelector('.odx-content p');
for (const f of ['rich.docx', 'shape-edit.docx', 'chart.docx', 'textbox.docx', 'watermark-picture.docx', 'wrap-float.docx', 'aptos.docx']) {
  await page.setInputFiles('input[type=file]', `test/fixtures/${f}`);
  await page.waitForFunction((n) => window.odx?.fileName === n, f);
  await page.waitForTimeout(1000);
}
for (const t of await page.locator('.odx-rtab').allTextContents()) { await page.click(`.odx-rtab:text-is("${t}")`); await page.waitForTimeout(150); }
const fonts = await page.evaluate(() => [...document.fonts].filter((f) => f.status === 'loaded').map((f) => f.family.replace(/["']/g, '')));
check(fonts.includes('odx-aptos'), `the stand-in fonts load from the page's own copy (${[...new Set(fonts)].join(', ')})`);
const pdf = await page.evaluate(async () => (await window.odx.exportPdf()).length);
const saved = await page.evaluate(async () => (await window.odx.save()).length);
check(pdf > 500 && saved > 1000, `a PDF exported (${pdf} bytes) and the document saved (${saved} bytes)`);
const violations = await page.evaluate(() => window.__violations);
check(!violations.length, `no Content-Security-Policy violations (${violations.join(' | ') || 'none'})`);
check(!outside.size, `no requests to other origins (${[...outside].join(' ') || 'none'})`);
check(!errors.length, `no page errors (${errors.join(' | ')})`);
await browser.close();
server.close();
console.log(failed ? `csp: ${failed} failed` : 'csp: all passed');
process.exit(failed ? 1 : 0);
