// Fonts: the missing-font report, and the bundled metric-compatible fonts standing in for one this computer lacks.
// Run after `npm run build` (serves demo/, which links dist/fonts/fonts.css).
import { chromium } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';

const kill = setTimeout(() => { console.log('timeout'); process.exit(2); }, 120000);
const root = path.resolve('demo');
const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.woff2': 'font/woff2' };
const server = http.createServer((req, res) => {
  const p = path.join(root, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
  if (!p.startsWith(root) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'content-type': types[path.extname(p)] || 'application/octet-stream' });
  fs.createReadStream(p).pipe(res);
}).listen(0);
const browser = await chromium.launch(process.env.BROWSER_CHANNEL ? { channel: process.env.BROWSER_CHANNEL } : process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {});
const page = await browser.newPage();
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
let failed = 0;
const check = (ok, msg) => { console.log(`${ok ? 'ok  ' : 'FAIL'} ${msg}`); if (!ok) failed++; };
await page.goto(`http://localhost:${server.address().port}/`);
await page.waitForSelector('.odx-content p');
await page.waitForTimeout(1000);

// A run in a font nothing has: reported, with no substitute; a known font isn't.
await page.click('.odx-content p');
await page.keyboard.type('Plain text. ');
const report = await page.evaluate(async () => {
  const v = window.odx.view;
  const before = await window.odx.missingFonts();
  // Mark the typed text in a font no computer has, through the editor's own command.
  const { state } = v;
  const mark = state.schema.marks.font.create({ val: 'Nonesuch Sans Test' });
  v.dispatch(state.tr.addMark(1, state.doc.content.size - 1, mark));
  await new Promise((r) => setTimeout(r, 300));
  return { before, after: await window.odx.missingFonts() };
});
check(!report.before.some((f) => f.font === 'Nonesuch Sans Test'), 'no missing font before');
const miss = report.after.find((f) => f.font === 'Nonesuch Sans Test');
check(!!miss && miss.substitute === null, `a font nothing has is reported (${JSON.stringify(report.after)})`);

// The bundled stand-ins load from fonts.css: Carlito (Calibri's widths) is available whether or not it's installed.
const carlito = await page.evaluate(async () => {
  await document.fonts.load('16px "Carlito"', 'Calibri');
  const c = document.createElement('canvas').getContext('2d');
  const w = (f) => { c.font = `40px ${f}`; return c.measureText('mmmmmmmmmmlli1WQ').width; };
  return { has: w('"Carlito", monospace') !== w('monospace'), faces: [...document.fonts].filter((f) => f.family.replace(/"/g, '') === 'Carlito').length };
});
check(carlito.has && carlito.faces >= 4, `Carlito available from the bundled fonts (${JSON.stringify(carlito)})`);
check(!errors.length, `no page errors ${errors.join(' | ')}`);
await browser.close();
server.close();
clearTimeout(kill);
console.log(failed ? `${failed} failures` : 'fonts: all passed');
process.exit(failed ? 1 : 0);
