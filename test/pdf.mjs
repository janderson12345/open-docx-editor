// PDF export in the browser (run with the e2e tests, after `npm run build`): each fixture opened in the demo and
// exported with src/pdf.js gives a PDF with a page per sheet, each the sheet's size.
import { chromium } from 'playwright';
import { PDFDocument, PDFName } from 'pdf-lib';
import fs from 'node:fs';
import path from 'node:path';
import { serveDemo, launchOptions } from './browser.mjs';

const { server } = serveDemo();
const launch = launchOptions();
const browser = await chromium.launch(launch);
const outDir = 'test/out';
fs.mkdirSync(outDir, { recursive: true });
let failed = 0;
for (const f of fs.readdirSync('test/fixtures').filter((n) => n.endsWith('.docx'))) {
  const page = await browser.newPage({ viewport: { width: 1300, height: 900 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto(`http://localhost:${server.address().port}/`);
  await page.waitForSelector('.odx-content p');
  await page.setInputFiles('input[type=file]', `test/fixtures/${f}`);
  await page.waitForFunction((n) => window.odx?.fileName === n, f);
  await page.waitForTimeout(1500);
  const sheets = await page.evaluate(() => [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => { const r = s.getBoundingClientRect(); return [r.width, r.height]; }));
  const bytes = Buffer.from(await page.evaluate(async () => [...await window.odx.exportPdf()]));
  fs.writeFileSync(path.join(outDir, f.replace(/\.docx$/, '.pdf')), bytes);
  const pdf = await PDFDocument.load(bytes);
  const sizes = pdf.getPages().map((p) => [p.getWidth() / 0.75, p.getHeight() / 0.75]);
  const ok = bytes.subarray(0, 5).toString() === '%PDF-' && sizes.length === sheets.length && sizes.every(([w, h], i) => Math.abs(w - sheets[i][0]) < 1 && Math.abs(h - sheets[i][1]) < 1);
  // Hebrew and Arabic in a font that has them (the Noto fallback, fetched when the run's font lacks them), not "?".
  const want = { 'rtl-hebrew.docx': 'NotoSansHebrew', 'rtl-arabic.docx': 'NotoSansArabic' }[f];
  const fontNames = pdf.context.enumerateIndirectObjects().map(([, o]) => o?.get?.(PDFName.of('BaseFont'))?.toString() || '').filter(Boolean);
  if (want && !fontNames.some((n) => n.includes(want))) { failed++; console.log(`FAIL ${f}: no ${want} fallback font (${fontNames.join(', ')})`); }
  if (!ok || errors.length) { failed++; console.log(`FAIL ${f}: ${sizes.length} PDF pages for ${sheets.length} sheets ${errors.join(' | ')}`); }
  else console.log(`ok   ${f}: ${sizes.length} pages, ${(bytes.length / 1024).toFixed(0)} KB`);
  await page.close();
}
await browser.close();
server.close();
process.exit(failed ? 1 : 0);
