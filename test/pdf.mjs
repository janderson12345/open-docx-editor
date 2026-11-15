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
  // (effects.docx with the installed fonts, which have outlines: its reflected words' mirrored copies are drawn.)
  const context = await browser.newContext({ viewport: { width: 1300, height: 900 } });
  if (f === 'effects.docx') await context.grantPermissions(['local-fonts']).catch(() => {});
  const page = await context.newPage();
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
  // A reflection (effects.docx's "Reflected words", w14:reflection): each word's mirrored copy faded by a soft mask of
  // its own, two of them, when the run's font was an installed one (the standard fonts have no outlines to mirror).
  if (f === 'effects.docx' && fontNames.some((n) => /Arial|Liberation|Arimo/i.test(n))) {
    const masks = pdf.context.enumerateIndirectObjects().filter(([, o]) => o?.get?.(PDFName.of('SMask'))?.get?.(PDFName.of('S'))?.toString() === '/Luminosity').length;
    if (masks !== 2) { failed++; console.log(`FAIL ${f}: ${masks} reflection masks for the two reflected words`); }
  }
  if (!ok || errors.length) { failed++; console.log(`FAIL ${f}: ${sizes.length} PDF pages for ${sheets.length} sheets ${errors.join(' | ')}`); }
  else console.log(`ok   ${f}: ${sizes.length} pages, ${(bytes.length / 1024).toFixed(0)} KB`);
  await context.close();
}
await browser.close();
server.close();
process.exit(failed ? 1 : 0);
