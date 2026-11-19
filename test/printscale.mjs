// Printing and PDF export of long documents of large tables (run with the e2e tests, after `npm run build`). A made
// document of shaded tables with a landscape page among portrait ones (so each page prints from a copy of its own,
// printcopies.js), at two lengths: what the print takes on is the same for each page whatever the length (the print
// layer hides the paragraphs it draws by one rule, not one each; a page's copy holds what that page shows, not the
// whole document), the print and the export give a page per sheet, and the document is as it was after.
import { chromium } from 'playwright';
import { PDFDocument } from 'pdf-lib';
import fs from 'node:fs';
import { serveDemo, launchOptions, openFixture, reporter } from './browser.mjs';
import { pack, para, run } from '../tools/doc-generator/ooxml.mjs';

const { check, finish } = reporter('printscale');
const portrait = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
const landscape = '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
const table = (rows, cols, tag) => {
  const w = Math.floor(9000 / cols);
  const cell = (r, c) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${(r + c) % 3 === 0 ? '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>' : ''}</w:tcPr>${para(run(`${tag} row ${r + 1} item ${c + 1}${r + c ? '' : '\t'}`))}</w:tc>`;
  const border = (s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`;
  return `<w:tbl><w:tblPr><w:tblW w:w="${w * cols}" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${`<w:gridCol w:w="${w}"/>`.repeat(cols)}</w:tblGrid>${Array.from({ length: rows }, (_, r) => `<w:tr>${Array.from({ length: cols }, (_, c) => cell(r, c)).join('')}</w:tr>`).join('')}</w:tbl>`;
};
// `n` tables of 40 rows before a landscape page and `n` after. (A cell of each ends with a tab: the editor puts its
// placeholder picture after it, ProseMirror-separator, which isn't a picture that can show outside its table.)
const make = (n) => {
  const part = (k) => Array.from({ length: n }, (_, i) => para(run(`Table ${k}.${i + 1}`)) + table(40, 6, `T${k}.${i + 1}`)).join('');
  const body = part(1) + `<w:p><w:pPr><w:sectPr>${portrait}</w:sectPr></w:pPr></w:p>`
    + para(run('A landscape page')) + table(8, 8, 'L') + `<w:p><w:pPr><w:sectPr>${landscape}</w:sectPr></w:pPr></w:p>`
    + part(2) + `<w:sectPr>${portrait}</w:sectPr>`;
  return pack({ body });
};

fs.mkdirSync('test/out', { recursive: true });
const { server, url } = serveDemo();
const browser = await chromium.launch(launchOptions());
const result = {};
for (const [name, n] of [['short', 1], ['long', 4]]) {
  const file = `test/out/printscale-${name}.docx`;
  fs.writeFileSync(file, make(n));
  const page = await browser.newPage({ viewport: { width: 1300, height: 900 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto(url);
  await page.waitForSelector('.odx-content p');
  await openFixture(page, file, 2000);
  const before = await page.evaluate(() => JSON.stringify(window.odx.view.state.doc.toJSON()));
  const sheets = await page.evaluate(() => document.querySelectorAll('.odx-sheets > .odx-sheet').length);
  // What the print works on, read as it starts (after the editor's own beforeprint).
  await page.evaluate(() => window.addEventListener('beforeprint', () => {
    const copies = [...document.querySelectorAll('.odx-print-pages > div')];
    const rules = Math.max(...[...document.querySelectorAll('.odx-printlines > style')].map((s) => s.sheet.cssRules.length));
    window.__printed = { copies: copies.length, elements: copies.reduce((t, c) => t + c.querySelectorAll('*').length, 0), rules, hidden: document.querySelectorAll('.odx-print-copy [data-odx-plh]').length };
  }));
  await page.emulateMedia({ media: 'print' });
  const printed = await PDFDocument.load(await page.pdf({ preferCSSPageSize: true, printBackground: true }));
  await page.emulateMedia({ media: 'screen' });
  const seen = await page.evaluate(() => window.__printed);
  const exported = await PDFDocument.load(Buffer.from(await page.evaluate(async () => [...await window.odx.exportPdf()])));
  const after = await page.evaluate(() => ({ doc: JSON.stringify(window.odx.view.state.doc.toJSON()), marks: document.querySelectorAll('[data-odx-plh]').length, layers: document.querySelectorAll('.odx-printlines').length }));
  result[name] = { sheets, ...seen };
  check(sheets > (n > 1 ? 12 : 3), `${name}: ${sheets} pages`);
  check(seen.copies === sheets && printed.getPageCount() === sheets, `${name}: printed a page per sheet, each from its copy (${printed.getPageCount()} pages, ${seen.copies} copies)`);
  check(exported.getPageCount() === sheets, `${name}: exported a page per sheet (${exported.getPageCount()})`);
  check(seen.rules === 1 && seen.hidden > 0, `${name}: the paragraphs the print layer draws hidden by one rule (${seen.rules} in a layer)`);
  check(after.doc === before && !after.marks && !after.layers, `${name}: the document as it was after printing and exporting, its marks and print layer gone`);
  check(!errors.length, `${name}: no page errors ${errors.join(' | ')}`);
  await page.close();
}
// A page's copy holds about as much in a long document as in a short one (each had held the whole document: 4 times
// the tables, 4 times as much in each copy, and 16 times in all).
const per = (r) => r.elements / r.copies;
check(per(result.long) < per(result.short) * 1.6, `a page's print copy as big in the long document as in the short (${per(result.long).toFixed(0)} and ${per(result.short).toFixed(0)} elements)`);
await browser.close();
server.close();
finish();
