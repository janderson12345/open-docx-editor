// Browser check of the smaller things Word draws (test/fixtures/mkfeatures.mjs, mkrtl.mjs, mkstrict.mjs), each laid
// out as Word's own PDF of it showed: picture bullets, line numbers, watermarks, a section's vertical alignment,
// emphasis marks and run borders making their lines taller, right-to-left paragraphs, a Strict document.
import { chromium, firefox } from 'playwright';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { serveDemo, launchOptions } from './browser.mjs';

const { server } = serveDemo();
const launch = launchOptions();
const browser = await chromium.launch(launch);
const page = await browser.newPage({ viewport: { width: 1300, height: 950 } });
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
let failed = 0;
const check = (ok, what) => { console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}`); if (!ok) failed++; };
await page.goto(`http://localhost:${server.address().port}/`);
await page.waitForSelector('.odx-content p');
const open = async (f) => {
  await page.setInputFiles('input[type=file]', `test/fixtures/${f}`);
  await page.waitForFunction((n) => window.odx?.fileName === n, f);
  await page.waitForTimeout(1200);
};

await open('numbering-legal.docx');
const labels = await page.evaluate(() => [...document.querySelectorAll('.odx-content p.odx-li')].map((p) => p.dataset.label));
check(labels.slice(0, 4).join(' ') === 'I. 1.1. 1.1.1 I.1.a.1', `legal numbering: ${labels.slice(0, 4).join(' ')}`);
const pic = await page.evaluate(() => { const p = document.querySelector('.odx-content p.odx-picbullet'); return p && getComputedStyle(p, '::before').content; });
check(/^url\("data:image\/svg\+xml/.test(pic || ''), 'a picture bullet shows its picture');

// Bold typed at the end of a list item carries on into the next item made with Enter, as in Word.
await page.click('.odx-content p.odx-li');
await page.keyboard.press('End');
await page.waitForTimeout(100); // (the caret move End makes reaches the editor as a selection change, after the key)
await page.keyboard.press('Control+b');
await page.keyboard.type(' more');
await page.keyboard.press('Enter');
await page.keyboard.type('carried');
await page.waitForTimeout(300);
const carried = await page.evaluate(() => { const w = document.createTreeWalker(document.querySelector('.odx-content'), NodeFilter.SHOW_TEXT); let n; while ((n = w.nextNode())) if (n.textContent === 'carried') return getComputedStyle(n.parentElement).fontWeight; return 'missing'; });
check(carried === '700', `bold carries into the list item Enter makes (${carried})`);

await open('line-numbers.docx');
const nums = await page.evaluate(() => [...document.querySelectorAll('.odx-lnnum')].map((s) => s.textContent));
// (Word's PDF: 32 on the first page, 10 on the second, then every second line (12 to 22) and 5 to 10 from w:start.)
check(nums.slice(0, 5).join(',') === '1,2,3,4,5' && nums.length === 54 && nums.slice(42, 48).join(',') === '12,14,16,18,20,22', `line numbers: ${nums.length} drawn, as Word's`);
const lnBox = await page.evaluate(() => { const s = document.querySelector('.odx-lnnum').getBoundingClientRect(); const c = document.querySelector('.odx-content').getBoundingClientRect(); return [s.right, c.left + parseFloat(getComputedStyle(document.querySelector('.odx-content')).paddingLeft)]; });
check(Math.abs(lnBox[1] - lnBox[0] - 24) < 1, `a line number ends 0.25in left of the text (${(lnBox[1] - lnBox[0]).toFixed(1)}px)`);

// Line numbers in columns (Word's prints of these, odx-test/mklinenumcols.ps1): each column's lines numbered left of
// it, a paragraph's lines carried to the next column numbered there; right of each column in a right-to-left section;
// a restart-each-page section after a continuous break going on down its page.
const lnAt = (n) => page.evaluate((t) => {
  const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
  return [...document.querySelectorAll('.odx-lnnum')].filter((s) => s.textContent === t).map((s) => {
    const r = s.getBoundingClientRect();
    const k = sheets.findIndex((q) => r.top >= q.top && r.top < q.bottom);
    return { sheet: k, left: r.left - sheets[k].left, right: r.right - sheets[k].left, top: r.top - sheets[k].top };
  });
}, String(n));
await open('line-numbers-columns.docx');
const colNums = await page.evaluate(() => [...document.querySelectorAll('.odx-lnnum')].map((s) => s.textContent));
const n35 = await lnAt(35);
// Column 1's numbers (right edge at 63pt, 84px) stop at its last line: carried lines' numbers had gone on below it,
// in the bottom margin (35, 36). (The count is a line either way with other systems' fonts: Edge on Windows draws
// Word's 139.)
const col1Max = await page.evaluate(() => { const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect()); return Math.max(...[...document.querySelectorAll('.odx-lnnum')].filter((s) => { const r = s.getBoundingClientRect(); const q = sheets.find((x) => r.top >= x.top && r.top < x.bottom); return q && Math.abs(r.right - q.left - 84) < 2; }).map((s) => Number(s.textContent))); });
// (Word: 1-68 on pages 1 and 2, 1-3 on page 3; 35 is column 2's first line, its right edge at 315.1pt, 420px.)
check(Math.abs(colNums.length - 139) <= 2 && col1Max <= 35 && n35.length === 2 && n35.every((p) => Math.abs(p.right - 420) < 1.5 && p.top < 120), `line numbers in columns: ${colNums.length} drawn, column 1 up to ${col1Max}, 35 atop column 2 (${n35.map((p) => `${p.right.toFixed(1)},${p.top.toFixed(0)}`).join(' ')})`);
await open('line-numbers-rtl.docx');
const r1 = (await lnAt(1))[0], r35 = (await lnAt(35))[0];
// (Word: left edges at 549.2pt and 297.2pt, 9pt right of each column.)
check(r1 && r35 && Math.abs(r1.left - 732.3) < 1.5 && Math.abs(r35.left - 396.2) < 1.5, `right-to-left columns: numbers right of each column (${r1?.left.toFixed(1)}, ${r35?.left.toFixed(1)})`);
await open('line-numbers-mixed.docx');
const mixNums = await page.evaluate(() => [...document.querySelectorAll('.odx-lnnum')].slice(0, 11).map((s) => s.textContent));
check(mixNums.join(',') === '1,2,3,4,5,6,7,8,9,10,11', `a new section restarting each page goes on down the page it starts on (${mixNums.join(',')})`);
// A paragraph longer than a column goes on over the columns after it, cut at each one's foot (Word: page 1's two
// columns and page 2's first). None of its lines or numbers below a column's foot (960px), where they had run on.
await open('columns-long-paragraph.docx');
const longCols = await page.evaluate(() => {
  const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
  const low = (els) => Math.max(...els.map((e) => { const r = e.getBoundingClientRect(); const q = sheets.find((x) => r.top >= x.top && r.top < x.bottom); return q ? r.top - q.top : 9999; }));
  return { sheets: sheets.length, num: low([...document.querySelectorAll('.odx-lnnum')]), parts: document.querySelectorAll('.odx-repeat').length };
});
check(longCols.sheets === 3 && longCols.num < 950 && longCols.parts >= 2, `a paragraph over three columns: ${longCols.parts} carried parts, 3 pages, the lowest number at ${longCols.num.toFixed(0)}px`);
// The same in a section of columns after a one-column section, and a table longer than a column (Word: rows 1-24,
// 25-49 on page 1, 50-74, 75-90 on page 2).
await open('columns-long-paragraph-section.docx');
const longSect = await page.evaluate(() => ({ sheets: document.querySelectorAll('.odx-sheets > .odx-sheet').length, parts: document.querySelectorAll('.odx-repeat').length }));
check(longSect.sheets === 3 && longSect.parts >= 2, `a long paragraph in a section of columns among others: ${longSect.parts} carried parts, ${longSect.sheets} pages`);
await open('columns-long-table.docx');
const longTable = await page.evaluate(() => {
  const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
  return { sheets: sheets.length, parts: document.querySelectorAll('.odx-repeat').length };
});
check(longTable.sheets === 2 && longTable.parts === 3, `a table over four columns: ${longTable.parts} carried parts, ${longTable.sheets} pages`);

// Page breaks as Word's prints of these show (test/fixtures/mkpagebreaks.mjs): where a paragraph of the given text is,
// by sheet, and its top and bottom on it.
const placeOf = (texts) => page.evaluate((ts) => {
  const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
  const ps = [...document.querySelectorAll('.odx-content p')];
  return { sheets: sheets.length, at: ts.map((t) => {
    const p = ps.find((q) => q.textContent === t);
    if (!p) return null;
    const r = p.getBoundingClientRect();
    const k = sheets.findIndex((q) => r.top >= q.top && r.top < q.bottom);
    return { sheet: k, top: r.top - (sheets[k]?.top ?? 0), bottom: r.bottom - (sheets[k]?.top ?? 0) };
  }) };
}, texts);
// Two floating tables a little narrower than the column: nothing wraps beside them, the second breaks at page 1's foot.
await open('pagebreak-float-table.docx');
const fts = await placeOf(['Row 23', 'Row 24', 'Row 40', 'After paragraph 1']);
check(fts.sheets === 2 && fts.at[0]?.sheet === 0 && fts.at[1]?.sheet === 1 && fts.at[3]?.sheet === 1 && fts.at[3].top >= fts.at[2].bottom, `a floating table breaks across pages row by row, the text after it below it (${JSON.stringify(fts)})`);
// A footer's closing paragraph at 0.06 lines takes a pixel, not a line: the page above it holds a line more.
await open('pagebreak-tiny-footer.docx');
const tf = await placeOf(['Body paragraph 54', 'Body paragraph 55']);
check(tf.at[0]?.sheet === 0 && tf.at[1]?.sheet === 1, `a footer's 0.06-line empty paragraph leaves the body its room (${JSON.stringify(tf.at)})`);
// A heading kept with a table longer than a page whose rows all keep with the next: on page 2 together, the table broken
// at that page's foot.
await open('pagebreak-keep-long.docx');
const kl = await placeOf(['Heading text', 'Row 1', 'Row 53', 'Row 54']);
check(kl.at.every(Boolean) && kl.at[0].sheet === 1 && kl.at[1].sheet === 1 && kl.at[2].sheet === 1 && kl.at[3].sheet === 2, `a heading kept with a long table starts the page with it (${JSON.stringify(kl.at.map((a) => a?.sheet))})`);
// A floating table set down from its text, its rows all keeping with the next, reaching past page 1's foot: broken there
// row by row (its keeps don't hold), the text after it under it on page 2.
await open('pagebreak-float-keeps.docx');
const fk = await placeOf(['Cell 4', 'Cell 5', 'Cell 11', 'After paragraph 1']);
check(fk.at.every(Boolean) && fk.at[0].sheet === 0 && fk.at[1].sheet === 1 && fk.at[2].sheet === 1 && fk.at[3].sheet === 1 && fk.at[3].top >= fk.at[2].bottom, `a floating table set down breaks at the page's foot, its keeps aside (${JSON.stringify(fk.at)})`);
// A table whose rows all keep with the next goes on with the paragraph after it that doesn't fit, though it fits itself.
await open('pagebreak-keep-table.docx');
const kt = await placeOf(['Lead text', 'R1 c1 text', 'R5 c3 text']);
check(kt.at.every(Boolean) && kt.at[0].sheet === 0 && kt.at[1].sheet === 1 && kt.at[2].sheet === 1, `a table kept with the next goes on with the paragraph after it (${JSON.stringify(kt.at.map((a) => a?.sheet))})`);

// Footnotes in columns (Word's prints of these, odx-test/mkrtlfootnotes.ps1): one band across each page's foot, in the
// section's columns (the first on the right in a right-to-left section), each note on the page of its reference; one
// band the text's width with w15:footnoteColumns 1; the separator at the right when its paragraph is right to left.
const bands = () => page.evaluate(() => [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => {
  const sr = s.getBoundingClientRect();
  const box = s.querySelector('.odx-notes');
  if (!box) return null;
  const sep = box.querySelector('.odx-note-sep').getBoundingClientRect();
  const cols = [...box.querySelectorAll('.odx-note-band > div')];
  const seen = (r, c) => r.top >= c.top - 1 && r.top < c.bottom - 1;
  const parts = cols.length ? cols.map((c) => { const cr = c.getBoundingClientRect(); return { x: Math.round(cr.left - sr.left), ids: [...c.querySelectorAll('.odx-note')].filter((n) => seen(n.getBoundingClientRect(), cr)).map((n) => n.dataset.id) }; })
    : [{ x: Math.round(box.getBoundingClientRect().left - sr.left), ids: [...box.querySelectorAll('.odx-note')].map((n) => n.dataset.id) }];
  return { sep: [Math.round(sep.left - sr.left), Math.round(sep.right - sr.left)], parts };
}));
await open('footnotes-columns-rtl.docx');
const fnr = await bands();
// (Word: page 1 notes 1 and 2 in the right column (x 324pt, 432px), note 3 in the left (x 72pt); page 2 notes 4 | 5.)
check(fnr[0] && fnr[0].parts.length === 2 && fnr[0].parts[0].x === 432 && fnr[0].parts[0].ids.join() === '1,2' && fnr[0].parts[1].x === 96 && fnr[0].parts[1].ids.join() === '3' && fnr[1]?.parts.map((q) => q.ids.join()).join('|') === '4|5', `footnotes in right-to-left columns: ${JSON.stringify(fnr.map((b) => b && b.parts))}`);
check(fnr[0] && fnr[0].sep[0] === 96 && fnr[0].sep[1] === 288, `the separator 2in long at the left (${fnr[0]?.sep})`);
await open('footnotes-columns-one-band.docx');
const f1 = await bands();
check(f1[0] && f1[0].parts.length === 1 && f1[0].parts[0].ids.join() === '1,2,3', `w15:footnoteColumns 1: one band, notes 1-3 on page 1 (${JSON.stringify(f1.map((b) => b && b.parts))})`);
// A note far too long for its page (fncols-word tall2a): the band stays in columns, under the text up to the reference
// evened out across them (paragraph 2, holding it, beside paragraph 1, not under it), each column filled, and the rest
// of the note in the next page's band, in columns again.
await open('footnotes-columns-tall.docx');
// (Its band is worked out over a few layout passes: until the first page's separator is above the middle.)
await page.waitForFunction(() => { const s = document.querySelector('.odx-sheets > .odx-sheet'); const sep = s?.querySelector('.odx-note-sep'); return sep && sep.getBoundingClientRect().top - s.getBoundingClientRect().top < s.getBoundingClientRect().height / 2; }, null, { timeout: 8000 }).catch(() => {});
await page.waitForTimeout(500);
const ft = (await page.evaluate(() => [...document.querySelectorAll('.odx-sheets > .odx-sheet')].slice(0, 2).map((s) => [...s.querySelectorAll('.odx-note-band > div')].map((c) => [...c.querySelectorAll('.odx-note')].map((n) => n.dataset.id)))));
const ftp = await page.evaluate(() => {
  const sh = document.querySelector('.odx-sheets > .odx-sheet');
  const s = sh.getBoundingClientRect();
  const sep = sh.querySelector('.odx-notes .odx-note-sep').getBoundingClientRect();
  const ps = [...document.querySelectorAll('.odx-content > p')].slice(0, 2).map((p) => { const r = p.getBoundingClientRect(); return { x: Math.round(r.left - s.left), y: Math.round(r.top - s.top) }; });
  return { sep: Math.round(sep.top - s.top), h: Math.round(s.height), ps };
});
const huge = ft[0]?.[0]?.[0];
check(!!huge && ft.length === 2 && ft.every((cs) => cs.length === 2 && cs.every((ids) => ids.includes(huge))), `a note too long for its page: in both band columns, going on in columns on page 2 (${JSON.stringify(ft)})`);
check(ftp.ps[0].x === 96 && ftp.ps[1].x === 432 && Math.abs(ftp.ps[0].y - ftp.ps[1].y) < 4 && ftp.sep < ftp.h / 2, `its reference's paragraph beside the one before, the band from under them (${JSON.stringify(ftp)})`);
await open('footnotes-separator-rtl.docx');
const fs2 = await bands();
check(fs2[0] && fs2[0].sep[0] === 528 && fs2[0].sep[1] === 720, `a right-to-left separator at the right (${fs2[0]?.sep})`);
// Compatibility mode 14 (Word 2010): each column its own note area and separator, under the column of the references
// (Word: page 1's right column notes 1 and 2, its separator at x 324pt; the left column note 3, at x 72pt).
await open('footnotes-columns-compat14.docx');
const c14 = await page.evaluate(() => {
  const s = document.querySelector('.odx-sheets > .odx-sheet');
  const sr = s.getBoundingClientRect();
  return [...s.querySelectorAll('.odx-notes')].map((b) => ({ x: Math.round(b.getBoundingClientRect().left - sr.left), sep: Math.round(b.querySelector('.odx-note-sep').getBoundingClientRect().left - sr.left), ids: [...b.querySelectorAll('.odx-note')].map((n) => n.dataset.id).join() })).sort((a, b) => b.x - a.x);
});
check(c14.length === 2 && c14[0].x === 432 && c14[0].ids === '1,2' && c14[1].x === 96 && c14[1].ids === '3', `compat 14: a note area under each column (${JSON.stringify(c14)})`);

await open('watermark-text.docx');
const wm = await page.evaluate(() => [...document.querySelectorAll('.odx-sheet > img')].map((i) => ({ t: getComputedStyle(i).transform, w: i.width })));
// (550px: the shape's 412.4pt width.)
const sheets = await page.evaluate(() => document.querySelectorAll('.odx-sheets > .odx-sheet').length);
check(wm.length === sheets && wm.every((x) => x.t !== 'none' && x.w === 550), `a text watermark on each of the ${sheets} pages, turned (${wm.length})`);

await open('valign.docx');
const vtop = await page.evaluate(() => { const p = document.querySelector('.odx-content > p'); const s = document.querySelector('.odx-sheet'); return p.getBoundingClientRect().top - s.getBoundingClientRect().top; });
check(vtop > 350 && vtop < 500, `a centred section's text sits in the middle of the page (${vtop.toFixed(0)}px down)`);

// Where a block is: its sheet (by the sheets' boxes) and its box on that sheet.
const sheetPlaces = () => page.evaluate(() => {
  const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
  return [...document.querySelectorAll('.odx-content > p')].map((p) => {
    const r = p.getBoundingClientRect();
    const n = sheets.findIndex((s) => r.top + 1 >= s.top && r.top < s.bottom);
    return { text: p.textContent, sheet: n, top: r.top - (sheets[n]?.top ?? 0), bottom: r.bottom - (sheets[n]?.top ?? 0), left: r.left - (sheets[n]?.left ?? 0), w: sheets[n]?.width ?? 0 };
  });
});
// A one-section document in two columns, centred (pages.js verticalAlignCols): its text moved down to the middle.
await open('valign-columns.docx');
const vcols = await sheetPlaces();
check(vcols[0].sheet === 0 && vcols[0].top > 250 && vcols[0].top < 500, `columns centred on their page: the text starts ${vcols[0].top.toFixed(0)}px down`);
// Two columns before their endnote (balanceEndCols): the second column used, the columns ending within a few lines of
// each other, and the endnote right under the lower of them (it had been drawn where the first pass left it, at the
// page's foot: drawSheets' key had no place for it).
await open('columns-endnotes.docx');
const ecols = (await sheetPlaces()).filter((x) => x.sheet === 0 && x.text.trim());
const half = ecols[0]?.w / 2;
const leftBottom = Math.max(...ecols.filter((x) => x.left < half).map((x) => x.bottom));
const rightCol = ecols.filter((x) => x.left >= half);
const rightBottom = rightCol.length ? Math.max(...rightCol.map((x) => x.bottom)) : -1;
const endnoteTop = await page.evaluate(() => {
  const n = document.querySelector('.odx-note[data-kind="endnote"]');
  const s = document.querySelector('.odx-sheets > .odx-sheet');
  return n && s ? n.getBoundingClientRect().top - s.getBoundingClientRect().top : null;
});
check(rightCol.length > 0 && Math.abs(leftBottom - rightBottom) < 80 && endnoteTop != null && endnoteTop >= Math.max(leftBottom, rightBottom) - 1 && endnoteTop < Math.max(leftBottom, rightBottom) + 40,
  `columns before an endnote end level and the endnote goes under them (${JSON.stringify({ leftBottom: Math.round(leftBottom), rightBottom: Math.round(rightBottom), endnoteTop: endnoteTop && Math.round(endnoteTop) })})`);
// Endnotes in right-to-left columns (fncols-word, Word's prints): on a page of their own after the text, at its top
// and in columns left to right (the first endnote in the left one), their separator at the right as its paragraph is
// right to left; the text before them not cut short for them. Going on onto a page of their own: a continuation
// separator over the first (left) column, the width of a column.
const endPage = () => page.evaluate(() => {
  const ss = [...document.querySelectorAll('.odx-sheets > .odx-sheet')];
  const s = ss[ss.length - 1], sr = s.getBoundingClientRect();
  const box = s.querySelector('.odx-notes');
  if (!box) return null;
  const sep = box.querySelector('.odx-note-sep').getBoundingClientRect();
  const cols = [...box.querySelectorAll('.odx-note-band > div:not(.odx-note-sep)')].map((c) => ({ x: Math.round(c.getBoundingClientRect().left - sr.left), ids: [...c.querySelectorAll('.odx-note')].map((n) => n.dataset.id) }));
  return { sheets: ss.length, sep: [Math.round(sep.left - sr.left), Math.round(sep.right - sr.left), Math.round(sep.top - sr.top)], cols, text: [...document.querySelectorAll('.odx-content > p')].filter((p) => p.getBoundingClientRect().top > sr.top).length };
});
await open('endnotes-columns-rtl.docx');
await page.waitForTimeout(1500);
const er = await endPage();
check(er && er.sheets === 3 && er.text === 0 && er.cols[0]?.x === 96 && er.cols[0].ids[0] === '1' && er.sep[0] === 528 && er.sep[2] < 120, `endnotes on a page of their own: at its top, left to right, the separator at the right (${JSON.stringify(er)})`);
await open('endnotes-columns-long-rtl.docx');
await page.waitForTimeout(1500);
const el = await endPage();
check(el && el.sheets === 3 && el.text === 0 && el.cols[0]?.x === 96 && el.sep[0] === 96 && el.sep[1] === 384 && el.sep[2] < 120, `endnotes going on onto a page of their own: a continuation separator a column wide over the left column (${JSON.stringify(el)})`);
// Odd- and even-page section breaks (parityPage): page 2 left blank, the second section on page 3, the third on 4.
await open('odd-even-breaks.docx');
const oe = (await sheetPlaces()).filter((x) => x.text.trim()).map((x) => x.sheet + 1);
const oeSheets = await page.evaluate(() => document.querySelectorAll('.odx-sheets > .odx-sheet').length);
check(oe.join(',') === '1,3,4' && oeSheets === 4, `odd- and even-page section breaks: the sections on pages ${oe.join(', ')} of ${oeSheets}`);
// Printed, a note's reference mark keeps its number (printlines.js: CSS counts them, and in the print layer the count
// would start again): find-all.docx's footnote 1, after its comment's end (it had been left out of the printed line),
// and endnote i.
await open('find-all.docx');
await page.emulateMedia({ media: 'print' });
await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
const printedMarks = await page.evaluate(() => [...document.querySelectorAll('.odx-printlines [data-kind="footnoteReference"], .odx-printlines [data-kind="endnoteReference"]')].map((m) => `${m.dataset.kind}:${m.dataset.n}`));
await page.evaluate(() => window.dispatchEvent(new Event('afterprint')));
await page.emulateMedia({ media: 'screen' });
check(printedMarks.join(' ') === 'footnoteReference:1 endnoteReference:i', `printed note marks keep their numbers (${printedMarks.join(' ')})`);

await open('effects.docx');
// (Two lines of 11pt Calibri at 1.08 are 38.7px; the marks add 0.31em, 4.5px, above the first or below it.)
const grow = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].slice(6, 10).map((p) => p.getBoundingClientRect().height));
check(grow.every((h) => h > 42.5 && h < 45), `emphasis marks make their line taller (${grow.map((h) => h.toFixed(1)).join(', ')}px)`);
// A reflection (w14:reflection, 36pt Arial bold): each word drawn again below it, flipped about its baseline and
// faded (as Word's print of effects-word/effects14b.docx), the copy as tall as the font's ascent (0.905em, 43.4px)
// and its foot on the word's baseline; not text: not in the paragraph's text, a selection or what a screen reader reads.
const refl = await page.evaluate(() => {
  const p = [...document.querySelectorAll('.odx-content > p')].find((x) => x.textContent.startsWith('Reflected'));
  const words = [...p.querySelectorAll('.odx-rf')];
  const a = getComputedStyle(words[0], '::after');
  const sel = getSelection();
  sel.selectAllChildren(p);
  const copied = sel.toString();
  sel.removeAllRanges();
  // (The word's baseline: its font's ascent below the top of its box, where the copy is placed.)
  const ctx = document.createElement('canvas').getContext('2d');
  const cs = getComputedStyle(words[0]);
  ctx.font = `${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
  const base = ctx.measureText('R').fontBoundingBoxAscent;
  return { words: words.map((w) => w.dataset.odxRf), content: a.content, transform: a.transform, mask: a.maskImage || a.webkitMaskImage, height: parseFloat(a.height), base, text: p.textContent, copied };
});
const reflAria = await page.locator('.odx-content > p', { hasText: 'Reflected' }).ariaSnapshot();
check(refl.words.join(',') === 'Reflected,words' && refl.content === '"Reflected" / ""' && /matrix\(1, 0, 0, -1, 0, 0\)/.test(refl.transform) && /linear-gradient/.test(refl.mask)
  && Math.abs(refl.height - 43.4) < 1.5 && Math.abs(refl.base - refl.height) < 1 && refl.text === 'Reflected words and plain after' && !/Reflected.*Reflected/.test(refl.copied) && (reflAria.match(/Reflected/g) || []).length === 1,
  `a reflection: each word mirrored below its baseline, faded, not text (${JSON.stringify({ ...refl, mask: undefined, aria: reflAria })})`);
// Its copy keeps the descenders (room for them below the baseline, the flip about the baseline still), and is softened
// by its blur (effects.docx's 0.5pt), as Word's print of wordart-leftovers/refl.docx.
const reflMore = await page.evaluate(() => {
  const w = document.querySelector('.odx-rf');
  const a = getComputedStyle(w, '::after');
  return { pad: parseFloat(a.paddingBottom), size: parseFloat(a.fontSize), origin: a.transformOrigin, height: parseFloat(a.height), filter: a.filter };
});
check(Math.abs(reflMore.pad - reflMore.size * 0.25) < 0.5 && Math.abs(parseFloat(reflMore.origin.split(' ')[1]) - reflMore.height) < 1 && /^blur\(0\.\d+px\)$/.test(reflMore.filter),
  `a reflection keeps its descenders and is blurred (${JSON.stringify(reflMore)})`);
// An inner shadow (a w14:shadow scaled to nothing: the gallery's Style 15, 36pt Arial bold): each word drawn again over
// itself through the page's SVG filter for it, which keeps only the shadow inside the letters (as Word's print of
// odx-test/mkinnershadow.ps1); not text, and no room: the line as tall as the plain 36pt line after it would be.
const pressed = await page.evaluate(() => {
  const p = [...document.querySelectorAll('.odx-content > p')].find((x) => x.textContent.startsWith('Pressed'));
  const words = [...p.querySelectorAll('.odx-is')];
  const b = getComputedStyle(words[0], '::before');
  const id = /url\("?#([\w-]+)"?\)/.exec(b.filter)?.[1];
  const f = id && document.getElementById(id);
  const sel = getSelection();
  sel.selectAllChildren(p);
  const copied = sel.toString();
  sel.removeAllRanges();
  const plain = [...document.querySelectorAll('.odx-content > p')].find((x) => x.textContent.startsWith('Reflected'));
  return { words: words.map((w) => w.dataset.odxIs), content: b.content, position: b.position, fill: b.webkitTextFillColor, filter: f?.tagName, offset: f?.querySelector('feOffset') && [f.querySelector('feOffset').getAttribute('dx'), f.querySelector('feOffset').getAttribute('dy')], height: p.getBoundingClientRect().height, text: p.textContent, copied, reflectedHeight: plain.getBoundingClientRect().height };
});
const pressedAria = await page.locator('.odx-content > p', { hasText: 'Pressed' }).ariaSnapshot();
// (Moved against the shadow's direction, a third of its 4pt distance: 1.33pt up and left is 1.26px each way.)
check(pressed.words.join(',') === 'Pressed,words' && pressed.content === '"Pressed" / ""' && pressed.position === 'absolute' && pressed.fill === 'rgb(0, 0, 0)' && pressed.filter === 'filter'
  && pressed.offset?.join() === '1.26,1.26' && pressed.text === 'Pressed words and plain after' && !/Pressed.*Pressed/.test(pressed.copied) && (pressedAria.match(/Pressed/g) || []).length === 1
  && pressed.height < pressed.reflectedHeight,
  `an inner shadow: each word drawn again through its filter, not text, no room (${JSON.stringify({ ...pressed, aria: pressedAria })})`);

await open('notes-long.docx');
// (Four notes of 45 paragraphs each, from one paragraph: they continue at the foot of the pages after it, as in Word,
// each page's notes inside its own margins, never over its body text.)
const notes = await page.evaluate(() => [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => {
  const n = s.querySelector('.odx-notes');
  return n ? n.getBoundingClientRect().top - s.getBoundingClientRect().top : null;
}));
check(notes.length >= 5 && notes.every((t) => t != null && t >= 96 + 20), `long footnotes continue over ${notes.length} pages, below each page's text (${notes.map((t) => Math.round(t)).join(',')})`);

// An input method (Chinese pinyin) composing with Track Changes on: the composed text is one tracked insertion, and
// nothing of the letters typed while composing is left behind.
await page.evaluate(() => { const s = document.querySelector('.odx-tracksel'); s.value = 'on'; s.dispatchEvent(new Event('change', { bubbles: true })); });
await page.click('.odx-content > p');
await page.evaluate(() => { const r = document.createRange(); r.selectNodeContents(document.querySelector('.odx-content > p')); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); });
const cdp = await page.context().newCDPSession(page);
for (const s of ['z', 'zh', 'zhong']) { await cdp.send('Input.imeSetComposition', { text: s, selectionStart: s.length, selectionEnd: s.length }); await page.waitForTimeout(60); }
await cdp.send('Input.insertText', { text: '中文' });
await page.waitForTimeout(400);
const ime = await page.evaluate(() => ({ text: document.querySelector('.odx-content > p').textContent, ins: [...document.querySelectorAll('.odx-content ins.odx-ins')].map((e) => e.textContent).join('|') }));
check(ime.text.endsWith('中文') && !/z/.test(ime.text.slice(-8)) && ime.ins === '中文', `IME input with Track Changes on is one tracked insertion (${JSON.stringify(ime)})`);
await page.evaluate(() => { const s = document.querySelector('.odx-tracksel'); s.value = 'off'; s.dispatchEvent(new Event('change', { bubbles: true })); });

await open('rtl-hebrew.docx');
const rtl = await page.evaluate(() => {
  const p = document.querySelectorAll('.odx-content > p')[1];
  const r = document.createRange(); r.selectNodeContents(p);
  const c = document.querySelector('.odx-content');
  return [getComputedStyle(p).direction, c.getBoundingClientRect().right - parseFloat(getComputedStyle(c).paddingRight) - Math.max(...[...r.getClientRects()].map((q) => q.right))];
});
check(rtl[0] === 'rtl' && rtl[1] < 2, `a right-to-left paragraph starts at the right margin (${rtl[1].toFixed(1)}px)`);

await open('fields.docx');
// A running head (STYLEREF) worked out for each page, as Word does; the body's fields keep their saved results until
// they're updated (F9), then match Word's (its PDF of the document with its fields updated).
const heads = await page.evaluate(() => [...document.querySelectorAll('.odx-sheet .odx-header')].map((h) => h.textContent.replace(/\s+/g, ' ').trim()));
check(heads.join(' | ') === 'Report — Introduction — page i | Report — Introduction — page 1 | Report — Results — page 2 | Report — Results — page 3', `running heads per page: ${heads.join(' | ')}`);
const bodyText = () => page.evaluate(() => window.odx.view.state.doc.textContent);
check((await bodyText()).includes('chapter 9, "Old title", on page 99'), 'the body keeps its saved field results until they are updated');
const updated = await page.evaluate(() => window.odx.updateFields());
const after = await bodyText();
check(updated >= 5 && after.includes('chapter 2, "Methods of measurement", on page 1') && after.includes('results are on page 2, on page 2, and the introduction on page 1') && after.includes('Error! Reference source not found.'), `Update Fields gives Word's values (${updated} updated)`);
await page.click('.odx-content p >> nth=1');
await page.keyboard.press('Control+z');
check((await bodyText()).includes('on page 99'), 'one Ctrl+Z puts the saved results back');
// An agent working in the editor: its edit appears as a tracked change under its name; the person accepts it.
const agentRun = await page.evaluate(() => {
  const a = window.odx.agent({ author: 'Assistant' });
  const head = a.outline()[2];
  const n = a.replace('Results', 'Findings', { id: head.id });
  const del = document.querySelector('.odx-content del.odx-del[data-author="Assistant"]')?.textContent;
  const ins = document.querySelector('.odx-content ins.odx-ins[data-author="Assistant"]')?.textContent;
  a.resolveChanges({ accept: true });
  return { n, del, ins, after: a.outline()[2].text, left: a.changes().length };
});
check(agentRun.n === 1 && agentRun.del === 'Results' && agentRun.ins === 'Findings' && agentRun.after === 'Findings' && agentRun.left === 0, `an agent's edit shows as a tracked change under its name, then accepted (${JSON.stringify(agentRun)})`);

// STYLEREF by heading level (mkstyleref.mjs): a caption's chapter number, as in Word's PDF of it.
await open('styleref-level.docx');
await page.evaluate(() => window.odx.updateFields());
{
  const text = await bodyText();
  check(text.includes('Table 3-1.') && text.includes('Plain: Place of Performance | heading 1 by name: Place of Performance | number by name: 3'),
    `STYLEREF 1 \\s and STYLEREF 1: Heading 1's number and text, as in Word (${text.slice(-160)})`);
}

await open('tiff.docx');
const tiffs = await page.evaluate(() => [...document.querySelectorAll('.odx-content img:not(.ProseMirror-separator)')].map((i) => i.complete && i.naturalWidth > 0));
check(tiffs.length === 3 && tiffs.every(Boolean), `TIFF pictures (uncompressed, LZW, fax) are shown (${tiffs.length})`);

// Restrict Editing, as Word enforces it.
const typeAt = async (text, keys) => { await page.click(`.odx-content p:has-text("${text}")`); await page.keyboard.press('End'); await page.keyboard.type(keys); await page.waitForTimeout(150); };
const note = () => page.evaluate(() => document.querySelector('.odx-protectnote span')?.textContent || '');
await open('protect-readonly.docx');
check(/can only be read/.test(await note()), 'a document protected read-only says so above the ribbon');
await typeAt('This paragraph is locked.', 'XX');
check(!(await bodyText()).includes('XX'), 'read-only: typing in a locked paragraph does nothing');
const perms = await page.evaluate(() => [...document.querySelectorAll('.odx-content .odx-perm')].map((e) => e.textContent).join('|'));
check(perms === 'Everyone may edit this.', `the part everyone may edit is highlighted, not the tester's (${perms})`);
// No keyboard trap: in a locked part, Tab moves the focus on instead of typing (WCAG 2.1.2).
await page.click('.odx-content p:has-text("This paragraph is locked.")');
await page.keyboard.press('Tab');
check(await page.evaluate(() => !document.activeElement.classList.contains('odx-content')), 'read-only: Tab leaves the document');
await page.click('.odx-content .odx-perm');
await page.keyboard.type('YY');
await page.waitForTimeout(150);
check((await bodyText()).includes('YY'), 'read-only: the part everyone may edit can be typed in');
check(await page.evaluate(() => window.odx.view.odx.session.comments.length === 0 && (document.querySelector('.odx-content p').click(), true)), 'no comments yet');
await page.keyboard.press('Control+Alt+m');
check(await page.evaluate(() => window.odx.view.odx.session.comments.length === 0), 'read-only: no comment can be added');
// Comments only, with a password; then stopped with it.
check(await page.evaluate(async () => { await window.odx.unprotect(); await window.odx.protect({ edit: 'comments', password: 'pw' }); return window.odx.protection.edit === 'comments' && window.odx.protection.password; }), 'Restrict Editing: protected for comments with a password');
await typeAt('This paragraph is locked too.', 'ZZ');
await page.keyboard.press('Control+Alt+m');
await page.waitForTimeout(200);
check(!(await bodyText()).includes('ZZ') && (await page.evaluate(() => window.odx.view.odx.session.comments.length)) === 1, 'comments only: no typing, but a comment is added');
check(await page.evaluate(async () => !(await window.odx.unprotect('wrong')) && window.odx.protection.enforced && (await window.odx.unprotect('pw')) && !window.odx.protection.enforced), 'Stop Protection needs the password');
// Tracked changes only: every edit tracked, Track Changes can't be turned off, nothing accepted.
await page.evaluate(() => window.odx.protect({ edit: 'trackedChanges' }));
await typeAt('This paragraph is locked.', 'TT');
const tracked = await page.evaluate(() => [...document.querySelectorAll('.odx-content ins.odx-ins')].map((e) => e.textContent).join(''));
await page.evaluate(() => { const t = document.querySelector('.odx-tracksel'); t.value = 'off'; t.dispatchEvent(new Event('change')); });
await page.waitForTimeout(150);
const stillOn = await page.evaluate(() => document.querySelector('.odx-tracksel').value);
await page.evaluate(() => window.odx.view.focus());
check(tracked.includes('TT') && stillOn === 'on', `tracked changes only: typing is tracked and Track Changes stays on (${tracked}, ${stillOn})`);
await page.evaluate(() => window.odx.unprotect());
// A protected form with an unprotected section.
await open('protect-sections.docx');
await typeAt('can be edited freely.', ' Yes.');
await typeAt('This text of the form is locked.', 'QQ');
// (A click on the form's own text puts the caret in its next field, as in Word: the typing goes there.)
const formText = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].map((p) => p.textContent));
check(formText[0].endsWith('freely. Yes.') && formText[3] === 'This text of the form is locked.', `a protected form: its unprotected section is edited, the form's own text is not (${formText.join(' / ')})`);
// Formatting limited to the allowed styles.
await open('protect-format.docx');
const styleNames = await page.evaluate(() => [...document.querySelectorAll('.odx-stile')].map((t) => t.title));
check(styleNames.includes('Heading 1') && !styleNames.includes('Heading 2'), 'limited formatting: a locked style is not offered');
await page.click('.odx-content p:has-text("Some text")');
await page.keyboard.press('Control+a');
await page.keyboard.press('Control+b');
await page.waitForTimeout(150);
await page.keyboard.press('Control+Home');
check(await page.evaluate(() => !document.querySelector('.odx-content strong, .odx-content b')), 'limited formatting: no direct formatting (Ctrl+B does nothing)');
await typeAt('Some text', ' More.');
check((await bodyText()).includes('formatting. More.'), 'limited formatting: the text can still be edited');
// Mark as Final: read-only until Edit Anyway, which takes the mark away.
await open('protect-final.docx');
await typeAt('marked as final', 'FF');
const finalNote = await note();
await page.click('.odx-protectnote button');
await typeAt('marked as final', 'GG');
check(/marked this document as final/.test(finalNote) && !(await bodyText()).includes('FF') && (await bodyText()).includes('GG'), 'Mark as Final: read-only, then Edit Anyway');

// Table of contents lines (mktoc.mjs): the page number sits right-aligned at the margin on its tab's line, also when
// the title leaves no room for it, where the tab starts the next line at the indent, as in Word's PDF of the file.
await open('toc-wrap.docx');
const toc = await page.evaluate(() => [...document.querySelectorAll('.odx-content p')].filter((p) => p.querySelectorAll('.odx-tab').length === 2).map((p) => {
  const tab = p.querySelectorAll('.odx-tab')[1].getBoundingClientRect(), pr = p.getBoundingClientRect();
  const num = [...p.querySelectorAll('.odx-t')].pop().getBoundingClientRect();
  return { sameLine: num.bottom > tab.top && num.top < tab.bottom, atRight: Math.abs(num.right - pr.right) < 2, startsLine: Math.abs(tab.left - pr.left) < 1, at: [p.textContent.slice(0, 24), Math.round(num.top), Math.round(tab.top), Math.round(num.right), Math.round(pr.right), getComputedStyle(p).fontFamily.slice(0, 30), p.querySelectorAll('.odx-tab')[1].style.cssText, [...p.querySelectorAll('.odx-t')].map((t) => `${t.textContent.slice(0, 12)}@${Math.round(t.getBoundingClientRect().left)}-${Math.round(t.getBoundingClientRect().right)}`).join(' '), (() => { const r = document.createRange(); r.setStartAfter(p.querySelectorAll('.odx-tab')[1]); r.setEnd(p, p.childNodes.length); return [...r.getClientRects()].map((x) => `${Math.round(x.left)}+${x.width.toFixed(1)}@${Math.round(x.top)}`).join(' '); })(), [...p.childNodes].slice(-4).map((n) => n.nodeName + (n.className ? `.${n.className}` : '')).join(' '),
    // (What the tab's width would be from the page as it is now: right if it was measured too early.)
    (() => { const cs = getComputedStyle(p), t = p.querySelectorAll('.odx-tab')[1]; const edge = pr.left - parseFloat(cs.marginLeft); const x = t.getBoundingClientRect().left - edge; const right = pr.right - parseFloat(cs.paddingRight) - edge; return { x: +x.toFixed(2), right: +right.toFixed(2), stop: +(9350 / 15).toFixed(2), rw: +num.width.toFixed(2), now: +(Math.min(9350 / 15, right) - x - num.width - 0.5).toFixed(2), dpr: devicePixelRatio, zoom: getComputedStyle(document.querySelector('.odx-sheet') || p).transform }; })()] };
}));
const tocBad = toc.find((x) => !(x.sameLine && x.atRight));
check(toc.length === 30 && !tocBad, `TOC: each page number right-aligned on its tab's line (${toc.filter((x) => x.sameLine && x.atRight).length}/${toc.length}${tocBad ? `; first not: ${JSON.stringify(tocBad)}` : ''})`);
check(toc.some((x) => x.startsLine), 'TOC: a title with no room left takes its tab and number to the next line');

// Find goes through every part of the document as Word's does (mkfindall.mjs): the main text, a text box, the header
// and footer, a footnote and an endnote, opening each part with its match selected; Replace All changes them all.
await open('find-all.docx');
await page.click('.odx-content p');
await page.keyboard.press('Control+Home');
await page.keyboard.press('Control+f');
await page.fill('.odx-findbar input[aria-label="Find"]', 'apple');
const visits = [];
const findSaid = []; // (what the find bar says at each, for screen readers: which match, and where)
for (let i = 0; i < 9; i++) {
  await page.press('.odx-findbar input[aria-label="Find"]', 'Enter');
  await page.waitForTimeout(250);
  findSaid.push(await page.textContent('.odx-findinfo'));
  findSaid[findSaid.length - 1] += await page.evaluate(() => (document.activeElement?.getAttribute('aria-label') === 'Find' ? '' : ' (focus left the find box)'));
  visits.push(await page.evaluate(() => {
    const ta = document.querySelector('.odx-markup textarea[data-odx-box^="e"]');
    if (ta) return `comment:${ta.value.slice(ta.selectionStart, ta.selectionEnd)}`;
    const v = window.odx.view.odx.hfView || window.odx.view;
    const { from, to } = v.state.selection;
    const part = v.dom.closest('.odx-txbxedit') ? 'box' : v.dom.closest('.odx-hfedit-header') ? 'header' : v.dom.closest('.odx-hfedit-footer') ? 'footer'
      : v.dom.closest('.odx-noteedit') ? (v.state.doc.textContent.includes('endnote') ? 'endnote' : 'footnote') : 'body';
    return `${part}:${v.state.doc.textBetween(from, to)}`;
  }));
}

const info = await page.textContent('.odx-findinfo');
check(findSaid[3] === '4 of 8, in the header' && findSaid[5] === '6 of 8, in a footnote' && findSaid[7] === '8 of 8, in a comment', `the find bar says which match and where (${findSaid.join(' | ')})`);
check(visits.join(' ') === 'body:apple body:apple box:apple header:apple footer:apple footnote:apple endnote:apple comment:apple body:apple' && info === '1 of 8',
  `Find goes through the body, text box, header, footer, notes and comments (${visits.join(' ')}; ${info})`);
await page.keyboard.press('Control+h');
await page.fill('.odx-findbar input[aria-label="Replace with"]', 'pear');
const findFor = await page.inputValue('.odx-findbar input[aria-label="Find"]'); // (what Replace All looks for: said if it fails)
await page.click('.odx-findbar button:has-text("Replace all")');
await page.waitForFunction(() => /results?$/.test(document.querySelector('.odx-findinfo').textContent));
const replaced = await page.textContent('.odx-findinfo');
const saved = unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))));
const partText = (f) => strFromU8(saved[f]).replace(/<w:t\b[^>]*>([^<]*)<\/w:t>|<[^>]+>/g, '$1');
const parts = ['word/document.xml', 'word/header1.xml', 'word/footer1.xml', 'word/footnotes.xml', 'word/endnotes.xml', 'word/comments.xml'];
const pears = parts.map((f) => (partText(f).match(/pear/g) || []).length);
check(replaced === '8 results' && parts.every((f) => !partText(f).includes('apple')) && pears.join() === '4,1,1,1,1,1',
  `Replace All in every part, saved (${replaced} for "${findFor}"; pears per part ${pears.join()}, the text box's twice: its shape and fallback copy; apples left ${parts.map((f) => (partText(f).match(/apple/g) || []).length).join()})`);
await page.keyboard.press('Escape');
// The find bar closed with Escape, then Ctrl+H with the text box selected: the search typed before stays, as in Word.
// (The browser's Escape in a search box emptied it; and a picture has no text to search for, where taking its "" did
// too: Replace All then found nothing.)
await page.evaluate(() => {
  const v = window.odx.view;
  let at = null;
  v.state.doc.descendants((n, pos) => { if (at == null && n.type.name === 'image') at = pos; });
  v.focus();
  v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.create(v.state.doc, at, at + 1)));
});
await page.keyboard.press('Control+h');
const kept = await page.inputValue('.odx-findbar input[aria-label="Find"]');
check(kept === 'apple', `Escape, then Ctrl+H with a picture selected: the search stays (${JSON.stringify(kept)})`);
await page.keyboard.press('Escape');

// Columns of their own widths (mkcolumns.mjs: 1.5in, 3in and 1.25in), laid out as Word's PDF of the file has them: each
// column at its place and width, a paragraph going on in a wider column laid out again at that width (P9's last
// lines: two at 3in), P24 starting the third column and P30 ending the first page.
await open('columns-unequal.docx');
const ucols = await page.evaluate(() => {
  const sheet = document.querySelector('.odx-sheet').getBoundingClientRect();
  const at = (n) => { const p = [...document.querySelectorAll('.odx-content > p')].find((x) => x.textContent.startsWith(`P${n} `)); const r = p.getBoundingClientRect(); return { x: Math.round(r.left - sheet.left), w: Math.round(r.width), top: Math.round(r.top - sheet.top) }; };
  const cont = document.querySelector('.odx-repeat .odx-contlines');
  return { p1: at(1), p10: at(10), p24: at(24), p30: at(30), p31: at(31), cont: cont && { w: Math.round(cont.getBoundingClientRect().width), lines: Math.round(cont.getBoundingClientRect().height / 19.33), text: cont.textContent.slice(0, 15) }, pages: window.odx.view.odx.pages.count };
});
check(ucols.p1.x === 96 && ucols.p1.w === 144 && ucols.p10.x === 264 && ucols.p10.w === 288 && ucols.p24.x === 600 && ucols.p24.w === 120,
  `columns of their own widths: each at its place and width (${JSON.stringify([ucols.p1, ucols.p10, ucols.p24])})`);
check(ucols.cont?.w === 288 && ucols.cont.lines === 2 && ucols.cont.text.startsWith('procedures'), `a paragraph going on in a wider column is laid out again at its width (${JSON.stringify(ucols.cont)})`);
check(ucols.pages === 3 && ucols.p30.top < 1056 && ucols.p31.top > 1056, `columns of their own widths: P30 ends the first page, as in Word (${ucols.pages} pages)`);

// Paragraphs in frames (mkframes.mjs) where Word's PDF of the file has them: frame A's text 2in down the page and 4.5in
// across, its two paragraphs one frame; sidebar B at the right of the text, Body 6 beside it; note C in the left margin.
await open('frames.docx');
const fr = await page.evaluate(() => {
  const sheet = document.querySelector('.odx-sheet').getBoundingClientRect();
  const para = (t) => [...document.querySelectorAll('.odx-content > p')].find((x) => x.textContent.startsWith(t));
  const text = (t) => { const r = document.createRange(); r.selectNodeContents(para(t)); const b = r.getClientRects()[0]; return { x: Math.round(b.left - sheet.left), y: Math.round(b.top - sheet.top) }; };
  const box = (t) => { const b = para(t).getBoundingClientRect(); return { left: Math.round(b.left - sheet.left), right: Math.round(b.right - sheet.left), top: Math.round(b.top - sheet.top) }; };
  const r = document.createRange(); r.selectNodeContents(para('Body 6.')); const body6 = Math.max(...[...r.getClientRects()].filter((x) => x.top < para('Frame B').getBoundingClientRect().bottom).map((x) => x.right)) - sheet.left;
  return { a1: text('Frame A, first'), a2: box('Frame A, second'), a1box: box('Frame A, first'), b: box('Frame B'), c: box('Note C'), body6: Math.round(body6) };
});
check(Math.abs(fr.a1.x - 432) <= 2 && fr.a1box.top === 192 && fr.a2.left === fr.a1box.left && fr.a2.top > fr.a1box.top,
  `a frame set down on the page: its text 4.5in across, the frame 2in down, its second paragraph below (${JSON.stringify([fr.a1, fr.a1box, fr.a2])})`);
check(fr.b.right === 720 && fr.body6 < fr.b.left && fr.c.left === 12 && fr.c.right <= 96,
  `a frame at the right of the text with the text beside it, and one in the margin (${JSON.stringify([fr.b, fr.body6, fr.c])})`);

// The Original view (mkoriginal.mjs): the document before its tracked changes, as Word's Review > Original shows it.
await open('original.docx');
const shown = () => page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].filter((p) => getComputedStyle(p).display !== 'none').map((p) => p.innerText.replace(/\s+/g, ' ').trim()).filter(Boolean));
await page.click('.odx-rtab:text-is("Review")');
await page.selectOption('.odx-markupsel', 'original');
await page.waitForTimeout(400);
const orig = await shown();
await page.selectOption('.odx-markupsel', 'none');
await page.waitForTimeout(400);
const final = await shown();
await page.selectOption('.odx-markupsel', 'simple');
await page.click('.odx-rtab:text-is("Home")');
check(orig.join(' | ') === 'Alpha words old end. | A paragraph deleted whole. | Split here and joined again. | First of two | Second of two.',
  `Original: insertions out, deletions in, paragraphs as they were (${orig.join(' | ')})`);
check(final.join(' | ') === 'Alpha new words end. | A paragraph inserted whole. | Split here by Enter | and joined again. | First of twoSecond of two.',
  `No Markup still shows the changed document (${final.join(' | ')})`);

// Tracked changes and hidden paragraph marks as Word's default PDF export gives them (mkrevprint.mjs: its PDF of the
// same file): a hidden mark's paragraph runs on at the next one's indent, a style separator's at its own; a list item
// whose mark is deleted isn't counted, one whose text is all deleted keeps its number; a mark inserted then deleted is
// gone. On screen the change bar and the formatting tint stay; on paper neither.
await open('revprint.docx');
const rp = await page.evaluate(() => {
  const ps = [...document.querySelectorAll('.odx-content > p')].filter((p) => getComputedStyle(p).display !== 'none');
  const find = (t) => ps.find((p) => p.innerText.includes(t));
  const x = (t) => { const p = find(t); const r = document.createRange(); r.selectNodeContents(p); return Math.round(r.getClientRects()[0].left - find('The end.').getBoundingClientRect().left); };
  const changed = document.querySelector('.odx-content p.odx-changed');
  return {
    texts: ps.map((p) => p.innerText.replace(/\s+/g, ' ').trim()),
    labels: ps.filter((p) => p.classList.contains('odx-li')).map((p) => p.dataset.label),
    hidden: x('Runs on at the margin.'), sep: x('Runs on at 2in.'),
    bar: changed && getComputedStyle(changed).boxShadow, tint: getComputedStyle(document.querySelector('.odx-content .odx-fmt')).backgroundColor,
  };
});
await page.emulateMedia({ media: 'print' });
const rpPrint = await page.evaluate(() => ({ bar: getComputedStyle(document.querySelector('.odx-content p.odx-changed')).boxShadow, tint: getComputedStyle(document.querySelector('.odx-content .odx-fmt')).backgroundColor }));
await page.emulateMedia({ media: 'screen' });
check(rp.texts.some((t) => t === 'Hidden mark A. Runs on at the margin.') && rp.hidden === 0 && rp.sep === 192,
  `a hidden paragraph mark runs on at the next paragraph's indent, a style separator at its own 2in (${rp.hidden}, ${rp.sep})`);
check(rp.labels.join(' ') === '1. 2. 3. 1. 2. 3.' && rp.texts.some((t) => t.replace(/\s/g, '').includes('Itemtwo.Itemthree.')) && !rp.texts.some((t) => t.includes('Inserted and deleted')),
  `list items: one whose mark is deleted not counted, one whose text is deleted numbered, a mark inserted then deleted gone (${rp.labels.join(' ')})`);
check(rp.bar && rp.bar !== 'none' && rp.tint !== 'rgba(0, 0, 0, 0)' && rpPrint.bar === 'none' && rpPrint.tint === 'rgba(0, 0, 0, 0)',
  `change bar and formatting tint on screen, neither on paper (${JSON.stringify([rp.bar, rp.tint, rpPrint])})`);

// Wrap Text (Layout > Wrap Text): the selected picture floats with the text square round it, goes behind the text, and
// back in line; each drawn so and saved as Word writes it.
await open('rich.docx');
await page.click('img[alt="A blue box"]'); // (selected; Wrap Text keeps it selected)
const wrapTo = async (name) => {
  await page.click('.odx-rtab:text-is("Layout")');
  await page.click('button[title^="Wrap Text"]');
  await page.click(`.odx-item:has-text("${name}")`);
  await page.waitForTimeout(400);
  const look = await page.evaluate(() => { const cs = getComputedStyle(document.querySelector('img[alt="A blue box"]')); return `${cs.float}/${cs.position}/${cs.zIndex}`; });
  const saved = strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/document.xml']);
  const at = saved.indexOf('name="blue box"');
  const open = Math.max(saved.lastIndexOf('<wp:inline', at), saved.lastIndexOf('<wp:anchor', at));
  const head = open < 0 ? '' : saved.slice(open, at); // (the drawing's start up to its name: its wrap element comes before)
  const wrapEl = /<wp:(wrap(?!Polygon)\w+)/.exec(head)?.[1];
  const saidSo = [head.startsWith('<wp:anchor') ? 'anchor' : 'inline', wrapEl, /behindDoc="1"/.test(head) && 'behind'].filter(Boolean).join(' ');
  return { look, saved: head ? saidSo : 'missing' };
};
const sq = await wrapTo('Square');
const behind = await wrapTo('Behind Text');
const inl = await wrapTo('In Line with Text');
await page.click('.odx-rtab:text-is("Home")');
check(sq.look.startsWith('left/') && sq.saved === 'anchor wrapSquare', `Wrap Text Square: floated, saved square (${JSON.stringify(sq)})`);
check(behind.look === 'none/absolute/-1' && behind.saved === 'anchor wrapNone behind', `Wrap Text Behind Text: behind the text, saved so (${JSON.stringify(behind)})`);
check(inl.look.startsWith('none/static') && inl.saved === 'inline', `Wrap Text In Line with Text: back in line (${JSON.stringify(inl)})`);

// Layout > Page Setup (mkpagesetup.mjs, two sections): margins, orientation and paper for the section the cursor is in,
// then the whole document, saved in each section's w:sectPr.
await open('page-setup.docx');
const sectsSaved = async () => {
  const xml = strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/document.xml']);
  return [...xml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((m) => {
    const a = (tag, k) => new RegExp(`<w:${tag}\\b[^>]*\\sw:${k}="(\\d+)"`).exec(m[0])?.[1];
    return `${a('pgSz', 'w')}x${a('pgSz', 'h')}${/w:orient="landscape"/.test(m[0]) ? 'L' : ''} t${a('pgMar', 'top')} l${a('pgMar', 'left')}`;
  });
};
const setup = async (vals, all) => {
  await page.click('.odx-content p');
  await page.keyboard.press('Control+Home');
  await page.click('.odx-rtab:text-is("Layout")');
  await page.click('button[title^="Page Setup"]');
  for (const [label, v] of Object.entries(vals)) {
    const sel = `.odx-pagesetup [aria-label="${label}"]`;
    if ((await page.$eval(sel, (e) => e.tagName)) === 'SELECT') await page.selectOption(sel, v); else await page.fill(sel, v);
  }
  await page.selectOption('.odx-pagesetup [aria-label="Apply to"]', all ? 'all' : 'section');
  await page.click('.odx-pagesetup button:has-text("OK")');
  await page.waitForFunction(() => window.odx?.fileName === 'page-setup.docx');
  await page.waitForTimeout(800);
};
await setup({ 'Left (inches)': '1.5', Orientation: 'landscape', 'Paper size': 'A4' }, false);
const one = await sectsSaved();
await setup({ 'Top (inches)': '0.5' }, true);
const both = await sectsSaved();
await page.click('.odx-rtab:text-is("Home")');
check(one.join(' / ') === '16838x11906L t1440 l2160 / 12240x15840 t1440 l1440', `Page Setup, this section: A4 landscape, left 1.5in, the other section as it was (${one.join(' / ')})`);
check(both.join(' / ') === '16838x11906L t720 l2160 / 16838x11906L t720 l2160', `Page Setup, whole document: the dialog's page in both sections (${both.join(' / ')})`);

// Imported content (w:altChunk, mkaltchunk.mjs): drawn as the paragraphs Word makes of it, in the cells and the body,
// with no placeholder left; Find finds its text.
await open('altchunk.docx');
const chunk = await page.evaluate(() => ({
  opaque: document.querySelectorAll('.odx-content .odx-opaque').length,
  cell: [...document.querySelectorAll('.odx-content td')].map((td) => td.textContent).find((t) => t.startsWith('Basis')),
  bullets: [...document.querySelectorAll('.odx-content p.odx-li')].map((p) => p.dataset.label + p.textContent).join('|'),
}));
check(chunk.opaque === 0 && chunk.cell === 'Basis of Estimate: hours from the last three task orders.', `imported content drawn as text, none as a placeholder (${JSON.stringify(chunk.cell)}, ${chunk.opaque} left)`);
check(chunk.bullets.includes('First bullet') && chunk.bullets.includes('1.First step'), `its lists are lists (${chunk.bullets})`);
await page.click('.odx-content p');
await page.keyboard.press('Control+f');
await page.fill('.odx-findbar input[aria-label="Find"]', 'task orders');
await page.press('.odx-findbar input[aria-label="Find"]', 'Enter');
await page.waitForTimeout(300);
check((await page.textContent('.odx-findinfo')) === '1 of 1', 'Find finds imported text');
await page.keyboard.press('Escape');

// Tables after page breaks stay on their pages when clicks move the caret between them (mkpagedtables.mjs): the
// table's own placement style must not wipe out its page push.
await open('paged-tables.docx');
const tableTops = () => page.evaluate(() => [...document.querySelectorAll('.odx-content > .tableWrapper')].map((t) => Math.round(t.getBoundingClientRect().top + document.querySelector('.odx-canvas').scrollTop)).join(' '));
const tops0 = await tableTops();
for (const at of ['A. SCOPE', 'B. MILESTONES', 'A. SCOPE', 'C. APPROACH']) {
  await page.click(`.odx-content td:has-text("${at}")`);
  await page.waitForTimeout(300);
}
const tops1 = await tableTops();
check(tops0 === tops1 && new Set(tops0.split(' ')).size === 3, `tables after page breaks keep their places after clicks (${tops0} -> ${tops1})`);

// List numbers in their level's bold, italic and colour, else the paragraph mark's (mklists2.mjs; as in Word's PDF).
await open('lists-more.docx');
const lbls = await page.evaluate(() => [...document.querySelectorAll('.odx-content p.odx-li')].slice(0, 6).map((p) => { const b = getComputedStyle(p, '::before'); return `${b.fontWeight >= 700 ? 'b' : ''}${b.fontStyle === 'italic' ? 'i' : ''}${b.color === 'rgb(192, 0, 0)' ? 'red' : ''}`; }));
check(lbls.join(',') === 'bred,bred,i,i,b,', `list numbers take their level's (or the paragraph mark's) bold, italic and colour (${lbls.join(',')})`);

// Auto spacing between items of one list, each side on its own; two nums of one abstract list as one list; auto spacing
// turned off outright, or kept under a plain w:after; an empty paragraph in its mark's theme font (mklists3.mjs; Word's
// PDF: 18.7, 26.6, 18.6, 28.0, 37.4px, 36.7px over two bullets as over two plain lines, and 38.0 and 36.7px over and under the empty lines).
await open('lists-spacing.docx');
const lsTops = await page.evaluate(() => {
  const ps = [...document.querySelectorAll('.odx-content > p')];
  const top = (t) => { const p = ps.find((x) => x.textContent === t); return p ? p.getBoundingClientRect().top : NaN; };
  return Object.fromEntries(['auto-a1', 'auto-a2', 'auto-a3', 'auto-a4', 'auto-b1', 'auto-b2', 'off-1', 'off-2', 'off-4', 'off-5', 'small-1', 'small-2', 'small-3', 'small-4', 'small-5', 'cour-1', 'cour-3', 'cour-5', 'msz-1', 'msz-3', 'msz-5', 'mark-1', 'mark-2', 'mark-3'].map((t) => [t, top(t)]));
});
const lsGap = (a, b) => +(lsTops[b] - lsTops[a]).toFixed(1);
check(lsGap('auto-a1', 'auto-a2') < 21 && lsGap('auto-a3', 'auto-a4') > 24 && lsGap('auto-a3', 'auto-a4') < 30 && lsGap('auto-b1', 'auto-b2') < 21,
  `auto spacing between items of one list goes, each side on its own (${lsGap('auto-a1', 'auto-a2')}, ${lsGap('auto-a3', 'auto-a4')}, ${lsGap('auto-b1', 'auto-b2')}px)`);
check(lsGap('off-1', 'off-2') > 25 && lsGap('off-1', 'off-2') < 31 && lsGap('off-4', 'off-5') > 34,
  `auto spacing turned off keeps the items' own space; a plain space after over it stays auto (${lsGap('off-1', 'off-2')}, ${lsGap('off-4', 'off-5')}px)`);
check(Math.abs(lsGap('small-1', 'small-3') - lsGap('small-3', 'small-5')) < 1,
  `a 10pt Symbol bullet over 12pt text adds nothing to its line (${lsGap('small-1', 'small-3')} over ${lsGap('small-3', 'small-5')}px)`);
check(Math.abs(lsGap('cour-1', 'cour-3') - lsGap('cour-3', 'cour-5')) < 1 && Math.abs(lsGap('msz-1', 'msz-3') - lsGap('msz-3', 'msz-5')) < 1,
  `a 10pt Courier New bullet over 12pt text, and a Symbol bullet sized by a 10pt mark over 11pt text, add nothing (${lsGap('cour-1', 'cour-3')} over ${lsGap('cour-3', 'cour-5')}, ${lsGap('msz-1', 'msz-3')} over ${lsGap('msz-3', 'msz-5')}px)`);
check(lsGap('mark-1', 'mark-2') - lsGap('mark-2', 'mark-3') > 0.6,
  `an empty paragraph is a line of its mark's theme font (${lsGap('mark-1', 'mark-2')} over ${lsGap('mark-2', 'mark-3')}px)`);

// Pasted from a web page or another editor: colour, size and font kept as marks; black (set on every run by some
// editors) and generic families left out.
await page.evaluate(() => window.odx.newDocument({ share: false }));
await page.click('.odx-content p');
const pasted = await page.evaluate(() => {
  const dt = new DataTransfer();
  dt.setData('text/html', '<p><span style="color:rgb(192, 0, 0); font-size:16pt; font-family:Georgia, serif">Red big Georgia</span> <span style="color:#000000; font-family:sans-serif; font-size:1.2em">plain</span></p>');
  dt.setData('text/plain', 'Red big Georgia plain');
  window.odx.view.dom.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  const marks = (word) => { let out = ''; window.odx.view.state.doc.descendants((n) => { if (n.isText && n.text.includes(word)) out = n.marks.map((m) => `${m.type.name}=${m.attrs.val}`).sort().join(' '); }); return out; };
  return [marks('Red'), marks('plain')];
});
check(pasted[0] === 'color=C00000 font=Georgia size=32' && pasted[1] === '', `pasted colour, size and font kept (${pasted.join(' | ')})`);

// Pasted headings written as styled <p>s (Word's clipboard HTML, other editors' classes, role=heading, a whole
// paragraph in bold at 2em): headings in the heading style they name; bold or MsoNormal text stays body text. Pasted
// into an empty paragraph, the first keeps its own style too, as in Word.
await page.evaluate(() => window.odx.newDocument({ share: false }));
await page.click('.odx-content p');
const pHeads = await page.evaluate(() => {
  const dt = new DataTransfer();
  dt.setData('text/html', '<p class="MsoHeading1">wordclass</p><p class="MsoNormal" style="mso-style-name:\'Heading 2\';margin:0">msoname</p>'
    + '<p class="MsoNormal" style="mso-outline-level:3">outline</p><p class="heading-2">editorclass</p><p style="font-size:2em;font-weight:bold">bigbold</p>'
    + '<p role="heading" aria-level="4">aria</p><p style="font-weight:bold">justbold</p><p class="MsoNormal">body</p>');
  dt.setData('text/plain', 'x');
  window.odx.view.dom.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  const out = [];
  window.odx.view.state.doc.forEach((n) => { if (n.textContent) out.push(`${n.textContent}=${n.attrs.styleId || ''}`); });
  return out.join(' ');
});
check(pHeads === 'wordclass=Heading1 msoname=Heading2 outline=Heading3 editorclass=Heading2 bigbold=Heading1 aria=Heading4 justbold= body=',
  `pasted <p> headings arrive as headings, the first in the empty paragraph it went into (${pHeads})`);
// Words from one paragraph (no paragraph mark of their own) into an empty heading: it stays a heading; one Undo takes
// away the paste and its style together.
const oneP = await page.evaluate(() => {
  const v = window.odx.view;
  const end = v.state.doc.content.size;
  const tr = v.state.tr.insert(end, v.state.schema.nodes.paragraph.create({ styleId: 'Heading2' }));
  v.dispatch(tr.setSelection(v.state.selection.constructor.create(tr.doc, end + 1)));
  const dt = new DataTransfer(); dt.setData('text/html', '<p>some words</p>'); dt.setData('text/plain', 'some words');
  v.dom.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  return v.state.doc.lastChild.textContent + '=' + v.state.doc.lastChild.attrs.styleId;
});
await page.keyboard.press('Control+Home');
await page.evaluate(() => {
  const v = window.odx.view;
  v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.create(v.state.doc, 1)));
});
await page.keyboard.press('Control+z'); // (the one-paragraph paste)
await page.keyboard.press('Control+z'); // (the inserted heading paragraph)
await page.keyboard.press('Control+z'); // (the first paste, style and all)
const undone = await page.evaluate(() => `${window.odx.view.state.doc.childCount}:${window.odx.view.state.doc.firstChild.attrs.styleId || ''}:${window.odx.view.state.doc.textContent}`);
check(oneP === 'some words=Heading2' && undone === '1::', `words from one paragraph keep the heading they go into; Undo takes a paste back whole (${oneP}, ${undone})`);
// A lone heading (a whole <h2>, Word's MsoHeading3) into an empty paragraph of body text: the heading's style, as in
// Word, not the paragraph's it went into.
const lone = await page.evaluate(() => {
  const v = window.odx.view;
  const out = [];
  for (const html of ['<h2>Lone heading</h2>', '<p class="MsoHeading3">Word heading</p>']) {
    const end = v.state.doc.content.size;
    const tr = v.state.tr.insert(end, v.state.schema.nodes.paragraph.create());
    v.dispatch(tr.setSelection(v.state.selection.constructor.create(tr.doc, end + 1)));
    const dt = new DataTransfer(); dt.setData('text/html', html); dt.setData('text/plain', 'x');
    v.dom.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
    out.push(`${v.state.doc.lastChild.textContent}=${v.state.doc.lastChild.attrs.styleId || ''}`);
  }
  return out.join(' ');
});
check(lone === 'Lone heading=Heading2 Word heading=Heading3', `a lone pasted heading takes its own heading style in an empty paragraph (${lone})`);

// Fit text (mkfittext.py): each fitted run, or group of runs, as wide as it says, the text after it starting there, as in
// Word's PDF (96, 192, 96, 72 and 144px on from the paragraph's left).
await open('fittext.docx');
const fits = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].slice(0, 5).map((p) => {
  const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  for (let t = w.nextNode(); t; t = w.nextNode()) {
    const i = t.data.indexOf('|');
    if (i < 0) continue;
    const r = document.createRange(); r.setStart(t, i); r.setEnd(t, i + 1);
    return Math.round(r.getBoundingClientRect().left - p.getBoundingClientRect().left);
  }
  return null;
}));
check(fits.join() === '96,192,96,72,144', `fitted text as wide as it says, as in Word (${fits.join()})`);

// Typing through an input method (Chinese, Japanese, Korean: composed, then committed) in each part of a document: the
// body, a text box, the header, the footer, a footnote and an endnote (find-all.docx, each reached with Find).
const typeIme = async (text) => {
  await cdp.send('Input.imeSetComposition', { text: text.slice(0, 1), selectionStart: 1, selectionEnd: 1 });
  await cdp.send('Input.imeSetComposition', { text, selectionStart: text.length, selectionEnd: text.length });
  await cdp.send('Input.insertText', { text });
  await page.waitForTimeout(150);
};
await open('find-all.docx');
await page.click('.odx-content p');
await page.keyboard.press('Control+Home');
await page.keyboard.press('Control+f');
await page.fill('.odx-findbar input[aria-label="Find"]', 'apple');
const imeParts = [];
for (let i = 0; i < 7; i++) {
  await page.press('.odx-findbar input[aria-label="Find"]', 'Enter');
  await page.waitForTimeout(250);
  if (i === 1) continue; // (the body's second apple)
  const v = await page.evaluate(() => { const v = window.odx.view.odx.hfView || window.odx.view; v.focus(); const { to } = v.state.selection; v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.near(v.state.doc.resolve(to)))); return true; });
  void v;
  await typeIme(`漢字${i}`);
  imeParts.push(await page.evaluate((i) => { const v = window.odx.view.odx.hfView || window.odx.view; return v.state.doc.textContent.includes(`漢字${i}`); }, i));
}
await page.keyboard.press('Escape');
const imeSaved = unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))));
const imeIn = ['word/document.xml', 'word/header1.xml', 'word/footer1.xml', 'word/footnotes.xml', 'word/endnotes.xml'].map((f) => (strFromU8(imeSaved[f]).match(/漢字\d/g) || []).length);
check(imeParts.every(Boolean) && imeIn.join() === '3,1,1,1,1', `input-method typing in the body, a text box, the header, the footer and notes, saved (${imeParts.join()}; ${imeIn.join()}: the text box's twice, its shape and fallback copy)`);

// Track Changes edge cases: a word deleted letter by letter, then typed over, reads in order (the new text before the
// deleted, where Word's caret is once it has gone back over it); undoing tracked typing
// takes its change away; Reject All in a table cell leaves no insertion behind.
await open('rich.docx');
await page.evaluate(() => window.odx.view.focus());
await page.click('.odx-content p >> nth=0');
await page.keyboard.press('End');
await page.keyboard.type(' alpha');
await page.evaluate(() => { const t = document.querySelector('.odx-tracksel'); t.value = 'on'; t.dispatchEvent(new Event('change')); });
await page.evaluate(() => window.odx.view.focus());
for (let i = 0; i < 5; i++) await page.keyboard.press('Backspace');
await page.keyboard.type('beta');
await page.waitForTimeout(150);
const order = await page.evaluate(() => { const p = window.odx.view.state.doc.firstChild; let out = ''; p.forEach((n) => { if (n.isText) out += n.marks.some((m) => m.type.name === 'del') ? `[-${n.text}]` : n.marks.some((m) => m.type.name === 'ins') ? `[+${n.text}]` : ''; }); return out; });
check(/\[\+beta\]\[-alpha\]$/.test(order), `a word deleted letter by letter and typed over reads in order (${order})`);
const insBefore = await page.evaluate(() => document.querySelectorAll('.odx-content ins.odx-ins').length);
await page.waitForTimeout(700); // (a new undo step, not the typing before)
await page.keyboard.type('X');
await page.keyboard.press('Control+z');
await page.waitForTimeout(150);
check(await page.evaluate((n) => document.querySelectorAll('.odx-content ins.odx-ins').length === n && !window.odx.view.state.doc.textContent.includes('betaX'), insBefore), 'undoing tracked typing takes its insertion away');
const cellOk = await page.evaluate(async () => {
  const v = window.odx.view;
  let at = null;
  v.state.doc.descendants((n, pos) => { if (at == null && n.type.name === 'table_cell') at = pos + 2; return at == null; });
  if (at == null) return 'no table';
  v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.near(v.state.doc.resolve(at))));
  v.focus();
  return 'ok';
});
await page.keyboard.type('cell text');
await page.evaluate(() => { const b = [...document.querySelectorAll('.odx-tracksel')][0]; void b; });
await page.evaluate(() => window.odx.agent().resolveChanges({ accept: false }));
const strayIns = /<w:ins\b/.test(strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/document.xml']));
check(cellOk === 'ok' && !strayIns && !(await bodyText()).includes('cell text'), `Reject All with a change in a table cell leaves no insertion behind (${cellOk}, ${strayIns})`);

// Inside a protected document's editable part: Enter, Tab and paste all work there.
await open('protect-readonly.docx');
// (The caret after "Everyone may edit", inside the part.)
await page.evaluate(() => {
  const v = window.odx.view;
  let at = null;
  v.state.doc.descendants((n, pos) => { if (at == null && n.isText && n.text.includes('may edit')) at = pos + n.text.indexOf('may edit') + 8; return at == null; });
  v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.near(v.state.doc.resolve(at))));
  v.focus();
});
await page.keyboard.press('Tab');
await page.keyboard.type('T');
await page.keyboard.press('Enter');
await page.keyboard.type('N');
await page.evaluate(() => { const dt = new DataTransfer(); dt.setData('text/plain', 'P'); window.odx.view.dom.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); });
await page.waitForTimeout(150);
const permText = await page.evaluate(() => { const d = window.odx.view.state.doc; return d.textBetween(0, d.content.size, '\n', (n) => (n.type.name === 'tab' ? '\t' : '')); });
check(/edit\tT\nNP this/.test(permText), `Tab, Enter and paste inside an editable part of a protected document (${JSON.stringify(permText.slice(0, 80))})`);

// Zoom (View > Zoom, zoom.js): only what's shown changes. The pages break the same at 150%, a click puts the caret where
// it lands, a double-click on the header edits it there, Ctrl+wheel zooms, and a PDF export is the same as at 100%.
await open('fields.docx');
const zoomSnap = () => page.evaluate(() => { const v = window.odx.view; const sh = v.dom.parentElement; const z = parseFloat(sh.style.zoom) || 1; const top = sh.getBoundingClientRect().top; return { pages: v.odx.pages.count, tops: [...v.dom.querySelectorAll(':scope > p')].filter((_, i) => i % 8 === 0).map((p) => Math.round((p.getBoundingClientRect().top - top) / z)) }; });
const z100 = await zoomSnap();
const pdf100 = await page.evaluate(async () => (await window.odx.exportPdf()).length);
await page.evaluate(() => window.odx.setZoom(1.5));
await page.waitForTimeout(1500);
const z150 = await zoomSnap();
// (Within a pixel, and a thousandth of how far down the page stack it is: zoomed, the browser lays every length out
// on a grid half again as fine, a 64th of a pixel a line or a paragraph, 2px by the fourth page here.)
check(z150.pages === z100.pages && z150.tops.every((t, i) => Math.abs(t - z100.tops[i]) <= 1 + z100.tops[i] / 1000), `at 150% the pages break as at 100% (${z100.pages} pages; ${z150.tops.join(',')} for ${z100.tops.join(',')})`);
await page.click('[role=tab]:text-is("View")');
check(await page.evaluate(() => document.querySelector('.odx-zoomsel').value === '1.5'), 'View > Zoom shows 150%');
await page.click('.odx-content p:has-text("Front matter paragraph 2.")', { position: { x: 30, y: 8 } });
await page.keyboard.type('@@');
check(await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].find((p) => p.textContent.includes('@@'))?.textContent.replace('@@', '').startsWith('Front matter paragraph 2.')), 'at 150% a click puts the caret where it lands');
await page.keyboard.press('Control+z');
const hdr = await page.evaluate(() => { const r = document.querySelector('.odx-sheet .odx-hf')?.getBoundingClientRect(); return r && { x: r.left + r.width / 2, y: r.top + r.height / 2 }; });
await page.mouse.dblclick(hdr.x, hdr.y);
await page.waitForTimeout(400);
const hfAt = await page.evaluate(() => { const e = document.querySelector('.odx-hfedit-header'); const h = document.querySelector('.odx-sheet .odx-hf'); return e && h ? Math.round(Math.abs(e.getBoundingClientRect().top - h.getBoundingClientRect().top)) : null; });
check(hfAt != null && hfAt < 30, `at 150% a double-click on the header edits it in place (${hfAt}px off)`);
await page.keyboard.press('Escape');
await page.mouse.move(640, 600);
await page.keyboard.down('Control');
await page.mouse.wheel(0, -100);
await page.keyboard.up('Control');
await page.waitForTimeout(300);
check(await page.evaluate(() => Math.abs(window.odx.zoom - 1.6) < 1e-6), `Ctrl+wheel zooms in 10% (${await page.evaluate(() => window.odx.zoom)})`);
const pdf160 = await page.evaluate(async () => (await window.odx.exportPdf()).length);
check(Math.abs(pdf160 - pdf100) < pdf100 * 0.01 && await page.evaluate(() => window.odx.view.dom.parentElement.style.zoom === '1.6'), `a PDF export at 160% is the 100% one (${pdf100} and ${pdf160} bytes), the zoom kept on screen`);
await page.evaluate(() => window.odx.setZoom('pageWidth'));
await page.waitForTimeout(300);
const fit = await page.evaluate(() => { const sh = window.odx.view.dom.parentElement; const c = sh.parentElement; const cs = getComputedStyle(c); return Math.round(c.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight) - sh.getBoundingClientRect().width); });
check(Math.abs(fit) <= 1, `Page Width fits the page to the window (${fit}px left over)`);
await page.evaluate(() => window.odx.setZoom(1));
await page.click('[role=tab]:text-is("Home")');

// Find into a comment, then the find bar closed: the comment closes too, and typing goes on in the document (the
// comment's box had taken the keys back at the next redraw).
await open('find-all.docx');
await page.click('.odx-content p');
await page.keyboard.press('Control+f');
await page.fill('.odx-findbar input[aria-label="Find"]', 'comment about');
await page.press('.odx-findbar input[aria-label="Find"]', 'Enter');
await page.waitForTimeout(300);
await page.press('.odx-findbar input[aria-label="Find"]', 'Escape');
await page.click('.odx-content p >> nth=1');
await page.keyboard.press('End');
await page.keyboard.type('QQ');
await page.waitForTimeout(600);
const afterFind = await page.evaluate(() => ({ inDoc: !!document.activeElement?.closest('.odx-content'), box: !!document.querySelector('.odx-markup textarea'), typed: window.odx.view.state.doc.textContent.includes('QQ') }));
check(afterFind.inDoc && !afterFind.box && afterFind.typed, `Find into a comment, closed: the comment closes and typing stays in the document (${JSON.stringify(afterFind)})`);

// Command search (Alt+Q, cmdsearch.js): every ribbon command by name; Enter runs it, Esc goes back to the document.
await open('rich.docx');
await page.click('.odx-content p');
await page.keyboard.press('Home');
await page.keyboard.press('Shift+End');
await page.waitForTimeout(300);
const boldBefore = await page.evaluate(() => { const v = window.odx.view; const { from, to } = v.state.selection; let b = false; v.state.doc.nodesBetween(from, to, (n) => { if (n.isText && n.marks.some((m) => m.type.name === 'bold')) b = true; }); return { b, sel: to - from }; });
await page.keyboard.press('Alt+q');
const csFocused = await page.evaluate(() => document.activeElement?.classList.contains('odx-cmdsearch'));
await page.keyboard.type('bold');
await page.waitForTimeout(200);
const csList = await page.evaluate(() => ({ expanded: document.querySelector('.odx-cmdsearch').getAttribute('aria-expanded'), first: document.querySelector('.odx-cmdopt')?.textContent, active: document.querySelector('.odx-cmdsearch').getAttribute('aria-activedescendant') }));
await page.keyboard.press('Enter');
await page.waitForTimeout(300);
const csBold = await page.evaluate(() => { const v = window.odx.view; const { from, to } = v.state.selection; let bold = false; v.state.doc.nodesBetween(from, to, (n) => { if (n.isText && n.marks.some((m) => m.type.name === 'bold')) bold = true; }); return { bold, sel: to - from, inDoc: !!document.activeElement?.closest('.odx-content') }; });
check(csFocused && csList.expanded === 'true' && /^Bold.*Ctrl\+B.*Home/.test(csList.first || '') && !csList.active && boldBefore.sel > 0 && csBold.bold !== boldBefore.b && csBold.inDoc,
  `Alt+Q finds Bold (with its shortcut and tab) and Enter runs it, back in the document (${JSON.stringify({ csFocused, ...csList, before: boldBefore, ...csBold })})`);
await page.keyboard.press('Control+z');
const pagesBefore = await page.evaluate(() => window.odx.view.state.doc.textContent.length);
await page.keyboard.press('End');
await page.keyboard.press('Alt+q');
await page.keyboard.type('page break');
await page.waitForTimeout(200);
await page.keyboard.press('Enter');
await page.waitForTimeout(300);
check(await page.evaluate(() => !!document.querySelector('.odx-content .odx-page-break, .odx-content [data-break="page"], .odx-content br.odx-pb')) || await page.evaluate(() => { let n = 0; window.odx.view.state.doc.descendants((x) => { if (x.type.name === 'hard_break' && x.attrs.type === 'page') n++; }); return n > 0; }), 'Alt+Q "page break" inserts one');
void pagesBefore;
await page.keyboard.press('Control+z');
await page.keyboard.press('Alt+q');
await page.keyboard.type('restrict');
await page.waitForTimeout(200);
await page.keyboard.press('Enter');
await page.waitForTimeout(500);
check(await page.evaluate(() => !!document.querySelector('.odx-pop:not([hidden]) .odx-protect')), 'Alt+Q "restrict" opens Restrict Editing');
await page.keyboard.press('Escape');
await page.evaluate(() => window.odx.view.focus());
await page.keyboard.press('Alt+q');
await page.keyboard.type('zoom in');
await page.waitForTimeout(200);
await page.keyboard.press('Enter');
await page.waitForTimeout(300);
check(await page.evaluate(() => window.odx.zoom > 1), `Alt+Q "zoom in" zooms (${await page.evaluate(() => window.odx.zoom)})`);
await page.evaluate(() => window.odx.setZoom(1));
await page.keyboard.press('Alt+q');
await page.keyboard.type('xyzzy');
await page.waitForTimeout(200);
await page.keyboard.press('Escape');
check(await page.evaluate(() => !!document.activeElement?.closest('.odx-content') && document.querySelector('.odx-cmdlist').hidden), 'Esc in the command search goes back to the document');
await page.click('[role=tab]:text-is("Home")');

// Insert > Table of Contents in a document without Word's TOC styles: they're added (TOC Heading, TOC 1-3, as Word's
// own), its lines in them, and saved with the document.
await open('fields.docx');
await page.click('.odx-content p');
await page.keyboard.press('Control+Home');
await page.evaluate(() => { const b = [...document.querySelectorAll('button')].find((x) => /Insert table of contents/.test(x.title || '')); b.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })); });
await page.waitForTimeout(1500);
const tocStyles = await page.evaluate(async () => {
  const styles = [...new Set([...document.querySelectorAll('.odx-content > p[data-style^="TOC"]')].map((p) => p.dataset.style))];
  const indent = getComputedStyle(document.querySelector('.odx-content > p[data-style="TOC2"]') || document.body).marginLeft;
  return { styles: styles.sort().join(','), indent };
});
const tocSaved = unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))));
const stylesXml = strFromU8(tocSaved['word/styles.xml']);
check(tocStyles.styles === 'TOC1,TOCHeading' && /<w:style [^>]*w:styleId="TOC2"[\s\S]*?<w:ind w:left="220"\/>/.test(stylesXml) && /w:styleId="TOC1"/.test(stylesXml) && /w:styleId="TOCHeading"/.test(stylesXml),
  `a table of contents gets Word's TOC styles when the document lacks them (${JSON.stringify(tocStyles)})`);

// A document from a localised Word (Spanish ids: "Ttulo1" is Heading 1, "TDC1" TOC 1): its table of contents in its own
// TOC styles, none added under the English ids, and a pasted <h1> in its own Heading 1.
await open('localised-styles.docx');
await page.click('.odx-content p');
await page.keyboard.press('Control+Home');
await page.evaluate(() => { const b = [...document.querySelectorAll('button')].find((x) => /Insert table of contents/.test(x.title || '')); b.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })); });
await page.waitForTimeout(1500);
await page.evaluate(() => {
  // (Into a new, empty paragraph at the end.)
  const v = window.odx.view;
  const end = v.state.doc.content.size;
  const tr = v.state.tr.insert(end, v.state.schema.nodes.paragraph.create());
  v.dispatch(tr.setSelection(v.state.selection.constructor.create(tr.doc, end + 1)));
  const dt = new DataTransfer(); dt.setData('text/html', '<p>Antes</p><h1>Pegado</h1><p>Mas</p>'); dt.setData('text/plain', 'Antes Pegado Mas');
  window.odx.view.dom.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
});
await page.waitForTimeout(300);
const loc = await page.evaluate(() => ({
  toc: [...new Set([...document.querySelectorAll('.odx-content > p[data-style^="T"]')].map((p) => p.dataset.style))].sort().join(','),
  pasted: [...document.querySelectorAll('.odx-content > p')].find((p) => p.textContent === 'Pegado')?.dataset.style,
}));
const locXml = strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/styles.xml']);
check(loc.toc === 'TDC1,TDC2,Ttulo1,Ttulo2,TtuloTDC' && loc.pasted === 'Ttulo1' && !/w:styleId="(TOC1|TOCHeading|Heading1)"/.test(locXml),
  `a localised document's TOC and pasted heading use its own style ids (${JSON.stringify(loc)})`);

// Update table on a TOC \u \h \z (toc-u-only.docx, made with Word COM): the field keeps its switches and lists what Word's
// own update did, outline levels set on paragraphs or through styles, 1-9, and no style based on Heading 1 at body text.
const updateToc = async (file) => {
  await open(file);
  await page.click('.odx-content p');
  await page.click('.odx-rtab:text-is("References")');
  await page.click('button[title="Table of contents"]');
  await page.click('.odx-pop .odx-item:has-text("Update table")');
  await page.waitForTimeout(500);
  return tocEntries();
};
const tocEntries = () => page.evaluate(() => {
  const words = [];
  let instr = null;
  window.odx.view.state.doc.forEach((p) => {
    p.forEach((c) => { if (instr == null && c.type.name === 'opaque_inline' && c.attrs.kind === 'instrText') instr = /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/.exec(c.attrs.xml)?.[1] ?? ''; }); // (its text, as merge.js reads it)
    const m = /^TOC(\d)$/.exec(p.attrs.styleId || '');
    if (!m) return;
    let t = '', tab = false;
    p.forEach((c) => { if (c.type.name === 'tab') tab = true; else if (!tab && c.isText) t += c.text; });
    words.push(`${t.trim()}:${m[1]}`);
  });
  return { instr: instr.trim().replace(/&quot;/g, '"'), words: words.slice(0, 9).join(', ') };
});
const tocU = await updateToc('toc-u-only.docx');
check(tocU.instr === 'TOC \\u \\h \\z' && tocU.words === 'H1 styled:1, H2 styled:2, H4 styled:4, Direct level 1:1, Direct level 5:5, Direct level 9:9, Custom style level 2:2, Custom style level 6:6, H1 styled:1',
  `a \\u-only TOC updates to Word's entries (${JSON.stringify(tocU)})`);
// And a TOC \t "Section Title,1,Deep Title,3" \h \z (toc-t-switch.docx): those two styles at those levels, nothing else
// (the next entry is the second table's, as Word left it).
const tocT = await updateToc('toc-t-switch.docx');
check(tocT.instr === 'TOC \\t "Section Title,1,Deep Title,3" \\h \\z' && tocT.words.startsWith('Custom style level 2:1, Custom style level 6:3, H1 styled:1'),
  `a \\t TOC updates to Word's entries (${JSON.stringify(tocT)})`);
// End on a TOC entry, then typing straight away: every key at the end of the entry, after its page number. (Its last
// node is the PAGEREF field's end, inside the entry's hyperlink: the browser's End stopped before it or went to the end
// of the document, and the first keys typed went one place, the rest the other.)
await open('toc-u-only.docx');
await page.click('.odx-content p >> nth=0');
await page.waitForTimeout(200);
await page.keyboard.press('End');
await page.keyboard.type(' typed');
await page.waitForTimeout(300);
const endTyped = await page.evaluate(() => `${window.odx.view.state.doc.firstChild.textContent}|${window.odx.view.state.doc.lastChild.textContent}`);
check(endTyped === 'H1 styled1 typed|Body text', `End on a TOC entry and typing: all at the entry's end (${endTyped})`);
// And End then Home on the next entry: at its start (the browser's Home from the end went to after the entry's tab, or
// out of the paragraph).
await page.click('.odx-content > p >> nth=1');
await page.keyboard.press('End');
await page.keyboard.press('Home');
await page.keyboard.type('Z!');
await page.waitForTimeout(300);
const homeTyped = await page.evaluate(() => `${window.odx.view.state.doc.child(1).textContent}|${window.odx.view.state.doc.textContent.split('Z!').length - 1}`);
check(homeTyped === 'Z!H2 styled1|1', `End, Home on a TOC entry and typing: all at the entry's start (${homeTyped})`);

// Insert > Table as Word's: in the Table Grid style (added to the document), lines drawn all round and between the
// cells, the cells' text with no space after and single spacing (rows 18.33px and 19px in Word's own print of it).
await page.evaluate(() => window.odx.newDocument({ share: false }));
await page.click('.odx-content p');
await page.click('.odx-rtab:text-is("Insert")');
await page.click('button[title="Insert table"]');
await page.click('.odx-pop .odx-item >> nth=0');
await page.waitForTimeout(500);
const grid = await page.evaluate(() => {
  const t = document.querySelector('.odx-content table');
  const td = t.querySelector('td');
  return { tstyle: t.dataset.tstyle, lines: getComputedStyle(td).boxShadow.split('inset').length - 1, after: getComputedStyle(td.querySelector('p')).marginBottom, rows: [...t.querySelectorAll('tr')].map((r) => Math.round(r.getBoundingClientRect().height)) };
});
const gridXml = unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))));
check(grid.tstyle === 'TableGrid' && grid.lines >= 3 && grid.after === '0px' && grid.rows.every((h) => h >= 18 && h <= 20)
  && /<w:tblStyle w:val="TableGrid"\/>/.test(strFromU8(gridXml['word/document.xml'])) && /w:styleId="TableGrid"[\s\S]*?<w:insideV w:val="single"/.test(strFromU8(gridXml['word/styles.xml'])),
  `an inserted table is in Word's Table Grid style, its lines drawn (${JSON.stringify(grid)})`);

// Merging cells keeps the first cell's formatting (shading, borders, alignment) and every cell's paragraphs, as Word.
await open('merge-format.docx');
await page.locator('.odx-content td p', { hasText: 'A1' }).first().click();
await page.locator('.odx-content td p', { hasText: 'A3' }).first().click({ modifiers: ['Shift'] });
await page.evaluate(() => document.querySelector('button[title="Merge cells"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })));
await page.waitForTimeout(400);
const merged = await page.evaluate(() => {
  const td = document.querySelector('.odx-content td');
  return { span: td.rowSpan, paras: td.querySelectorAll('p').length, bg: getComputedStyle(td).backgroundColor, va: getComputedStyle(td).verticalAlign, align: getComputedStyle(td.querySelector('p')).textAlign };
});
const mergedXml = strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/document.xml']);
check(merged.span === 3 && merged.paras === 3 && merged.bg === 'rgb(255, 230, 153)' && merged.va === 'middle' && merged.align === 'center'
  && /<w:vMerge w:val="restart"\/><w:tcBorders><w:left w:val="single" w:sz="24"[^>]*\/><\/w:tcBorders><w:shd [^>]*w:fill="FFE699"\/><w:vAlign w:val="center"\/>/.test(mergedXml),
  `merged cells keep their formatting (${JSON.stringify(merged)})`);

// Layout > Columns with a selection: for the selected text only (continuous section breaks before and after it, on the
// paragraphs' own marks, none split), as Word's menu; "Two, from here on": from the cursor on.
const colsDoc = async (sel) => {
  await page.evaluate(() => window.odx.newDocument({ share: false }));
  await page.click('.odx-content p');
  for (const t of ['First paragraph, one column.', 'Second paragraph, in two columns.', 'Third paragraph, one column again.']) { await page.keyboard.type(t); await page.keyboard.press('Enter'); }
  await page.keyboard.type('Fourth.');
  await page.evaluate(sel);
  await page.click('.odx-rtab:text-is("Layout")');
  await page.click('button[title="Columns"]');
};
const colsSaved = async () => {
  await page.waitForTimeout(1200);
  const x = strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/document.xml']);
  return [...x.matchAll(/<w:p\b[\s\S]*?<\/w:p>|<w:sectPr\b[\s\S]*?<\/w:sectPr>(?=<\/w:body>)/g)].map((m) => {
    const t = (m[0].match(/<w:t[^>]*>([^<]*)/g) || []).map((s) => s.replace(/<w:t[^>]*>/, '')).join('').slice(0, 6);
    const s = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(m[0])?.[0];
    return s ? `${t}|${/<w:cols\b[^>]*w:num="(\d)"/.exec(s)?.[1] || 1}${/w:type w:val="continuous"/.test(s) ? 'c' : ''}` : t;
  }).join(' / ');
};
await colsDoc(() => {
  // (The second paragraph's text selected.)
  const v = window.odx.view; let a = null, b = null;
  v.state.doc.forEach((n, pos, i) => { if (i === 1) { a = pos + 1; b = pos + 1 + n.content.size; } });
  v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.create(v.state.doc, a, b)));
});
await page.click('.odx-pop .odx-item:has-text("Two") >> nth=0');
const colsSel = await colsSaved();
await colsDoc(() => {
  const v = window.odx.view; let a = null;
  v.state.doc.forEach((n, pos, i) => { if (i === 2) a = pos + 1; });
  v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.create(v.state.doc, a)));
});
await page.click('.odx-pop .odx-item:has-text("Two, from here on")');
const colsFwd = await colsSaved();
check(colsSel === 'First |1 / Second|2c / Third  / Fourth / |1c' && colsFwd === 'First  / Second|1 / Third  / Fourth / |2c',
  `Columns for the selected text, and from here on (${colsSel} || ${colsFwd})`);

// Automatic hyphenation (w:autoHyphenation): words' break points shown as soft hyphens, so a narrow paragraph's lines
// end in a hyphen as in Word; not in a paragraph with suppressAutoHyphens; nothing of it saved; printed with its hyphens.
await open('hyphenation.docx');
await page.waitForFunction(() => document.querySelectorAll('.odx-content .odx-shy, .odx-content p.odx-hyph').length > 0, null, { timeout: 5000 }).catch(() => {});
const hy = await page.evaluate(() => {
  const [a, b] = document.querySelectorAll('.odx-content > p');
  const lines = (p) => { const r = document.createRange(); r.selectNodeContents(p); return new Set([...r.getClientRects()].map((x) => Math.round(x.top))).size; };
  // (Soft hyphens where the browser can't hyphenate itself, its own hyphens:auto where it can: hyphenate.js.)
  const on = (p) => p.querySelectorAll('.odx-shy').length > 10 || p.classList.contains('odx-hyph');
  return { onA: on(a), onB: on(b), linesA: lines(a), linesB: lines(b) };
});
const hySaved = strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/document.xml']);
await page.emulateMedia({ media: 'print' });
await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
const hyPrinted = await page.evaluate(() => [...document.querySelectorAll('.odx-printlines p')].filter((p) => /-$/.test(p.textContent)).length);
await page.evaluate(() => window.dispatchEvent(new Event('afterprint')));
await page.emulateMedia({ media: 'screen' });
check(hy.onA && !hy.onB && hy.linesA > 1 && hy.linesA <= hy.linesB && !hySaved.includes('\u00ad') && /<w:suppressAutoHyphens\/>/.test(hySaved) && hyPrinted > 0,
  `automatic hyphenation: break points shown, not saved, printed with hyphens (${JSON.stringify({ ...hy, hyPrinted })})`);
// Nothing the layout measures with is put inside the document's own paragraphs: the editor reads a change there as an
// edit made in the page and draws the paragraph again from it. A probe measuring a run in another font than its
// paragraph's went there, and with Track Changes on a document of tables with merged cells saved, untouched, its tables
// twice over as tracked deletions and insertions (mktrackmixed.mjs: tracked, merged cells, mixed fonts).
{
  await page.evaluate(() => {
    window.__inContent = 0;
    const content = () => document.querySelector('.odx-content');
    window.__probeWatch?.disconnect();
    window.__probeWatch = new MutationObserver((list) => {
      for (const m of list) for (const n of m.addedNodes) {
        if (n.nodeType === 1 && n.getAttribute('aria-hidden') === 'true' && /visibility:\s*hidden/.test(n.getAttribute('style') || '') && content()?.contains(m.target)) window.__inContent++;
      }
    });
    window.__probeWatch.observe(document.body, { subtree: true, childList: true });
  });
  await open('track-mixed-table.docx');
  await page.waitForTimeout(1500);
  const r = await page.evaluate(() => {
    window.__probeWatch.disconnect();
    return { probes: window.__inContent, tracked: document.querySelectorAll('.odx-content ins.odx-ins, .odx-content del.odx-del').length };
  });
  check(r.probes === 0 && r.tracked === 0, `measuring probes stay out of the document's paragraphs; an untouched tracked document gains no changes (${JSON.stringify(r)})`);
}
// A first-page header (and no other) with its distance far below a small top margin, holding a picture placed against
// the page, wrapped square (mkhdrplace.mjs). Word's PDF of it: the picture 48px down and 192px in, the first page's
// text from under the header's line, the other pages' from the top margin (their first line's text 38.7px down in Word's
// PDF, 140px on the first page).
{
  await open('header-page-placed.docx');
  const r = await page.evaluate(() => {
    const sheets = [...document.querySelectorAll('.odx-sheet')];
    const img = sheets[0]?.querySelector('.odx-hf img');
    const s0 = sheets[0].getBoundingClientRect(), s1 = sheets[1]?.getBoundingClientRect();
    const firstText = (top) => { for (const p of document.querySelectorAll('.odx-content > p')) { const q = p.getBoundingClientRect(); if (q.top >= top && q.height) return Math.round((q.top - top) * 10) / 10; } return null; };
    const ir = img?.getBoundingClientRect();
    return { imgTop: ir && Math.round((ir.top - s0.top) * 10) / 10, imgLeft: ir && Math.round((ir.left - s0.left) * 10) / 10, page1: firstText(s0.top), page2: s1 && firstText(s1.top) };
  });
  check(Math.abs(r.imgTop - 48) <= 1 && Math.abs(r.imgLeft - 192) <= 1 && r.page1 > 100 && r.page2 < 40,
    `a header's picture placed against the page, as Word draws it; pages without a header from the top margin (${JSON.stringify(r)})`);
}
// A header's empty paragraph spaced 6pt before holding a picture wrapped top and bottom that reaches past the top margin,
// in compatibility mode 15 (mkhdrtb.mjs). Word's PDF of it: the body's first line from 86.7pt (115.6px) down, its
// paragraph's space before and its mark's line under the picture; it had been 9.6px higher, right under the picture.
{
  await open('header-top-bottom.docx');
  const r = await page.evaluate(() => {
    const s0 = document.querySelector('.odx-sheet').getBoundingClientRect();
    const p = [...document.querySelectorAll('.odx-content > p')].find((e) => e.getBoundingClientRect().height);
    return { body: p && Math.round((p.getBoundingClientRect().top - s0.top) * 10) / 10 };
  });
  check(Math.abs(r.body - 115.6) <= 1.5, `a header's picture wrapped top and bottom: the body under its paragraph's space and line, as in Word (${JSON.stringify(r)})`);
}
// A table style's hanging indent (mktblind.mjs), its own and its base's: in Word's PDF each Normal cell's first line at
// the cell's edge and the rest 18pt in, the same under a style based on Normal, none under a style of its own indent,
// and with a direct 0.5in left indent the first line 18pt in from that.
{
  await open('table-style-indent.docx');
  const r = await page.evaluate(() => [...document.querySelectorAll('.odx-content td p')].filter((p) => p.textContent.length > 10).map((p) => {
    const cs = getComputedStyle(p);
    return `${p.textContent.split(':')[0]} ${Math.round(parseFloat(cs.marginLeft) * 0.75)}/${Math.round(parseFloat(cs.textIndent) * 0.75)}`;
  }).join(', '));
  check(r === 'Normal 18/-18, Body 18/-18, Own 0/0, Direct 36/-18, Normal 18/-18, Body 18/-18, Own 0/0, Direct 36/-18', `a table style's indent on its cells' paragraphs, as in Word (${r})`);
}
// Contextual spacing across a table's edges (mktableedges.mjs): with it, the paragraph above, the cell's and the one
// under the table meet with no space (Word: each line 13.5pt under the one above, and the border); without it, each
// keeps its 10pt (the cell's line 33.75pt under the one above, the line under the table 33.75pt under the cell's).
{
  await open('table-edges.docx');
  const r = await page.evaluate(() => {
    const at = (t) => [...document.querySelectorAll('.odx-content p')].find((p) => p.textContent === t)?.getBoundingClientRect();
    const gap = (a, b) => { const x = at(a), y = at(b); return x && y ? Math.round((y.top - x.bottom) * 10) / 10 : null; };
    return { withTop: gap('Above with', 'Cell with'), withBottom: gap('Cell with', 'Below with'), withoutTop: gap('Above without', 'Cell without'), withoutBottom: gap('Cell without', 'Below without') };
  });
  check(r.withTop < 4 && r.withBottom < 4 && r.withoutTop > 24 && r.withoutBottom > 24, `contextual spacing at a table's edges, as in Word (${JSON.stringify(r)})`);
}
// A row of empty cells that leave their end marks out of its height (w:hideMark; mktableedges.mjs): no room, as in
// Word (the third row's line 14.25pt under the first's), where the same row without it keeps a line (36.75pt).
{
  await open('hide-mark.docx');
  const r = await page.evaluate(() => {
    const at = (t) => [...document.querySelectorAll('.odx-content p')].find((p) => p.textContent === t)?.getBoundingClientRect().top;
    return { hidden: Math.round((at('Third hidden') - at('First hidden')) * 10) / 10, shown: Math.round((at('Third shown') - at('First shown')) * 10) / 10 };
  });
  check(Math.abs(r.hidden - 19) <= 1.5 && Math.abs(r.shown - 49) <= 1.5, `a row of empty cells with hidden end marks takes no room, as in Word (${JSON.stringify(r)})`);
}
// Autofit columns and cells' auto spacing as Word lays them (mkautofit.mjs; Word's prints of the same files).
// A 40px first column holding empty bulleted paragraphs, or "x" after a first-line indent, stays 40px; "x" bulleted at
// a 720 indent widens it to 68.8px (its indent, word and margins).
{
  await open('autofit-indents.docx');
  const ws = await page.evaluate(() => [...document.querySelectorAll('.odx-content table')].map((t) => +t.rows[0].cells[0].getBoundingClientRect().width.toFixed(2)));
  check(ws.length === 3 && Math.abs(ws[0] - 40) < 1 && Math.abs(ws[1] - 40) < 1 && Math.abs(ws[2] - 68.8) < 1, `an autofit column needs no room for an empty paragraph's indent or a first-line indent (${ws})`);
}
// A Normal (Web) paragraph last in its cell has no auto space under it (its row as high as the next, but for the
// table's top line); first of two, it keeps 14pt, and so does a list's last item ending a cell.
{
  await open('cell-auto-space.docx');
  const r = await page.evaluate(() => {
    const ts = [...document.querySelectorAll('.odx-content table')];
    const b = [...document.querySelectorAll('.odx-content td p')].find((q) => q.textContent === 'B1b');
    const a = b?.previousElementSibling;
    const item = [...document.querySelectorAll('.odx-content td p')].find((q) => q.textContent === 'B2b');
    return { rows: ts[0] ? [...ts[0].rows].map((tr) => +tr.getBoundingClientRect().height.toFixed(2)) : [], gap: a ? +(b.getBoundingClientRect().top - a.getBoundingClientRect().bottom).toFixed(2) : null, item: item ? parseFloat(getComputedStyle(item).marginBottom) : null };
  });
  check(r.rows.length === 2 && Math.abs(r.rows[0] - r.rows[1]) < 1 && Math.abs(r.gap - 18.67) < 0.5 && Math.abs(r.item - 18.67) < 0.5, `no auto space after a cell's last paragraph but a list item, 14pt after its first (${JSON.stringify(r)})`);
}
// Table rows and notes as Word lays them (mktablerows.mjs; Word's prints of the same files).
// A note at the foot of its own row's page, not of the table's first page: page 1 has notes 1-15, page 3 notes 32-40.
{
  await open('table-row-notes.docx');
  const b = await bands();
  const ids = (k) => (b[k]?.parts || []).flatMap((q) => q.ids).join();
  const want = (a, z) => Array.from({ length: z - a + 1 }, (_, i) => String(a + i)).join();
  check(ids(0) === want(1, 15) && ids(2) === want(32, 40), `a table's notes on the pages of their rows (${[0, 1, 2].map(ids).join(' | ')})`);
}
// Each cell's text under its row's widest top line: a 3pt line over the one-line cell, 0.5pt over the three-line one,
// both 3pt down, the rows 62.67px apart.
{
  await open('table-row-band.docx');
  const r = await placeOf(['Alpha1', 'B1 one', 'B2 one']);
  const [a, b1, b2] = r.at;
  check(a && b1 && b2 && Math.abs(a.top - b1.top) < 0.5 && Math.abs(b2.top - b1.top - 62.67) < 0.5, `a row's text under its widest top line, its own thinner one or not (${JSON.stringify(r.at.map((x) => x && +x.top.toFixed(2)))})`);
}
// A row of at least 32.25pt with 33.75pt left at the page's foot, its own 1pt top line and the next row's under it: it
// splits there (its first two lines on page 1), as in Word.
{
  await open('table-row-foot.docx');
  // (Where each line's text is: a line carried on is drawn under a spacer at its paragraph's start.)
  const at = await page.evaluate((ts) => {
    const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
    return ts.map((t) => {
      const p = [...document.querySelectorAll('.odx-content p')].find((q) => q.textContent === t);
      const node = p && [...p.querySelectorAll('span.odx-t')].pop()?.firstChild;
      if (!node) return null;
      const rg = document.createRange();
      rg.selectNodeContents(node);
      const top = rg.getBoundingClientRect().top;
      return sheets.findIndex((q) => top >= q.top && top < q.bottom);
    });
  }, ['R1 first', 'R1 second', 'R1 third']);
  check(at.join() === '0,0,1', `a row of at least height splits at the page's foot with room for it and its own lines (${at})`);
}
// Notes in a Times New Roman style, their text in Calibri after a Times quote mark: Calibri's lines, 16.28px apart (the
// browser's own would be 17px; the style's Times 15.33px).
{
  await open('notes-mixed-font.docx');
  const hs = await page.evaluate(() => [...document.querySelectorAll('.odx-sheet .odx-notes .odx-note p')].map((p) => +p.getBoundingClientRect().height.toFixed(2)));
  check(hs.length === 3 && hs.every((h) => Math.abs(h - 2 * 16.276) < 0.6), `notes' lines at their text's font, not their style's (${hs}: two lines each)`);
}
// Normal in 16pt: the separator a line of it, page 1 keeps 30 lines of 20pt over the note (31 under Normal in 12pt).
{
  await open('notes-separator-16pt.docx');
  const n = await page.evaluate(() => {
    const s = document.querySelector('.odx-sheets > .odx-sheet').getBoundingClientRect();
    return [...document.querySelectorAll('.odx-content p')].filter((p) => p.textContent.startsWith('Body line') && p.getBoundingClientRect().top < s.bottom).length;
  });
  check(n === 30, `the note separator a line of the Normal style: ${n} lines on page 1 (Word: 30)`);
}
// A form's merged cells, framed text boxes and columns as Word lays them out (mkrowspan.mjs; Word's prints of the same
// files). Where a paragraph's text is: its sheet, and its top on it (a moved column's too: its transform counts).
{
  const tops = (ts) => page.evaluate((list) => {
    const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
    const ps = [...document.querySelectorAll('.odx-content p')];
    return list.map((t) => {
      const p = ps.find((q) => q.textContent === t);
      if (!p) return null;
      const top = p.getBoundingClientRect().top;
      const k = sheets.findIndex((q) => top >= q.top && top < q.bottom);
      return { sheet: k, top: +(top - (sheets[k]?.top ?? 0)).toFixed(2) };
    });
  }, ts);
  // A merged cell whose text fits the rows' own minimum heights doesn't make the last of them taller (Word: "After
  // table" 132.7px under "Before").
  await open('rowspan-mixed.docx');
  let [a, b] = await tops(['Before', 'After table']);
  check(a && b && Math.abs(b.top - a.top - 132.7) < 1, `rows beside a merged cell keep their minimum heights when its text fits them (${a && b && (b.top - a.top).toFixed(2)}; Word: 132.7)`);
  // The table's 1.5pt inside line under each hidden part of a merged cell (its top nil) is in its row's height, over
  // the side cells' own 0.75pt lines (Word: "Side 6" 108px under "Side 1").
  await open('rowspan-inside.docx');
  [a, b] = await tops(['Side 1', 'Side 6']);
  check(a && b && Math.abs(b.top - a.top - 108) < 1, `rows beside a merged cell take the inside line its parts sit under (${a && b && (b.top - a.top).toFixed(2)}; Word: 108)`);
  // An inline VML group of text boxes, 120pt tall: its line that tall (Word: "After group" 174.7px under "Before"; ours
  // 2px less, the font's descent under the box: TODO.md).
  await open('vml-inline-group.docx');
  [a, b] = await tops(['Before', 'After group']);
  check(a && b && Math.abs(b.top - a.top - 174.7) < 3, `an inline VML group takes its own size in its line (${a && b && (b.top - a.top).toFixed(2)}; Word: 174.7)`);
  // Rows of a minimum height under dashed 1pt and dotted 0.75pt lines (mkrowlines.mjs): each row its minimum and its
  // top line, the text under such a line as under a solid one (Word: "A row 2" 14.08px under "A row 1", 14.07px a row
  // to "A row 12"; "C row 12" 11.0px a row under "C row 1").
  await open('row-dash-lines.docx');
  const rl = await tops(['A row 1', 'A row 2', 'A row 12', 'C row 1', 'C row 12']);
  check(rl.every(Boolean) && Math.abs(rl[1].top - rl[0].top - 14.08) < 0.5 && Math.abs((rl[2].top - rl[0].top) / 11 - 14.065) < 0.05 && Math.abs((rl[4].top - rl[3].top) / 11 - 11.0) < 0.05, `rows under dashed and dotted lines take their minimum and the line (${JSON.stringify(rl)})`);
  // Tables 341tw in, lines of 0.5 to 6pt, in a Word 2007 document (mktblcompat.mjs): each first cell's text where the
  // indents put it, 29.5px in from the paragraphs' (Word's print: "A one", "C one", "D one", "E one" all at 125.5px).
  await open('table-left-compat12.docx');
  const xs = await page.evaluate(() => ['Before A', 'A one', 'C one', 'D one', 'E one'].map((t) => {
    const p = [...document.querySelectorAll('.odx-content p')].find((q) => q.textContent === t);
    const r = document.createRange(); r.selectNodeContents(p); return p ? +r.getClientRects()[0].left.toFixed(2) : null;
  }));
  check(xs.every((x) => x != null) && xs.slice(1).every((x) => Math.abs(x - xs[0] - 29.5) < 0.5), `a Word 2007 document's tables set their text at the indent, whatever their lines (${JSON.stringify(xs)}; Word: 29.5 in)`);
  // A hanging paragraph "Label<tab><tab>Text" whose own stop lies 0 to 20tw past the hanging indent (mktabnear.mjs): the
  // second tab to that stop however near, to the next default one where the stop is the indent itself (Word: the text
  // at 297.5-298.8px, and 336.1px for the stop at the indent).
  await open('tab-near-stop.docx');
  const tn = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].map((p) => {
    const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (/^Text after/.test(n.data) ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP) });
    const t = w.nextNode(); if (!t) return null; const r = document.createRange(); r.selectNodeContents(t);
    return +(r.getClientRects()[0].left - p.getBoundingClientRect().left).toFixed(2);
  }));
  check(tn.length === 5 && tn.slice(1).every((x) => Math.abs(x - tn[1]) < 1.5) && tn[0] - tn[1] > 30, `a tab goes to a stop however near it (${JSON.stringify(tn)}; Word: the stop at the indent 38.6px further)`);
  // Cells of text turned on its side (mkturned.mjs): a paragraph's space before across the cell, its indent along it
  // (Word: "Before 105" 7px further across its cell than "Plain one", "Left 335" starting 22.4px further up hers).
  await open('turned-cells.docx');
  const tc = await page.evaluate(() => ['Plain one', 'Before 105', 'Left 335'].map((t) => {
    const p = [...document.querySelectorAll('.odx-content td p')].find((q) => q.textContent === t);
    if (!p) return null;
    const r = document.createRange(); r.selectNodeContents(p); const q = r.getClientRects()[0];
    return { left: +q.left.toFixed(2), bottom: +q.bottom.toFixed(2), cell: +p.closest('td').getBoundingClientRect().left.toFixed(2) };
  }));
  check(tc.every(Boolean) && Math.abs((tc[1].left - tc[1].cell) - (tc[0].left - tc[0].cell) - 7) < 1 && Math.abs(tc[0].bottom - tc[2].bottom - 22.4) < 1.5, `turned cells: space before across the cell, indents along it (${JSON.stringify(tc)})`);
  // A row carried to the next page starting with a paragraph spaced 195tw before (mksplitspace.mjs): its space kept there
  // (Word: "Spaced item" on page 2, its baseline 122px down, 12.7px under where a plain line starts, 109.3px).
  await open('split-row-space.docx');
  // (Where its text is, not its box: pagination moves a cell's lines inside their paragraph, by a spacer in it.)
  const ss = await page.evaluate((list) => {
    const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
    return list.map((t) => {
      const p = [...document.querySelectorAll('.odx-content p')].find((q) => q.textContent === t);
      if (!p) return null;
      const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.data.trim() && !n.parentElement.closest('.odx-push') ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP) });
      const tn = w.nextNode(); if (!tn) return null;
      const r = document.createRange(); r.selectNodeContents(tn); const top = r.getClientRects()[0]?.top;
      const k = sheets.findIndex((q) => top >= q.top && top < q.bottom);
      return { sheet: k, top: +(top - (sheets[k]?.top ?? 0)).toFixed(2) };
    });
  }, ['Line 55', 'Spaced item', 'After 1']);
  check(ss.every(Boolean) && ss[1].sheet === ss[0].sheet + 1 && Math.abs(ss[2].top - ss[1].top - 15.36) < 1 && Math.abs(ss[1].top - 109.5) < 1.5, `a paragraph starting a row's part on the next page keeps its space before (${JSON.stringify(ss)}; Word: its line 13px down)`);
  // Text-less VML shapes wrapped top and bottom (mkbandshape.mjs): an empty paragraph's line over a grey band 14pt down,
  // the next paragraph under the band; a ruled line 15pt under a line of text, the next paragraph under it (Word: "After
  // band" 73.28px under "Before", "After rule" 20.96px under the rule's line of text).
  await open('band-shapes.docx');
  const bs = await tops(['Before', 'After band', 'Text of the anchor paragraph', 'After rule']);
  check(bs.every(Boolean) && Math.abs(bs[1].top - bs[0].top - 73.28) < 1 && Math.abs(bs[3].top - bs[2].top - 20.96) < 1, `text-less bands wrapped top and bottom keep the text after them under them (${JSON.stringify(bs)})`);
  // Text at 110% (w:w) in one run, then in runs of one to five letters (mkcharscale.mjs): each as wide as Word sets it,
  // the one-letter runs scaled too (Word: 474.55, 477.53 and 482.97px; one-letter runs left at the estimate, ~2px more).
  await open('char-scale.docx');
  const ws = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].slice(0, 3).map((p) => {
    const r = document.createRange(); r.selectNodeContents(p); const rs = [...r.getClientRects()].filter((q) => q.width);
    return +(Math.max(...rs.map((q) => q.right)) - Math.min(...rs.map((q) => q.left))).toFixed(2);
  }));
  check(ws.length === 3 && [474.55, 477.53, 482.97].every((w, i) => Math.abs(ws[i] - w) < 1.2), `scaled text as wide as Word's, in runs of one letter too (${JSON.stringify(ws)}; Word: 474.55, 477.53, 482.97)`);
  // Two columns part way down a portrait page, the last section landscape, a paragraph starting with a column break:
  // the second column level with the first, on its page, and the text after the columns 16.6px under their last lines
  // (Word: "Left item 2" and "Right item 1" level; "After columns" 16.6px under "Right item 3").
  await open('cols-break-landscape.docx');
  const [l2, r1, r3, ac] = await tops(['Left item 2', 'Right item 1', 'Right item 3', 'After columns']);
  check(l2 && r1 && r3 && ac && r1.sheet === l2.sheet && Math.abs(r1.top - l2.top) < 1 && ac.sheet === r3.sheet && Math.abs(ac.top - r3.top - 16.6) < 1, `a column after a column break level with the first, the text after them under them (${JSON.stringify([l2, r1, r3, ac])})`);
  // The same columns in sections of a 1440tw bottom margin and a footer 720tw up, part way down a page of 600tw and a
  // footer 405tw up: the page filled to its own margin (Word: 62 of the lines after them on page 1; a line's leeway for
  // fonts that differ from Windows').
  await open('cols-footer.docx');
  const n = await page.evaluate(() => {
    const s = document.querySelector('.odx-sheets > .odx-sheet').getBoundingClientRect();
    return [...document.querySelectorAll('.odx-content p')].filter((p) => /^Line \d+$/.test(p.textContent) && p.getBoundingClientRect().top < s.bottom).length;
  });
  check(n >= 61 && n <= 62, `a continuous section starting part way down a page keeps the page's own margins on it: ${n} lines on page 1 (Word: 62; its own, 58)`);
}
// A form's header and footer lines in floating VML text boxes placed against the page, in one paragraph of a 14/240
// line (mkvmlhf.mjs): drawn at their places out of the flow, the footer a line of 1px, so the body runs to its bottom
// margin (Word: 59 lines on page 1, "Page 1 of 68" 10.5px under "TEST 4/20.B"'s line).
{
  await open('vml-hf-boxes.docx');
  const r = await page.evaluate(() => {
    const s = document.querySelector('.odx-sheets > .odx-sheet').getBoundingClientRect();
    const n = [...document.querySelectorAll('.odx-content p')].filter((p) => /^Line \d+$/.test(p.textContent) && p.getBoundingClientRect().top < s.bottom).length;
    const foot = document.querySelector('.odx-footer');
    const imgs = [...foot.querySelectorAll('img')].map((i) => i.getBoundingClientRect().top - s.top);
    return { n, foot: foot.querySelector('p').getBoundingClientRect().height, imgs: imgs.map((v) => +v.toFixed(1)), text: foot.textContent.trim() };
  });
  check(r.n === 59 && r.foot < 2 && !r.text && r.imgs.length === 3 && Math.abs(r.imgs[0] - 981.5) < 1.5 && Math.abs(r.imgs[2] - 1006.5) < 1.5, `floating VML text boxes in a footer drawn at their places, out of the flow (${JSON.stringify(r)}; Word: 59 lines, boxes at 981.5 and 1006.5px)`);
}
// An empty next-page section-break paragraph holding two bookmarks after a table ending 20px into its room
// (mksectbreak.mjs): kept on its page past the margin, as Word does ("Next section" on page 2, of 2).
{
  await open('sect-after-table.docx');
  const s2 = await page.evaluate(() => {
    const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((e) => e.getBoundingClientRect());
    const p = [...document.querySelectorAll('.odx-content p')].find((q) => q.textContent === 'Next section');
    return [sheets.length, p ? sheets.findIndex((q) => p.getBoundingClientRect().top >= q.top && p.getBoundingClientRect().top < q.bottom) + 1 : 0];
  });
  check(s2[0] === 2 && s2[1] === 2, `an empty section-break paragraph with bookmarks, after a table, kept on its page past the margin (sheets, "Next section" on: ${s2}; Word: 2, 2)`);
}
// A cover's lines (mkcover.mjs): a bare break in a 48pt style over 36pt text, its line the style's size (Word: 195px for
// the paragraph); a 20pt line closing on a 36pt space, which sizes nothing (Word: 65px for its two lines).
{
  await open('cover-lines.docx');
  const h = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].map((p) => +p.getBoundingClientRect().height.toFixed(2)));
  check(Math.abs(h[1] - 195) < 0.6 && Math.abs(h[2] - 65) < 0.6, `a leading break's line in its style's size, a closing space no taller line (${h.slice(1, 3)}; Word 195, 65)`);
}
// Symbol bullets over 10pt Calibri at least 13pt: 13pt apart, the bullet's reach inside the set height (Word: 17.33px).
{
  await open('list-bullet-atleast.docx');
  const t = await page.evaluate(() => { const ps = [...document.querySelectorAll('.odx-content > p')]; return ps.length === 30 ? (ps[29].getBoundingClientRect().top - ps[0].getBoundingClientRect().top) / 29 : null; });
  check(t != null && Math.abs(t - 17.333) < 0.05, `bullets under at-least spacing as Word spaces them (${t && t.toFixed(3)}px; Word 17.33)`);
}
// Shapes behind the text in a drawing order unlike the document's: each over the page's white sheet, not under it.
{
  await open('behind-shapes-order.docx');
  const z = await page.evaluate(() => {
    const sheets = Number(getComputedStyle(document.querySelector('.odx-sheets')).zIndex);
    return [...document.querySelectorAll('.odx-content img')].filter((i) => getComputedStyle(i).position === 'absolute').map((i) => Number(getComputedStyle(i).zIndex) > sheets);
  });
  check(z.length === 4 && z.every(Boolean), `shapes behind the text drawn over the page, in any drawing order (${z})`);
  // (The fourth, 600pt wide from the column, cut off at the page's right edge: 1in margin + 800px, 97px past 816px.)
  const cut = await page.evaluate(() => { const i = [...document.querySelectorAll('.odx-content img')].filter((x) => getComputedStyle(x).position === 'absolute').pop(); const s = document.querySelector('.odx-sheets > .odx-sheet').getBoundingClientRect(); const m = /inset\(\S+ (\S+)px/.exec(getComputedStyle(i).clipPath); return m ? +(i.getBoundingClientRect().right - Number(m[1]) - s.right).toFixed(1) : null; });
  check(cut != null && Math.abs(cut) < 1, `a drawing past the page's edge cut off there, as Word shows it (${cut}px from the edge)`);
}
// A grey picture shown in black and white at 75% (a:biLevel, mkcover.mjs): drawn through the page's threshold filter
// (Word: its mid-grey half black, its light half white).
{
  await open('picture-bilevel.docx');
  const f = await page.evaluate(() => { const i = document.querySelector('.odx-content img:not(.ProseMirror-separator)'); const m = /url\("?#([\w-]+)"?\)/.exec(i ? getComputedStyle(i).filter : ''); const el = m && document.getElementById(m[1]); return el ? el.querySelector('feFuncR')?.getAttribute('tableValues').split(' ').indexOf('1') : null; });
  check(f === 75, `a picture in black and white through a threshold at its 75% (${f})`);
}
// A floating table with no vertical anchor, 2in down (mktableedges.mjs): against the margin, as Word reads it, with the
// paragraphs after it above it until it (Word: "After 1" 139.5pt down the page, under "Before 5"; the table's first
// row at 216.75pt, 2in under the top margin).
{
  await open('float-no-anchor.docx');
  const r = await page.evaluate(() => {
    const sheet = document.querySelector('.odx-sheet').getBoundingClientRect();
    const at = (t) => { const p = [...document.querySelectorAll('.odx-content p')].find((x) => x.textContent === t); return p ? Math.round(p.getBoundingClientRect().top - sheet.top) : null; };
    return { after1: at('After 1'), row1: at('Floating row 1'), before5: at('Before 5') };
  });
  check(r.after1 > r.before5 && r.after1 < r.row1 && Math.abs(r.row1 - 288) <= 3, `a floating table with no vertical anchor against the margin, the text after it above it (${JSON.stringify(r)})`);
}
// A table as wide as the text, placed 451tw down from the paragraph after it (mktableedges.mjs): that paragraph above
// it, the empty one after under it, as in Word ("Text 2" 41.25pt under the table's first row; with the empty paragraph
// left beside the table, in no room, 13.5pt less).
{
  await open('float-text-noroom.docx');
  const r = await page.evaluate(() => {
    const at = (t) => { const p = [...document.querySelectorAll('.odx-content p')].find((x) => x.textContent === t); return p ? p.getBoundingClientRect().top : null; };
    return { text1: Math.round(at('Text 1') - at('Placed row 1')), text2: Math.round(at('Text 2') - at('Placed row 1')) };
  });
  check(r.text1 < 0 && Math.abs(r.text2 - 55) <= 2, `an empty paragraph after a table placed down from the text with no room beside it goes under it (${JSON.stringify(r)})`);
}
// Tables beside floats (mkfloatroom.mjs): one indented past the room a page-top logo leaves beside it starts under the
// logo, as in Word's PDF of it (the logo's foot at 113.9px, the table's first line at 128.7px; it had stayed beside the
// logo, 6px higher); one that fits beside a logo placed from its paragraph stays beside it, moved over to its wrap
// distance.
{
  await open('float-room.docx');
  const r = await page.evaluate(() => {
    const logos = [...document.querySelectorAll('.odx-content img')].map((i) => i.getBoundingClientRect()).filter((q) => q.height > 50 && q.width > 50).sort((a, b) => a.top - b.top);
    const tables = [...document.querySelectorAll('.odx-content > .tableWrapper > table, .odx-content > table')].map((t) => t.getBoundingClientRect());
    const r1 = (v) => Math.round(v * 10) / 10;
    return logos.length === 2 && tables.length === 2 ? { under: r1(tables[0].top - logos[0].bottom), besideTop: r1(tables[1].top - logos[1].bottom), besideLeft: r1(tables[1].left - logos[1].right) } : { logos: logos.length, tables: tables.length };
  });
  check(r.under >= -0.5 && r.under < 2 && r.besideTop < -40 && r.besideLeft >= 11 && r.besideLeft < 13.5,
    `a table that doesn't fit beside a float starts under it; one that fits stays beside it, at its wrap distance (${JSON.stringify(r)})`);
}
// A header holding a logo wrapped square beside its paragraph, reaching into the body's top margin, and a floating table
// placed 10pt above the paragraph after it (mkfloatroom.mjs): Word starts the body at its top margin (its first line's
// text 110px down in Word's PDF; it had been under the logo, 145px) and draws the table 13.3px above that paragraph.
{
  await open('header-float.docx');
  const r = await page.evaluate(() => {
    const s0 = document.querySelector('.odx-sheet').getBoundingClientRect();
    const hf = document.querySelector('.odx-sheet .odx-hf');
    const t = hf?.querySelector('table');
    const after = t?.nextElementSibling;
    const first = document.querySelector('.odx-content > p');
    const r1 = (v) => Math.round(v * 10) / 10;
    return { body: r1(first.getBoundingClientRect().top - s0.top), table: t && after ? r1(t.getBoundingClientRect().top - after.getBoundingClientRect().top) : null };
  });
  check(Math.abs(r.body - 96) < 1.5 && r.table != null && Math.abs(r.table + 13.33) < 1.5,
    `a header's floats don't push the body down; its floating table placed above its paragraph is drawn there (${JSON.stringify(r)})`);
}
// A section's closing paragraph holding only a picture, its top above the footer's room and its foot in it
// (mkfloatroom.mjs): Word keeps it on its page, two pages in all (it had gone on alone: a blank page, three).
{
  await open('sect-hang.docx');
  const sheets = await page.evaluate(() => document.querySelectorAll('.odx-sheet').length);
  check(sheets === 2, `a section's closing paragraph of only a picture stays on its page, as an empty one does (${sheets} pages; Word 2)`);
}
// Where the body starts under a header (mkhfbody.mjs; Calibri 11pt, the header 24px down, Word's PDFs of documents made
// like these): two pictures with tabs between them, a line as tall as the taller picture (body at 110px; the tabs had
// added their descent, 4.4px); a paragraph closing with a line break over a 6pt mark, its empty line the mark's (69.6px;
// it had been lost); a rule of no height wrapped top and bottom 13pt below its paragraph's top, no text kept clear of it
// (77.7px; it had pushed the body 13px down); a positional tab, a tight logo across the line, spaces and a positional
// tab, the spaces' line and the tab's below the logo (155.6px; the body had started at the top margin, 96px).
for (const [f, want, what] of [['hf-pic-tabs.docx', 110, 'tabs beside header pictures size no line'],
  ['hf-end-break.docx', 69.6, "a header paragraph closing with a line break keeps its mark's empty line"],
  ['hf-zero-rule.docx', 77.7, 'a header rule of no height wrapped top and bottom keeps no text clear of it'],
  ['hf-ptab-logo.docx', 155.6, 'a header line of positional tabs round a logo across it goes on below the logo']]) {
  await open(f);
  const body = await page.evaluate(() => {
    const s0 = document.querySelector('.odx-sheet').getBoundingClientRect();
    return Math.round((document.querySelector('.odx-content > p').getBoundingClientRect().top - s0.top) * 10) / 10;
  });
  check(Math.abs(body - want) < 1, `${what} (body at ${body}px; Word ${want})`);
}
// A header line of positional tabs (mkhfbody.mjs hf-ptabs): the centred one's text centred between the margins, the
// right one's ending at the right margin, as Word's PDF has them (408.2px and 720.3px from the page's left).
{
  await open('hf-ptabs.docx');
  const r = await page.evaluate(() => {
    const s0 = document.querySelector('.odx-sheet').getBoundingClientRect();
    const spans = [...document.querySelectorAll('.odx-sheet .odx-hf.odx-header span.odx-t')];
    const at = (t) => { const s = spans.find((x) => x.textContent === t); if (!s) return null; const q = s.getBoundingClientRect(); return { mid: Math.round((q.left + q.right) / 2 - s0.left), right: Math.round(q.right - s0.left), top: Math.round(q.top - s0.top) }; };
    return { left: at('Left words'), centre: at('Centre words'), right: at('Right words') };
  });
  check(r.centre && r.right && Math.abs(r.centre.mid - 408) <= 1 && Math.abs(r.right.right - 720) <= 1 && r.right.top === r.left.top,
    `positional tabs set their text centred and at the right margin, on one line (${JSON.stringify(r)})`);
}
// The hyphenation zone (Word 2010's layout: a word left whole where wrapping it leaves no more than the zone at the
// line's end; 1in leaves almost none hyphenated) and the consecutive-hyphen limit (here one line in a row), against the
// same text with neither (mkhyphrules.mjs). Where the browser hyphenates itself these are its own, not checked.
const hyLines = async (f) => {
  await open(f);
  await page.waitForFunction(() => document.querySelectorAll('.odx-content .odx-shy, .odx-content p.odx-hyph').length > 0, null, { timeout: 5000 }).catch(() => {});
  await page.waitForTimeout(400);
  return page.evaluate(() => {
    if (document.querySelector('.odx-content p.odx-hyph')) return null;
    const rectAt = (t, i) => { const r = document.createRange(); r.setStart(t, i); r.setEnd(t, i + 1); return r.getClientRects()[0]; };
    const near = (s, dir) => {
      const w = document.createTreeWalker(s.closest('p'), NodeFilter.SHOW_TEXT);
      w.currentNode = s.firstChild;
      for (let t = dir > 0 ? w.nextNode() : w.previousNode(); t; t = dir > 0 ? w.nextNode() : w.previousNode()) if (t.data && !t.parentElement.closest('.odx-shy')) return rectAt(t, dir > 0 ? 0 : t.data.length - 1);
      return null;
    };
    let breaks = 0, run = 0, most = 0, last = null;
    for (const s of document.querySelectorAll('.odx-content .odx-shy')) {
      const a = near(s, -1), b = near(s, 1);
      if (!a || !b || Math.abs(a.top - b.top) < a.height / 2) continue;
      breaks++;
      run = last && Math.abs(last.top - a.top) < a.height / 2 ? run + 1 : 1;
      most = Math.max(most, run);
      last = b;
    }
    return { breaks, most };
  });
};
const hyFree = await hyLines('hyphenation-free.docx');
const hyZone = await hyLines('hyphenation-zone.docx');
const hyLimit = await hyLines('hyphenation-limit.docx');
check(!hyFree || (hyFree.most > 1 && hyZone.breaks < hyFree.breaks / 3 && hyLimit.most === 1 && hyLimit.breaks > 0),
  `hyphenation zone and consecutive-hyphen limit (${JSON.stringify({ hyFree, hyZone, hyLimit })})`);
// Settled, not drawn again for good: the limit's justified paragraphs were measured in a way the editor read as a change
// in the page, which redrew them, which measured them again (every 85ms, as long as the document stayed open).
{
  const redraws = await page.evaluate(() => new Promise((res) => {
    let first = document.querySelector('.odx-content p');
    let n = 0;
    const tick = () => { const p = document.querySelector('.odx-content p'); if (p !== first) { first = p; n++; } };
    const id = setInterval(tick, 20);
    setTimeout(() => { clearInterval(id); res(n); }, 1500);
  }));
  check(redraws === 0, `a hyphenated document with the consecutive-hyphen limit settles (first paragraph redrawn ${redraws} times in 1.5s)`);
}
// The same in Firefox, which hyphenates itself (hyphens:auto; the words left whole marked hyphens:manual): lines that
// end inside a word, its letters on two lines. (Where Playwright's Firefox is installed: CI installs only Chromium.)
{
  let ff = null;
  try { ff = await firefox.launch(); } catch { console.log('skip hyphenation zone and limit in Firefox (not installed)'); }
  if (ff) {
    const fp = await ff.newPage({ viewport: { width: 1300, height: 950 } });
    await fp.goto(`http://localhost:${server.address().port}/`);
    await fp.waitForSelector('.odx-content p');
    const ffLines = async (f) => {
      await fp.setInputFiles('input[type=file]', `test/fixtures/${f}`);
      await fp.waitForFunction((n) => window.odx?.fileName === n, f);
      await fp.waitForTimeout(1500);
      return fp.evaluate(() => {
        const same = (a, b) => Math.abs(a.top - b.top) < a.height / 2;
        let breaks = 0, most = 0;
        for (const p of document.querySelectorAll('.odx-content > p')) {
          let run = 0, last = null;
          const walk = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
          for (let t = walk.nextNode(); t; t = walk.nextNode()) {
            const rect = (i) => { const r = document.createRange(); r.setStart(t, i); r.setEnd(t, i + 1); return r.getClientRects()[0]; };
            for (const m of t.data.matchAll(/[A-Za-z]+/g)) {
              for (let k = 1, prev = rect(m.index); k < m[0].length; k++) {
                const r = rect(m.index + k);
                if (prev && r && !same(prev, r)) { breaks++; run = last && same(last, prev) ? run + 1 : 1; most = Math.max(most, run); last = r; break; }
                prev = r;
              }
            }
          }
        }
        return { native: !!document.querySelector('.odx-content p.odx-hyph'), breaks, most };
      });
    };
    const fFree = await ffLines('hyphenation-free.docx');
    const fZone = await ffLines('hyphenation-zone.docx');
    const fLimit = await ffLines('hyphenation-limit.docx');
    check(fFree.native && fFree.most > 1 && fZone.breaks < fFree.breaks / 3 && fLimit.most === 1 && fLimit.breaks > 0,
      `hyphenation zone and consecutive-hyphen limit in Firefox (${JSON.stringify({ fFree, fZone, fLimit })})`);
    await ff.close();
  }
}

// Bold over the cloud SemiBold and Demi families (mksemibold.mjs): Word fakes it at the Regular's widths, so each
// paragraph's lines end on the words of Word's PDF of it, the Bold ones as the Regular (widths.js fakeBold). With the
// open-metric-fonts stand-ins from jsDelivr: skipped where they don't load (no network).
const WORD_LINES = [
  'by of waltzes grapes.', 'by of waltzes grapes.', 'by of waltzes grapes.', 'by of waltzes grapes.',
  'stored 0123456789 black and mixed grapes.', 'rewritten jugs; jaunty grapes.', 'were dozen quack, grapes.',
  'procedures with vixens boxes grapes.',
];
// (Read until two readings half a second apart agree: the stand-ins' files, one per alphabet, can arrive late, and
// the lines meanwhile are set in the fallback font.)
const lineEnds = (pg) => pg.evaluate(async () => {
  const read = () => [...document.querySelectorAll('.odx-content > p')].map((p) => {
    const ends = [];
    let top = null, word = '';
    const walk = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let t = walk.nextNode(); t; t = walk.nextNode()) {
      for (const m of t.data.matchAll(/\S+/g)) {
        const r = document.createRange();
        r.setStart(t, m.index);
        r.setEnd(t, m.index + 1);
        const at = r.getClientRects()[0]?.top;
        if (top != null && at - top > 3) ends.push(word);
        top = at;
        word = m[0];
      }
    }
    return [...ends, word].join(' ');
  });
  let last = null;
  for (let i = 0; i < 30; i++) {
    await document.fonts.ready;
    await new Promise((r) => setTimeout(r, 500));
    const faces = [...document.fonts].filter((f) => /odx-omf/.test(f.family) && f.style === 'normal' && f.status === 'loaded');
    const now = new Set(faces.map((f) => f.family.replace(/"/g, ''))).size >= 4 ? JSON.stringify(read()) : null;
    if (now && now === last) return JSON.parse(now);
    last = now;
  }
  return null;
});
// (Word's words on Windows, where the browser sets the stand-ins at their design widths as Word does. Elsewhere, whose
// fonts can be set differently, the Bold paragraphs' lines end where the Regular's and the Italic's do.)
const sameAsWord = (ends) => (process.platform === 'win32' ? JSON.stringify(ends) === JSON.stringify(WORD_LINES)
  : ends[1] === ends[0] && ends[3] === ends[2]);
await open('semibold-bold.docx');
{
  const ends = await lineEnds(page);
  if (!ends) console.log('skip SemiBold with Bold (the open-metric-fonts stand-ins did not load, or the lines did not settle)');
  else check(sameAsWord(ends), `SemiBold with Bold: lines end as in Word's PDF (${JSON.stringify(ends)})`);
  let ff = null;
  try { ff = await firefox.launch(); } catch { console.log('skip SemiBold with Bold in Firefox (not installed)'); }
  if (ff) {
    const fp = await ff.newPage({ viewport: { width: 1300, height: 950 } });
    await fp.goto(`http://localhost:${server.address().port}/`);
    await fp.waitForSelector('.odx-content p');
    await fp.setInputFiles('input[type=file]', 'test/fixtures/semibold-bold.docx');
    await fp.waitForFunction(() => window.odx?.fileName === 'semibold-bold.docx');
    await fp.waitForTimeout(2500);
    const fends = await lineEnds(fp);
    if (!fends) console.log('skip SemiBold with Bold in Firefox (the open-metric-fonts stand-ins did not load, or the lines did not settle)');
    else check(sameAsWord(fends), `SemiBold with Bold in Firefox: lines end as in Word's PDF (${JSON.stringify(fends)})`);
    await ff.close();
  }
}

// Text spaced under a single line (mksubsingle.mjs): each paragraph that multiple of the font's single line, however tall
// its text. Word's PDF: "Single A" to "Single C" 276.8px; the 16pt headings at 0.5 lines under 12pt before, 29.0px apart.
// (At the 1.15em floor of a paragraph, 353.6px and 40.5px.)
await open('sub-single.docx');
{
  const tops = await page.evaluate(() => Object.fromEntries([...document.querySelectorAll('.odx-content p')].map((p) => [p.textContent, p.getBoundingClientRect().top])));
  const span = tops['Single C'] - tops['Single A'], heads = tops['Head three'] - tops['Head two'];
  check(Math.abs(span - 276.8) < 1.5 && Math.abs(heads - 29.0) < 1, `text under a single line: as tall as Word's lines (${span?.toFixed(2)}px, headings ${heads?.toFixed(2)}px)`);
}

// Aptos Narrow (mkaptosnarrow.mjs): drawn in the open-metric-fonts stand-in, at Aptos Narrow's widths, ahead of the
// editor's Arial Narrow (6.5% narrower). Word's PDF: the lines end on "limited", "keeping", "year." and "as", "two",
// "period.". Skipped where the stand-in doesn't load (no network); elsewhere than Windows, the font's order only.
await open('aptos-narrow.docx');
{
  const r = await page.evaluate(async () => {
    const span = document.querySelector('.odx-content p span');
    const family = span ? getComputedStyle(span).fontFamily : '';
    let loaded = false;
    for (let i = 0; i < 30 && !loaded; i++) {
      await document.fonts.ready;
      await new Promise((res) => setTimeout(res, 500));
      loaded = [...document.fonts].some((f) => f.family.replace(/"/g, '') === 'odx-omf Aptos Narrow' && f.status === 'loaded');
    }
    await new Promise((res) => setTimeout(res, 1000));
    const ends = [];
    for (const p of document.querySelectorAll('.odx-content > p')) {
      let top = null, word = '';
      const walk = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      for (let t = walk.nextNode(); t; t = walk.nextNode()) {
        for (const m of t.data.matchAll(/\S+/g)) {
          const rg = document.createRange();
          rg.setStart(t, m.index);
          rg.setEnd(t, m.index + 1);
          const at = rg.getClientRects()[0]?.top;
          if (top != null && at - top > 3) ends.push(word);
          top = at;
          word = m[0];
        }
      }
      ends.push(word);
    }
    return { loaded, ahead: family.indexOf('odx-omf Aptos Narrow') >= 0 && family.indexOf('odx-omf Aptos Narrow') < family.indexOf('odx-aptos-narrow'), ends };
  });
  check(r.ahead, `Aptos Narrow: the open-metric-fonts stand-in ahead of Arial Narrow (${JSON.stringify(r)})`);
  if (!r.loaded) console.log('skip Aptos Narrow line ends (the open-metric-fonts stand-in did not load)');
  else if (process.platform === 'win32') check(r.ends.join(' ') === 'limited keeping year. as two period.', `Aptos Narrow: lines end as in Word's PDF (${r.ends.join(' ')})`);
}

// Mailings: recipients from the host, Preview Results (values in place of the fields on screen, the document keeping
// its fields), the next record, Insert Merge Field, a CSV file of recipients, and Finish & Merge's merged document.
await open('mailmerge.docx');
await page.evaluate(() => window.odx.setMergeData([{ FirstName: 'Ada', 'Last Name': 'Lovelace', City: 'London', Amount: '1234.5' }, { FirstName: 'Alan', 'Last Name': 'Turing', City: 'Wilmslow', Amount: '7' }]));
await page.click('.odx-rtab:text-is("Mailings")');
await page.click('button[title^="Preview results"]');
await page.waitForTimeout(400);
const firstText = () => page.evaluate(() => document.querySelector('.odx-content p').innerText.trim());
const pv1 = await firstText();
await page.click('button[title="Next record"]');
await page.waitForTimeout(300);
const pv2 = await firstText();
const pvStatus = await page.evaluate(() => [...document.querySelectorAll('.odx-panel:not([hidden]) .odx-cap')].map((e) => e.textContent).join('|'));
const pvSaved = strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/document.xml']);
await page.click('button[title^="Preview results"]');
await page.waitForTimeout(300);
const pvOff = await firstText();
// Insert Merge Field at the end of the first paragraph.
await page.click('.odx-content p >> nth=0');
await page.keyboard.press('End');
await page.keyboard.type(' ');
await page.click('button[title="Insert merge field"]');
await page.click('.odx-pop .odx-item:has-text("City")');
await page.waitForTimeout(300);
const inserted = await firstText();
// A CSV file of recipients.
await page.setInputFiles('input[accept=".csv,text/csv"]', { name: 'people.csv', mimeType: 'text/csv', buffer: Buffer.from('FirstName,Last Name,City\r\n"Grace","Hopper, Rear Adm.",Arlington\r\n') });
await page.waitForTimeout(300);
const csvCount = await page.evaluate(() => window.odx.view.odx.session.merge.records.length);
const [download] = await Promise.all([
  page.waitForEvent('download', { timeout: 5000 }),
  (async () => { await page.click('button[title^="Finish & Merge"]'); await page.click('.odx-pop .odx-item:has-text("Edit individual documents")'); })(),
]);
const mmXml = strFromU8(unzipSync(new Uint8Array(fs.readFileSync(await download.path())))['word/document.xml']);
check(pv1 === 'Ada LOVELACE' && pv2 === 'Alan TURING' && /Record 2 of 2/.test(pvStatus) && /MERGEFIELD FirstName/.test(pvSaved) && pvOff === '«FirstName» «Last Name»'
  && inserted === '«FirstName» «Last Name» «City»' && csvCount === 1 && /Grace/.test(mmXml) && /HOPPER, REAR ADM\./.test(mmXml) && !/MERGEFIELD/.test(mmXml) && /download/.test(download.suggestedFilename() + 'download'),
  `mail merge: preview, next record, insert field, CSV recipients, merged document (${JSON.stringify({ pv1, pv2, pvStatus, pvOff, inserted, csvCount, file: download.suggestedFilename() })})`);
// Preview Results with IF fields and a header's and footer's merge fields (mergeif.docx; Word's merge gives these).
await open('mergeif.docx');
await page.evaluate(() => window.odx.setMergeData([{ Name: 'Ada', Amount: '150', City: 'New York' }, { Name: 'alan', Amount: '50', City: 'Boston' }]));
await page.click('.odx-rtab:text-is("Mailings")');
await page.click('button[title^="Preview results"]');
await page.waitForTimeout(500);
const ifShown = () => page.evaluate(() => ({
  lines: [...document.querySelectorAll('.odx-content p')].slice(0, 2).map((p) => p.innerText.trim()).concat(document.querySelectorAll('.odx-content p')[15]?.innerText.trim()),
  header: document.querySelector('.odx-header')?.innerText.trim(),
  footer: document.querySelector('.odx-footer')?.innerText.trim(),
}));
const if1 = await ifShown();
await page.click('button[title="Next record"]');
await page.waitForTimeout(500);
const if2 = await ifShown();
await page.click('button[title^="Preview results"]');
await page.waitForTimeout(500);
const ifOff = await ifShown();
check(JSON.stringify(if1) === JSON.stringify({ lines: ['gt: big', 'lt: not under', 'nested: big'], header: 'Statement for Ada', footer: 'City: New York / big' })
  && JSON.stringify(if2) === JSON.stringify({ lines: ['gt: small', 'lt: under', 'nested: mid'], header: 'Statement for alan', footer: 'City: Boston / small' })
  && ifOff.header === 'Statement for «Name»' && ifOff.lines[0] === 'gt: small',
  `mail merge preview: IF fields, header and footer per record (${JSON.stringify({ if1, if2, ifOff })})`);
// NEXTIF and NEXT in the preview (mergenext.docx; Word's preview shows these): the fields after them from the records
// after the one shown, the header from the record the body ended on.
await open('mergenext.docx');
await page.evaluate(() => window.odx.setMergeData([['Ada', 'x', '5'], ['Bob', 'x', '20'], ['Cy', 'y', '30'], ['Di', 'x', '2'], ['Ed', 'y', '40']].map(([Name, Group, Amount]) => ({ Name, Group, Amount }))));
await page.click('.odx-rtab:text-is("Mailings")');
await page.click('button[title^="Preview results"]');
await page.waitForTimeout(500);
const nextShown = () => page.evaluate(() => ({
  lines: [0, 2, 4].map((i) => document.querySelectorAll('.odx-content p')[i]?.innerText.replace(/\s+/g, ' ').trim()),
  header: document.querySelector('.odx-header')?.innerText.trim(),
}));
const nx1 = await nextShown();
await page.click('button[title="Next record"]');
await page.waitForTimeout(500);
const nx2 = await nextShown();
await page.click('button[title^="Preview results"]');
check(JSON.stringify(nx1) === JSON.stringify({ lines: ['a: 1 Ada', 'b: 2 Bob', 'c: 3 Cy'], header: 'H: Cy' })
  && JSON.stringify(nx2) === JSON.stringify({ lines: ['a: 2 Bob', 'b: 3 Cy', 'c: 4 Di'], header: 'H: Di' }),
  `mail merge preview: NEXTIF and NEXT moving on to the next records (${JSON.stringify({ nx1, nx2 })})`);
await page.click('.odx-rtab:text-is("Home")');

// Insert > Shapes: picked from the keyboard, a rectangle in line with the text; with the mouse, drawn on the page by
// dragging, as in Word: an arrow dragged up and to the right (its end where the drag ended), an oval floating in front
// of the text where it was dragged out, one Undo taking it away; Esc stops drawing. Theme-styled as Word's (accent 1,
// darker outline), and saved as Word writes them.
await page.evaluate(() => window.odx.newDocument({ share: false }));
await page.click('.odx-content p');
await page.keyboard.type('A paragraph for the shapes to be drawn over.');
await page.click('.odx-rtab:text-is("Insert")');
await page.click('button[title^="Shapes"]');
await page.focus('.odx-pop button[title="Rectangle"]');
await page.keyboard.press('Enter');
await page.waitForTimeout(300);
// (Each shape's own box: less the room drawn round a floating one's outline or arrowheads, reader.js padPlace's --odx-pad.)
const imgs = () => page.evaluate(() => [...document.querySelectorAll('.odx-content img:not(.ProseMirror-separator)')].map((i) => { const r = i.getBoundingClientRect(); const pad = Number(/--odx-pad:\s*([\d.]+)px/.exec(i.getAttribute('style') || '')?.[1]) || 0; return { w: i.width - 2 * pad, h: i.height - 2 * pad, x: r.left + pad, y: r.top + pad, svg: decodeURIComponent(i.src.split(',')[1] || '') }; }));
const p0 = await page.evaluate(() => { const r = document.querySelector('.odx-content p').getBoundingClientRect(); return { x: r.left, y: r.bottom }; });
const drawOn = async (title, x0, y0, x1, y1) => {
  await page.click('button[title^="Shapes"]');
  await page.click(`.odx-pop button[title="${title}"]`);
  await page.mouse.move(x0, y0);
  await page.mouse.down();
  await page.mouse.move((x0 + x1) / 2, (y0 + y1) / 2);
  await page.mouse.move(x1, y1);
  await page.mouse.up();
  await page.waitForTimeout(400);
};
await drawOn('Arrow', p0.x + 40, p0.y + 120, p0.x + 190, p0.y + 60);
await drawOn('Oval', p0.x + 250, p0.y + 140, p0.x + 370, p0.y + 220);
const shp = await imgs();
// (Each drawn one goes in at the start of its paragraph, before the rectangle: found by size.)
const oval = shp.find((x) => x.w === 120) || {};
const rect = shp.find((x) => x.w === 96) || {};
const arrowImg = shp.find((x) => x.w === 150) || { svg: '' };
const shpXml = strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/document.xml']);
const anchors = [...shpXml.matchAll(/<wp:anchor\b[\s\S]*?<\/wp:anchor>/g)].map((m) => m[0]);
await page.keyboard.press('Control+z');
await page.waitForTimeout(300);
const afterUndo = (await imgs()).length;
await page.click('button[title^="Shapes"]');
await page.click('.odx-pop button[title="Rectangle"]');
const drawingOn = await page.evaluate(() => !!document.querySelector('.odx-draw'));
await page.keyboard.press('Escape');
await page.waitForTimeout(100);
const drawingOff = await page.evaluate(() => !document.querySelector('.odx-draw'));
check(shp.length === 3 && /fill="#156082"/.test(rect.svg || '') && /<a:prstGeom prst="rect">/.test(shpXml) && /<a:fillRef idx="1"><a:schemeClr val="accent1"\/>/.test(shpXml)
  && /<a:prstGeom prst="line">[\s\S]*?<a:tailEnd type="triangle"\/>/.test(shpXml) && /<path d="M[^"]*Z" fill="#/.test(arrowImg.svg)
  && anchors.length === 2 && anchors.every((x) => /<wp:wrapNone\/>/.test(x) && /behindDoc="0"/.test(x) && /<wp:positionH relativeFrom="column"><wp:posOffset>-?\d+</.test(x) && /<wp:positionV relativeFrom="paragraph"><wp:posOffset>-?\d+</.test(x))
  && anchors.some((x) => /<a:xfrm flipV="1">[\s\S]*?prst="line"/.test(x) && /<wp:extent cx="1428750" cy="571500"\/>/.test(x))
  && anchors.some((x) => /prst="ellipse"/.test(x)) && Math.abs(arrowImg.x - (p0.x + 40)) <= 1 && Math.abs(arrowImg.y - (p0.y + 60)) <= 1 && Math.abs(oval.x - (p0.x + 250)) <= 1 && Math.abs(oval.y - (p0.y + 140)) <= 1 && oval.w === 120 && oval.h === 80
  && afterUndo === 2 && drawingOn && drawingOff,
  `Insert > Shapes: a rectangle from the keyboard, an arrow and an oval drawn by dragging, saved as Word's (${JSON.stringify({ shp: shp.map((s) => [s.w, s.h, Math.round(s.x - p0.x), Math.round(s.y - p0.y)]), anchors: anchors.length, afterUndo, drawingOn, drawingOff })})`);
await page.click('.odx-rtab:text-is("Home")');

// Insert > WordArt: fifteen styles, each previewed in its effects; one picked puts "Your text here" in a floating text box,
// open for typing with its text selected; what's typed replaces it, the box widening to fit, saved with the effects.
await page.evaluate(() => window.odx.newDocument({ share: false }));
await page.click('.odx-content p');
await page.keyboard.type('A paragraph under the WordArt.');
await page.click('.odx-rtab:text-is("Insert")');
await page.click('button[title^="WordArt"]');
const waGallery = await page.evaluate(() => [...document.querySelectorAll('.odx-pop .odx-wa')].map((b) => ({ name: b.title, shadow: getComputedStyle(b.firstChild).textShadow })));
await page.click('.odx-pop .odx-wa[title="Fill: White; Outline: Accent 1; Glow: Accent 1"]');
await page.waitForTimeout(400);
const waEditing = await page.evaluate(() => !!document.querySelector('.odx-txbxedit') && String(getSelection()));
await page.keyboard.type('Hello WordArt, a little longer');
await page.keyboard.press('Escape');
await page.waitForTimeout(300);
const waImg = await page.evaluate(() => { const i = document.querySelector('.odx-content img:not(.ProseMirror-separator)'); return i && { w: i.width, abs: getComputedStyle(i).position }; });
const waXml = strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/document.xml']);
check(waGallery.length === 15 && waGallery[0].shadow !== 'none' && waEditing === 'Your text here' && waImg && waImg.w > 295 && waImg.abs === 'absolute'
  && /<wp:anchor\b[\s\S]*?<wp:wrapNone\/>[\s\S]*?<w:t[^>]*>Hello WordArt, a little longer<\/w:t>/.test(waXml) && /<w14:glow\b[^>]*\bw14:rad="38100"/.test(waXml) && /wrap="none"[\s\S]*?<a:spAutoFit\/>/.test(waXml) && !/Your text here/.test(waXml),
  `Insert > WordArt: picked, typed over, widened to fit, saved with its effects (${JSON.stringify({ styles: waGallery.length, waEditing, waImg })})`);
// WordArt wider than the column wraps there, as Word's (wordart-leftovers/wide.docx: its box no wider than the column,
// taller for the lines it wraps onto), while it's typed and once drawn; the gallery's Soft Bevel is drawn as a lit
// relief of the letters (an SVG lighting filter), as Word prints it.
await page.evaluate(() => window.odx.newDocument({ share: false }));
await page.click('.odx-content p');
await page.click('.odx-rtab:text-is("Insert")');
await page.click('button[title^="WordArt"]');
await page.click('.odx-pop .odx-wa[title="Fill: Accent 4; Soft Bevel"]');
await page.waitForTimeout(400);
await page.keyboard.type('This WordArt is far too long to fit on one line of its column');
const waTyping = await page.evaluate(() => { const e = document.querySelector('.odx-txbxedit'); return e && { w: e.getBoundingClientRect().width, h: e.getBoundingClientRect().height }; });
await page.keyboard.press('Escape');
await page.waitForTimeout(500);
const waWide = await page.evaluate(() => {
  const v = document.querySelector('.odx-content');
  const cs = getComputedStyle(v);
  const col = v.clientWidth - parseFloat(cs.getPropertyValue('--odx-page-left')) - parseFloat(cs.getPropertyValue('--odx-page-right'));
  const i = document.querySelector('.odx-content img:not(.ProseMirror-separator)');
  return i && { col, w: i.getBoundingClientRect().width, h: i.getBoundingClientRect().height, lit: /feDiffuseLighting/.test(decodeURIComponent(i.getAttribute('src'))) };
});
check(waTyping && waWide && waTyping.w <= waWide.col + 1 && waTyping.h > 100 && waWide.w <= waWide.col + 1 && waWide.w > waWide.col * 0.6 && waWide.h > 100 && waWide.lit,
  `WordArt wider than the column wraps at it; a bevel is lit (${JSON.stringify({ waTyping, waWide })})`);
await page.click('.odx-rtab:text-is("Home")');

// The stand-in fonts: their stylesheet linked from fonts.base (here the demo's own copy), and the Aptos stand-in (Morrow
// Sans's letters at Aptos's fitted widths) loaded from its file when asked for.
const fontsLoad = await page.evaluate(async () => {
  const link = document.querySelector('link[data-odx-fonts]');
  await document.fonts.load('16px odx-aptos', 'Image Test');
  const faces = [...document.fonts].filter((f) => f.family.replace(/["']/g, '') === 'odx-aptos' && f.status === 'loaded');
  return { href: link?.href.replace(location.origin, ''), rules: link?.sheet?.cssRules.length || 0, loaded: faces.length };
});
check(fontsLoad.href === '/dist/fonts/fonts.css' && fontsLoad.rules > 20 && fontsLoad.loaded >= 1, `stand-in fonts linked from fonts.base and loaded when used (${JSON.stringify(fontsLoad)})`);

// Shapes' text in the stand-in fonts: the SmartArt (theme font Aptos, missing here and on CI) shown with the Aptos
// stand-in embedded in its picture, the node keeping the picture as drawn.
await open('smartart.docx');
await page.waitForFunction(() => [...document.querySelectorAll('.odx-content img')].some((i) => (i.getAttribute('src') || '').includes('%40font-face')), null, { timeout: 8000 }).catch(() => {});
const svgFonts = await page.evaluate(() => {
  const img = [...document.querySelectorAll('.odx-content img')].find((i) => (i.getAttribute('src') || '').startsWith('data:image/svg+xml'));
  const shown = decodeURIComponent(img?.getAttribute('src').split(',')[1] || '');
  let nodeSrc = '';
  window.odx.view.state.doc.descendants((n) => { if (!nodeSrc && n.type.name === 'image') nodeSrc = n.attrs.src || ''; });
  return { embedded: /@font-face\{font-family:"odx-aptos"[^}]*src:url\(data:/.test(shown), named: /odx-aptos/.test(shown), nodeClean: !!nodeSrc && !nodeSrc.includes('font-face') };
});
check(svgFonts.embedded && svgFonts.named && svgFonts.nodeClean, `a shape's text drawn in the stand-in font embedded in its picture, the document unchanged (${JSON.stringify(svgFonts)})`);

// A picture turned with Alt+Right (15 degrees each), as Word's: its pic:spPr a:xfrm rot and the room round it written
// as Word writes them, drawn turned; at 90 degrees it takes its turned box's room (taller than wide, no effect extent).
await open('rich.docx');
const picAt = await page.evaluate(() => { let at = -1; window.odx.view.state.doc.descendants((n, pos) => { if (at < 0 && n.type.name === 'image' && /<pic:pic\b/.test(n.attrs.xml)) at = pos; }); const r = window.odx.view.nodeDOM(at).getBoundingClientRect(); return { at, x: r.x + r.width / 2, y: r.y + r.height / 2 }; });
await page.mouse.click(picAt.x, picAt.y);
await page.waitForTimeout(300);
const picTurn = async (n) => {
  for (let i = 0; i < n; i++) { await page.keyboard.press('Alt+ArrowRight'); await page.waitForTimeout(120); }
  return page.evaluate((at) => { const n = window.odx.view.state.doc.nodeAt(at); const img = window.odx.view.nodeDOM(at); const r = img.getBoundingClientRect(); return { rot: /<pic:spPr\b[^>]*>\s*<a:xfrm\b[^>]*\brot="(\d+)"/.exec(n.attrs.xml)?.[1], fx: /<wp:effectExtent\b[^>]*\bl="(\d+)"[^>]*\bt="(\d+)"/.exec(n.attrs.xml)?.slice(1).join(','), turned: /rotate\((\d+)deg\)/.exec(img.style.transform)?.[1], handle: !!document.querySelector('.odx-imgsize-turns'), wide: Math.round(r.width), tall: Math.round(r.height) }; }, picAt.at);
};
const pic30 = await picTurn(2);
const pic90 = await picTurn(4);
check(pic30.rot === '1800000' && pic30.fx === '66675,247650' && pic30.turned === '30' && pic30.handle && pic90.rot === '5400000' && pic90.fx === '0,0' && pic90.turned === '90' && pic90.tall > pic90.wide,
  `a picture turned with Alt+Right: written and drawn as Word does (${JSON.stringify({ pic30, pic90 })})`);

// Long footnotes (made in Word: four notes of about twelve lines each): a paragraph whose note won't fit below it keeps
// its lines above the reference on the page, the reference's line going on with its note, as Word lays it out.
await open('footnotes-long.docx');
const fnLong = await page.evaluate(() => {
  const sheets = [...document.querySelectorAll('.odx-sheet')].map((s) => s.getBoundingClientRect());
  const sheetOf = (y) => sheets.findIndex((r) => y >= r.top && y < r.bottom);
  const p4 = [...document.querySelectorAll('.odx-content > p')].find((p) => p.textContent.startsWith('Text 4.'));
  const rg = document.createRange();
  rg.selectNodeContents(p4);
  const rects = [...rg.getClientRects()].filter((r) => r.height > 0);
  const ref = p4.querySelector('[data-kind="footnoteReference"]').getBoundingClientRect();
  return { first: sheetOf(rects[0].top), ref: sheetOf((ref.top + ref.bottom) / 2), pages: sheets.length };
});
check(fnLong.first === 0 && fnLong.ref === 1, `a long note's paragraph split above its reference, as Word (${JSON.stringify(fnLong)})`);

// A footnote inserted in 14pt Georgia text (notes-georgia14.docx, mknotes.mjs), typed in, then typed after: the caret
// stays after the reference, the note is drawn under the text of its reference's page (the text making room), the
// typed text after it is the body's 14pt Georgia, and the note keeps the space after its number ("¹Typed": collapsed,
// the browser dropped it with the first letter).
await open('notes-georgia14.docx');
const g14 = () => page.evaluate(() => {
  const v = window.odx.view;
  const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
  const sheetOf = (y) => sheets.findIndex((r) => y >= r.top && y < r.bottom);
  const ref = document.querySelector('.odx-content [data-kind="footnoteReference"]');
  const rr = ref?.getBoundingClientRect();
  const s0 = rr ? sheetOf((rr.top + rr.bottom) / 2) : 0;
  const lines = [];
  const w = document.createTreeWalker(document.querySelector('.odx-content'), NodeFilter.SHOW_TEXT);
  for (let n; (n = w.nextNode());) { const rg = document.createRange(); rg.selectNodeContents(n); lines.push(...[...rg.getClientRects()].filter((r) => r.height > 0 && sheetOf(r.top) === s0)); }
  const note = document.querySelector('.odx-note[data-kind="footnote"]');
  const nr = note?.getBoundingClientRect();
  const ed = document.querySelector('.odx-noteedit');
  const nv = v.odx.hfView;
  const c = v.coordsAtPos(v.state.selection.from);
  const top = sheets[s0].top;
  return {
    sel: v.state.selection.from, refAt: (() => { let at = null; v.state.doc.descendants((n, pos) => { if (n.attrs?.kind === 'footnoteReference') at = pos; }); return at; })(),
    before: v.state.doc.textBetween(Math.max(0, v.state.selection.from - 12), v.state.selection.from),
    caret: { sheet: sheetOf(c.top), onRefLine: !!rr && c.top < rr.bottom && c.bottom > rr.top, dx: rr ? Math.round(c.left - rr.right) : null },
    ref: rr && { sheet: s0, font: getComputedStyle(ref).fontFamily.split(',')[0], size: parseFloat(getComputedStyle(ref).fontSize) },
    textBottom: Math.round(Math.max(...lines.map((r) => r.bottom)) - top),
    note: nr && { sheet: sheetOf(nr.top), top: Math.round(nr.top - top), bottom: Math.round(nr.bottom - top), text: note.textContent, size: parseFloat(getComputedStyle(note.querySelector('.odx-t') || note).fontSize), lines: new Set([...note.querySelectorAll('.odx-t')].flatMap((t) => { const rg = document.createRange(); rg.selectNodeContents(t); return [...rg.getClientRects()].map((r) => Math.round(r.top)); })).size },
    editor: ed && nv && { top: Math.round(ed.getBoundingClientRect().top - top), atEnd: nv.state.selection.from === nv.state.doc.content.size - 1, focused: nv.hasFocus() },
    focus: v.hasFocus(), pages: sheets.length,
  };
});
await page.evaluate(() => {
  const v = window.odx.view; let at = null;
  v.state.doc.descendants((n, pos) => { if (at == null && n.isText && n.text.startsWith('Paragraph 2.')) at = pos + n.text.indexOf('orchard') + 'orchard'.length; });
  v.focus();
  v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.near(v.state.doc.resolve(at))));
});
const g14Before = await g14();
await page.keyboard.press('Control+Alt+f');
await page.waitForSelector('.odx-noteedit');
await page.waitForTimeout(300);
const g14New = await g14();
await page.keyboard.type('Typed in the note: a note long enough to take three lines of its own at the foot of the page, under the text of the page its reference is on, as Word draws a note of this length there, the text above it making room for it. One more sentence of the note, so that it runs on to a third line of its own.');
await page.waitForTimeout(300);
const g14Typing = await g14();
await page.keyboard.press('Escape');
await page.waitForTimeout(600);
const g14Closed = await g14();
await page.keyboard.type(' typed after');
await page.waitForTimeout(800);
const g14After = await g14();
const g14Run = await page.evaluate(() => { const w = document.createTreeWalker(document.querySelector('.odx-content'), NodeFilter.SHOW_TEXT); for (let n; (n = w.nextNode());) if (n.textContent.startsWith(' typed after')) { const cs = getComputedStyle(n.parentElement); return { font: cs.fontFamily.split(',')[0], size: parseFloat(cs.fontSize), va: cs.verticalAlign }; } return null; });
const g14Bytes = new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save())));
const g14Xml = strFromU8(unzipSync(g14Bytes)['word/footnotes.xml']);
// (The pages as typed against the saved file's pages laid out afresh.)
const g14File = path.join(os.tmpdir(), 'odx-notes-georgia14-typed.docx');
fs.writeFileSync(g14File, g14Bytes);
await page.setInputFiles('input[type=file]', g14File);
await page.waitForFunction(() => window.odx?.fileName === 'odx-notes-georgia14-typed.docx');
await page.waitForTimeout(1200);
const g14Fresh = await g14();
fs.rmSync(g14File, { force: true });
const PAGE_FOOT = 1056 - 96; // (a Letter page's 1in bottom margin)
check(g14New.refAt === g14Before.sel && g14New.sel === g14Before.sel + 1 && g14New.ref.font === 'Georgia' && Math.abs(g14New.ref.size - 18.67 * 0.66) < 0.2 && g14New.editor?.focused && g14New.editor.atEnd && g14New.editor.top === g14New.note.top,
  `a footnote inserted in 14pt Georgia: its reference at the cursor in Georgia, raised, its note open for typing over its drawn copy (${JSON.stringify({ before: g14Before.sel, ...g14New })})`);
check(g14Typing.editor?.atEnd && g14Typing.editor.top === g14Typing.note.top && g14Typing.sel === g14New.sel,
  `typing in the new note: its caret at the end of what's typed, the editor where the note is, the text's cursor still after the reference (${JSON.stringify(g14Typing)})`);
check(g14Closed.focus && g14Closed.sel === g14New.sel && g14Closed.note.sheet === g14Closed.ref.sheet && g14Closed.note.text.startsWith('1 Typed in the note') && Math.abs(g14Closed.note.size - 13.33) < 0.1
  && g14Closed.note.lines === 3 && g14Closed.textBottom < g14Closed.note.top - 10 && g14Closed.note.bottom <= PAGE_FOOT + 1 && g14Closed.textBottom < g14Before.textBottom,
  `the typed note under its reference's page's text, which made room for its three lines; the caret back after the reference (${JSON.stringify(g14Closed)})`);
check(g14After.sel === g14New.sel + 12 && g14After.before.endsWith(' typed after') && g14After.caret.sheet === g14After.ref.sheet && g14After.caret.onRefLine && g14After.caret.dx > 40
  && g14Run?.font === 'Georgia' && Math.abs(g14Run.size - 18.67) < 0.1 && g14Run.va === 'baseline'
  && g14After.note.sheet === g14After.ref.sheet && g14After.textBottom < g14After.note.top - 10 && g14After.note.bottom <= PAGE_FOOT + 1 && g14After.pages === g14Before.pages,
  `typed after the reference: 14pt Georgia on its line, the caret after it, the note still under the page's text (${JSON.stringify({ g14After, g14Run })})`);
check(g14Fresh.textBottom === g14After.textBottom && g14Fresh.note?.top === g14After.note.top && g14Fresh.note.bottom === g14After.note.bottom && g14Fresh.note.lines === g14After.note.lines && g14Fresh.pages === g14After.pages,
  `the pages as typed are the saved file's pages laid out afresh (${JSON.stringify({ typed: [g14After.textBottom, g14After.note], fresh: [g14Fresh.textBottom, g14Fresh.note] })})`);
check(/<w:footnoteRef\/><\/w:r><w:r><w:t xml:space="preserve"> Typed in the note/.test(g14Xml), `the new note saved with the space after its number (${/<w:footnote w:id="1">[\s\S]*?<\/w:footnote>/.exec(g14Xml)?.[0].slice(0, 260)})`);

// Delete Cells ("shift cells left") with Track Changes on: the cell recorded as deleted (w:cellDel), still shown; with it
// off, the cell goes and the row ends in a grid column of its own (w:gridAfter), as Word's dialog does.
await open('rich.docx');
const cellDelOf = async (tracked) => {
  await page.evaluate((on) => { const s = document.querySelector('.odx-tracksel'); s.value = on ? 'on' : 'off'; s.dispatchEvent(new Event('change', { bubbles: true })); }, tracked);
  const td = await page.$('.odx-content td:nth-child(2) p');
  await td.click();
  await page.click('button[title^="Delete cells"]');
  await page.waitForTimeout(300);
  return page.evaluate(async () => {
    let cells = 0, del = 0;
    window.odx.view.state.doc.descendants((n) => { if (n.type.name === 'table_cell' && !n.attrs.phantom) { cells++; if ((n.attrs.tcPr || []).some((t) => t[0] === 'cellDel')) del++; } });
    return { cells, del };
  });
};
const cellsBefore = await page.evaluate(() => { let c = 0; window.odx.view.state.doc.descendants((n) => { if (n.type.name === 'table_cell' && !n.attrs.phantom) c++; }); return c; });
const cdTracked = await cellDelOf(true);
await page.keyboard.press('Control+z');
await page.waitForTimeout(200);
const cdPlain = await cellDelOf(false);
const cdBytes = await page.evaluate(async () => Array.from(await window.odx.save()));
const cdXml = strFromU8(unzipSync(new Uint8Array(cdBytes))['word/document.xml']);
check(cdTracked.cells === cellsBefore && cdTracked.del === 1 && cdPlain.cells === cellsBefore - 1 && cdPlain.del === 0 && /<w:gridAfter w:val="\d+"\/>/.test(cdXml), // (the first row's second cell spans two columns: gridAfter 2)
  `Delete Cells: tracked, the cell recorded as deleted; not tracked, it goes and its row ends in a gridAfter (${JSON.stringify({ cellsBefore, cdTracked, cdPlain })})`);

// A line holding a picture beside text (picfix.js): Word's is the picture (36pt), the font's descent and all of the
// "multiple" extra (picline.docx, Word-made: ..\picline-word). The advance to the next paragraph, each 0 apart.
await open('picline.docx');
const picAdv = await page.evaluate(() => {
  const ps = [...document.querySelectorAll('.odx-content > p')];
  return ps.filter((p) => /picture line/.test(p.textContent)).map((p) => [p.textContent.trim().slice(0, 24), +(ps[ps.indexOf(p) + 1].getBoundingClientRect().top - p.getBoundingClientRect().top).toFixed(2)]);
});
// (descent: Aptos 0.2817em, Calibri 0.2686em (its Windows descent), Times New Roman 0.2163em; extra: (multiple - 1) x line)
const picWant = [48 + 4.507 + 3.093, 48 + 4.507, 48 + 4.507 + 9.766, 48 + 4.507 + 19.531, 48 + 3.939 + 1.418, 48 + 3.939, 48 + 3.461, 48 + 3.461 + 9.199];
check(picAdv.length === 8 && picAdv.every((a, i) => Math.abs(a[1] - picWant[i]) < 0.75), `picture lines as tall as Word's: ${JSON.stringify(picAdv.map((a, i) => [a[0], a[1], +picWant[i].toFixed(2)]))}`);

await open('strict.docx');
const strictText = await page.evaluate(() => document.querySelector('.odx-content td')?.textContent);
check(strictText === 'Half the page wide', 'a Strict document opens');

// The ribbon's layout chosen by the page (mountDocxEditor's `ribbon`, the demo's ?ribbon=): over the user's last
// choice, still switched by the ribbon's own button, and by setRibbon().
const ribbonPage = await browser.newPage();
await ribbonPage.addInitScript(() => { try { localStorage.setItem('odx-ribbon', 'classic'); } catch { /* storage blocked */ } });
await ribbonPage.goto(`http://localhost:${server.address().port}/?ribbon=simple`);
await ribbonPage.waitForSelector('.odx-content p');
const simpleNow = await ribbonPage.evaluate(() => [document.querySelector('.odx-toolbar').classList.contains('odx-simple'), window.odx.ribbon]);
await ribbonPage.click('.odx-ribbon-mode');
const afterBtn = await ribbonPage.evaluate(() => [document.querySelector('.odx-toolbar').classList.contains('odx-simple'), window.odx.ribbon]);
const afterApi = await ribbonPage.evaluate(() => { window.odx.setRibbon('simple'); return [document.querySelector('.odx-toolbar').classList.contains('odx-simple'), window.odx.ribbon]; });
check(simpleNow.join() === 'true,simple' && afterBtn.join() === 'false,classic' && afterApi.join() === 'true,simple', `ribbon: 'simple' starts the one-line ribbon over the user's classic; the button and setRibbon() switch it (${simpleNow} | ${afterBtn} | ${afterApi})`);
await ribbonPage.close();

// Thin compound borders (thinborders-word.docx, made in Word, its print checked: boxes of triple, medium-gap, double and
// small-gap borders at 0.25 to 1pt, 14 of them under 2.25pt in all): each drawn by a picture of its lines, the box as tall as Word lays it out: its
// line, and above and below it the border's whole width (lines and gaps) and its 1pt space, not CSS double's 2.25pt.
await open('thinborders-word.docx');
const thin = await page.evaluate(() => {
  const ps = [...document.querySelectorAll('.odx-content p')];
  return ps.flatMap((p, i) => {
    const m = /^Box \d+ (\w+) (\d+)/.exec(p.textContent);
    if (!m || !ps[i + 1]) return [];
    const cs = getComputedStyle(p);
    return [{ val: m[1], sz: +m[2], extra: p.getBoundingClientRect().height - ps[i + 1].getBoundingClientRect().height, image: /svg/.test(cs.borderImageSource), widths: [cs.borderTopWidth, cs.borderLeftWidth].map(parseFloat) }];
  });
});
const wordWidth = { triple: (w) => 5 * w, double: (w) => 3 * w, thinThickMediumGap: (w) => 2 * w, thickThinMediumGap: (w) => 2 * w, thinThickThinMediumGap: (w) => 3 * w, thinThickSmallGap: (w) => w + 1.5, thickThinSmallGap: (w) => w + 1.5 };
const thinOnes = thin.filter((b) => wordWidth[b.val](b.sz / 8) < 2.25);
// (Wider ones the same since 0.8.3: CSS double drew a triple as two lines, a thick-thin pair's lines equal.)
const off = thin.map((b) => ({ ...b, want: 2 * ((wordWidth[b.val](b.sz / 8) + 1) * 4) / 3 })).filter((b) => !b.image || Math.abs(b.extra - b.want) > 0.1 || b.widths.some((w) => w !== Math.round(w)));
check(thinOnes.length === 14 && thin.length === 18 && !off.length,
  `compound borders drawn as Word prints them, at Word's width (${thinOnes.length} thin of ${thin.length}; off: ${JSON.stringify(off)})`);

// Table cells' compound borders (tableborders-word.docx, made in Word, its print checked: twelve tables, each with one
// kind of compound border all round and inside, 0.25pt to 3pt): drawn by pictures behind the cells, not CSS's collapsed
// double (2.25pt at least, two lines); each row as much taller than its text as Word's line between rows is wide, the
// text half the side line's width further in, as Word has it.
await open('tableborders-word.docx');
const cellBorders = await page.evaluate(() => [...document.querySelectorAll('.odx-content table')].slice(0, 11).flatMap((t) => [...t.rows].slice(0, 2).map((tr) => {
  const td = tr.cells[1], p = td.querySelector('p'), cs = getComputedStyle(td);
  return { pic: /svg/.test(cs.backgroundImage), css: cs.borderTopStyle + cs.borderLeftStyle, extra: td.getBoundingClientRect().height - p.getBoundingClientRect().height, inset: p.getBoundingClientRect().left - td.getBoundingClientRect().left };
})));
const tableWord = [1.25, 1, 2, 1.5, 3.25, 2.25, 9, 15, 4.5, 5.25, 9]; // (each kind's width, pt: COMPOUND in props.js)
const cellOff = cellBorders.filter((c, i) => { const w = (tableWord[i >> 1] * 4) / 3; return !c.pic || c.css !== 'nonenone' || Math.abs(c.extra - w) > 0.1 || Math.abs(c.inset - w / 2 - cellBorders[0].inset + (tableWord[0] * 2) / 3) > 0.1; });
check(cellBorders.length === 22 && !cellOff.length, `table cells' compound borders drawn as Word prints them, rows and text where Word's widths put them (off: ${JSON.stringify(cellOff)})`);

// Where printed text sits in its line, as Word's PDFs put it (lines.docx, Times New Roman, mklines.mjs): an exactly 40pt
// line's baseline four fifths of the way down it; a line of 10pt text with a 12pt run, raised 3pt or not, on the 12pt
// run's ascent (plus its raise), not the first run's. Each measured from the line's top as laid out (the print shift
// taken off), the 12pt run's baseline found from its text's box with a probe in its font (the raised run's own
// baseline: 3pt over its line's, whose top is 3pt higher too).
await open('lines.docx');
await page.emulateMedia({ media: 'print' });
await page.waitForTimeout(300);
const printedBase = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].slice(1, 4).map((p) => {
  // (The exact line's text; the 12pt run in the others.)
  const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  let t = w.nextNode();
  while (t && p.textContent.includes('Large') && t.data !== 'Large') t = w.nextNode();
  const cs = getComputedStyle(t.parentElement);
  const q = document.createElement('div');
  q.style.cssText = `position:absolute;left:-9999px;top:0;line-height:0;font:${cs.font};white-space:nowrap`;
  q.innerHTML = '<span>X<i style="display:inline-block;width:0;height:0"></i></span>';
  q.firstChild.style.lineHeight = cs.lineHeight;
  document.body.append(q);
  const above = q.querySelector('i').getBoundingClientRect().top - q.firstChild.getBoundingClientRect().top; // baseline below the text's box
  q.remove();
  const r = document.createRange();
  r.selectNodeContents(t);
  const top = p.getBoundingClientRect().top - (parseFloat(getComputedStyle(p).top) || 0);
  return +(r.getClientRects()[0].top + above - top).toFixed(2);
}));
await page.emulateMedia({ media: 'screen' });
// (Times New Roman's ascent 0.9336em, its print offset 0.34px: baseline.js.)
// (Exact lines: 0.1px over 0.8 of the line, baseline.js.)
const baseWant = [0.8 * 800 / 15 + 0.1, 0.9336 * 16 + 0.34, 0.9336 * 16 + 0.34];
check(printedBase.length === 3 && printedBase.every((b, i) => Math.abs(b - baseWant[i]) < 0.75), `printed baselines where Word's are: exact 40pt at 0.8 of the line, a larger run's ascent over the first run's (${JSON.stringify(printedBase)}, want ${JSON.stringify(baseWant.map((b) => +b.toFixed(2)))})`);
// A line of runs in two sizes as tall as Word makes it (mixlines.js): the highest top (ascent plus raise) and the lowest
// foot (descent less raise), each in its own font, and 1.15's extra from the taller font below them (Word's PDFs of the
// same lines, mixed-word: 14.932pt, 17.111pt, 18.438pt).
const mixedH = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].slice(5, 8).map((p) => [+p.getBoundingClientRect().height.toFixed(2), p.dataset.odxMix === '1']));
const mixedWant = [14.932, 17.111, 18.438].map((pt) => +(pt * 4 / 3).toFixed(2));
check(mixedH.length === 3 && mixedH.every(([h, on], i) => on && Math.abs(h - mixedWant[i]) < 0.15), `lines of mixed sizes, raised and lowered, as tall as Word's (${JSON.stringify(mixedH)}, want ${JSON.stringify(mixedWant)})`);
// The same in a header and a footer (lhfix.js storyLineFix): 10pt Calibri with a 12pt run raised 3pt; 10pt Times New
// Roman with a 12pt raised run between 12pt tabs, which don't size the line (hf-lines.docx, mkhflines.mjs; Word's PDFs:
// 17.111pt, 16.366pt).
await open('hf-lines.docx');
await page.waitForTimeout(500);
const hfH = await page.evaluate(() => ['.odx-header p', '.odx-footer p'].map((s) => { const p = document.querySelector(s); return p ? [+p.getBoundingClientRect().height.toFixed(2), p.dataset.odxMix === '1'] : null; }));
const hfWant = [17.111, 16.366].map((pt) => +(pt * 4 / 3).toFixed(2));
check(hfH.every((x, i) => x && x[1] && Math.abs(x[0] - hfWant[i]) < 0.15), `header and footer lines of mixed sizes as tall as Word's (${JSON.stringify(hfH)}, want ${JSON.stringify(hfWant)})`);
// A large document (over 1,500 blocks: those far off screen skip layout) still gets its line heights fixed (lhfix.js):
// its probes are laid out, as many of its paragraphs fixed as in the same lines in a document a quarter as long.
// (Large as it opens: once drawn and the browser idle, it is laid out in full, pages.js, and no longer marked so.)
const fixedShare = async (f) => { await open(f); const large = await page.evaluate(() => document.querySelector('.odx-large') != null); await page.waitForTimeout(1500); return page.evaluate((lg) => { const ps = [...document.querySelectorAll('.odx-content > p')].slice(0, 400); return [ps.filter((p) => p.dataset.odxLh).length, lg]; }, large); };
const [smallFixed] = await fixedShare('lines-small.docx');
const [largeFixed, isLarge] = await fixedShare('lines-large.docx');
check(isLarge && largeFixed >= smallFixed, `a large document's line heights fixed as a small one's (${largeFixed} of its first 400 paragraphs, ${smallFixed} in the small one)`);

// Lines of nothing but spaces are the paragraph mark's line in Word, whatever size the spaces are (mixlines.js blankOnly,
// pararules.js lineSize; mkspacing.mjs, Word: 15pt apart as the "X" lines between them, 21.75pt over an 18pt mark);
// a no-break space is a letter, sizing its line (22.5pt).
await open('spaces-lines.docx');
const blankH = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].map((p) => +(p.getBoundingClientRect().height * 0.75).toFixed(2)));
const xH = blankH[0];
const blankOk = blankH.length >= 13 && [2, 4, 6, 8].every((i) => Math.abs(blankH[i - 1] - xH) < 0.4) && Math.abs(blankH[9] - 22.35) < 0.5 && Math.abs(blankH[11] - 22.35) < 0.5;
check(blankOk, `lines of spaces as tall as their paragraph mark's line, as in Word (${JSON.stringify(blankH)})`);
// A cell's lone space, its paragraph of no size of its own, keeps its table style's size (mkspacing.mjs): as tall as the
// cell of text beside it.
await open('spaces-table-style.docx');
const cellH = await page.evaluate(() => [...document.querySelectorAll('.odx-content td p')].map((p) => +p.getBoundingClientRect().height.toFixed(2)));
check(cellH.length === 2 && Math.abs(cellH[0] - cellH[1]) < 0.3 && cellH[0] < 17, `a cell's lone space keeps its table style's size (${JSON.stringify(cellH)})`);
// An empty section-break paragraph that is all a columns section holds keeps its line, as in one column (mkspacing.mjs;
// Word: the text after it two lines under the line before it).
await open('columns-empty-section.docx');
const colGap = await page.evaluate(() => { const ps = [...document.querySelectorAll('.odx-content > p')]; const t = (i) => ps[i].getBoundingClientRect().top; return ps.length >= 4 ? [+(t(3) - t(1)).toFixed(2), +(t(1) - t(0)).toFixed(2)] : null; });
check(colGap && Math.abs(colGap[0] - 2 * colGap[1]) < 1, `a lone empty paragraph in a columns section keeps its line (${JSON.stringify(colGap)})`);
// An empty paragraph carrying a continuous section break takes no room (why-tests gen2.py; Word: the line under it
// 28pt below the line above, an 8pt space after that one and a 6pt space before this one both kept); one that is all
// of its section keeps its line (three in a row: the first none, the next two a line each).
const tops = () => page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].map((p) => { const r = p.getBoundingClientRect(); return [+r.top.toFixed(2), +r.height.toFixed(2), +parseFloat(getComputedStyle(p).marginTop).toFixed(2)]; }));
await open('section-break-empty.docx');
const sb = await tops();
check(sb.length === 5 && sb[2][1] === 0 && Math.abs(sb[3][0] - sb[1][0] - (sb[1][1] + 10.67 + 8)) < 0.5, `an empty continuous section break paragraph takes no room, the spaces around it kept (${JSON.stringify(sb)})`);
await open('section-break-only.docx');
const so = await tops();
check(so.length === 7 && so[2][1] === 0 && so[3][1] > 15 && so[4][1] > 15 && Math.abs(so[5][0] - so[1][0] - 3 * so[1][1]) < 0.5, `a section break paragraph that is all of its section keeps its line (${JSON.stringify(so)})`);
// A run of paragraphs each kept with the next that is longer than a page breaks where the page ends (why-tests gen2.py
// keepchain; Word: 10 lines, then the 55 kept paragraphs from the next page's top, 51 on it and 4 with the 6 after).
await open('keep-chain-long.docx');
await page.waitForTimeout(600);
const kc = await page.evaluate(() => [document.querySelectorAll('.odx-sheets > .odx-sheet').length, document.querySelectorAll('.odx-content > p').length]);
check(kc[0] === 3 && kc[1] === 71, `a keep-with-next run longer than a page takes the pages its lines need (${JSON.stringify(kc)}, want 3 pages)`);
// A list from the paragraph's style, the style setting a left indent itself: that indent, not the list level's
// (why-tests gen2.py styleind; Word: the label at the margin, the level's left 2.375in).
await open('list-style-indent.docx');
const si = await page.evaluate(() => { const ps = [...document.querySelectorAll('.odx-content > p')]; return [ps[1].getBoundingClientRect().left - ps[0].getBoundingClientRect().left, parseFloat(getComputedStyle(ps[1]).textIndent)]; });
check(Math.abs(si[0] - 24) < 0.5 && Math.abs(si[1] + 24) < 0.5, `a style's own indent over its list level's (${JSON.stringify(si)}, want 24px in, hanging 24px)`);
// A list label in a font of its own makes Word's line taller by what its top reaches above the text's (why-tests gen2.py
// lblfont, lblsym; Word, 12pt Calibri text: a Segoe UI bullet 2.03px a line, a Courier New "o" nothing, 14pt 0.31px, a
// Verdana number 0.85px).
await open('list-label-font.docx');
const lf = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].map((p) => +p.getBoundingClientRect().height.toFixed(2)));
const lfNear = (from, add) => [0, 1, 2].every((k) => Math.abs(lf[from + k] - lf[0] - add) < 0.12);
check(lf.length === 18 && lfNear(3, 2.03) && lfNear(6, 0) && lfNear(9, 0.31) && lfNear(12, 0.85) && lfNear(15, 0), `a label's own font makes its line as much taller as Word does (${JSON.stringify(lf)})`);
// Spaces with no size of their own before larger text in a header keep their style's size (gen2.py fix3; Word: 28 spaces
// of 11pt Calibri, 93px, before 20pt text, the whole on one line).
await open('header-leading-spaces.docx');
const hs = await page.evaluate(() => { const p = document.querySelector('.odx-hf.odx-header p'); const t = p && [...p.querySelectorAll('.odx-t')].find((s) => s.textContent.trim()); return t ? [+p.getBoundingClientRect().height.toFixed(1), +(t.getBoundingClientRect().left - p.getBoundingClientRect().left).toFixed(1)] : null; });
check(hs && hs[0] < 40 && Math.abs(hs[1] - 92.8) < 1.5, `a header's unsized spaces stay in their style's size (${JSON.stringify(hs)}, want one line, the text 92.8px in)`);
// An empty paragraph whose line reaches past the bottom margin only by its line spacing's extra stays on its page, as a
// line of text does (gen2.py fix3; Word: 46 lines, the empty 1.15-line paragraph under them, the next text at page 2's top).
const onSheets = (a, b) => page.evaluate(([i, j]) => { const sh = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((x) => x.getBoundingClientRect()); const ps = [...document.querySelectorAll('.odx-content > p')]; const of = (k) => { const r = ps[k].getBoundingClientRect(); const n = sh.findIndex((x) => r.top >= x.top - 1 && r.top < x.bottom); return [n + 1, n < 0 ? null : +(r.top - sh[n].top).toFixed(1)]; }; return [sh.length, of(i), of(j)]; }, [a, b]);
await open('empty-line-page-foot.docx');
await page.waitForTimeout(600);
const ef = await onSheets(46, 47);
check(ef[0] === 2 && ef[1][0] === 1 && ef[2][0] === 2 && Math.abs(ef[2][1] - 96) < 1.5, `an empty line at a page's foot stays there (${JSON.stringify(ef)}, want it on page 1, the next text at page 2's top)`);
// A footer is measured with its tabs at their stops, as it is drawn (gen2.py fix3; Word: three footer lines, the first
// text after a right tab that just fits, 46 body lines above them).
await open('footer-tab-fits.docx');
await page.waitForTimeout(600);
const ftab = await onSheets(45, 46);
check(ftab[0] === 2 && ftab[1][0] === 1 && ftab[2][0] === 2 && Math.abs(ftab[2][1] - 96) < 1.5, `a footer as tall as it is drawn, its tabs laid out (${JSON.stringify(ftab)}, want 46 lines on page 1)`);
// A paragraph kept with the next that may split sends only its last lines on with the block that moved (why-tests gen2.py
// keepsplit; Word: 30 lines, four 4-line paragraphs each kept with the next, the fourth split two lines and two, the
// four after it on page 2; the whole run had gone to page 2).
await open('keep-next-splits.docx');
await page.waitForTimeout(600);
const kn = await onSheets(30, 34);
const knCut = await page.evaluate(() => { const sh = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((x) => x.getBoundingClientRect()); const r = document.querySelectorAll('.odx-content > p')[33].getBoundingClientRect(); return r.top < sh[0].bottom && r.bottom > sh[1].top; });
check(kn[0] === 2 && kn[1][0] === 1 && kn[2][0] === 2 && knCut, `a kept paragraph splits to stay with the next (${JSON.stringify(kn)}, split ${knCut})`);

// A page's last line is fitted exactly (why-tests gen2.py fitedge; Word: 56 single-spaced 10pt Arial lines, the first
// with space after; with 81 twips of it the last line ends on the bottom margin and stays, with 82 it goes to page 2.
// And 46 lines of exactly 14pt: 80 twips stay, 81 go).
const sheetsOf = async (f) => { await open(f); await page.waitForTimeout(400); return page.evaluate(() => document.querySelectorAll('.odx-sheets > .odx-sheet').length); };
const fitPages = [await sheetsOf('page-foot-line-fits.docx'), await sheetsOf('page-foot-line-passes.docx'), await sheetsOf('page-foot-exact-fits.docx'), await sheetsOf('page-foot-exact-passes.docx')];
check(fitPages.join() === '1,2,1,2', `a page's last line stays when it ends on the margin and goes a twip past it (${fitPages}, want 1,2,1,2 pages)`);
// Spacing adds up down a page as Word's does (gen2.py sprnd; Word: 28 Calibri paragraphs at 1.08 lines with 6.65pt
// after, the last 761.23px under the first; cut to the browser's 64ths it came out half a pixel short).
await open('spacing-adds-up.docx');
const sa = await page.evaluate(() => { const ps = [...document.querySelectorAll('.odx-content > p')]; return [ps.length, +(ps[ps.length - 1].getBoundingClientRect().top - ps[0].getBoundingClientRect().top).toFixed(2)]; });
check(sa[0] === 28 && Math.abs(sa[1] - 761.23) < 0.3, `paragraph spacing adds up to Word's down a page (${JSON.stringify(sa)}, want 761.23px)`);
// A table with no alignment of its own takes its style's, and starts half its left line's width in (gen2.py tbljc,
// cellx; Word: the centred style's table's cells start their text 215.2 and 415.3px in; a table with 6pt lines, 107.23
// and 373.99px).
const cellStarts = () => page.evaluate(() => { const sh = document.querySelector('.odx-sheets > .odx-sheet').getBoundingClientRect(); return [...document.querySelectorAll('.odx-content table tr:first-child td')].map((td) => { const r = document.createRange(); r.selectNodeContents(document.createTreeWalker(td, NodeFilter.SHOW_TEXT).nextNode()); return +(r.getBoundingClientRect().left - sh.left).toFixed(2); }); });
await open('table-style-centred.docx');
const tsc = await cellStarts();
check(Math.abs(tsc[0] - 215.2) < 0.6 && Math.abs(tsc[1] - 415.3) < 0.6, `a table centred by its style (${JSON.stringify(tsc)}, want 215.2 and 415.3px)`);
await open('table-left-line-half.docx');
const tlh = await cellStarts();
check(Math.abs(tlh[0] - 107.23) < 0.6 && Math.abs(tlh[1] - 373.99) < 0.6, `a table starts half its left line's width in (${JSON.stringify(tlh)}, want 107.23 and 373.99px)`);
// A table style's header row repeats on the next page (gen2.py stylehdr; Word: two pages, the second starting with it).
await open('table-style-header-row.docx');
await page.waitForTimeout(400);
const tsh = await page.evaluate(() => [document.querySelectorAll('.odx-sheets > .odx-sheet').length, document.querySelectorAll('.odx-repeat tr').length]);
check(tsh[0] === 2 && tsh[1] === 1, `a table style's header row repeats (${JSON.stringify(tsh)}, want 2 pages and one repeated row)`);
// In an exact-height row the centred cell's last paragraph's space after counts in the centring (gen2.py exva; Word: a
// 20pt row, the centred one-line cell with 6pt after it 2.1px under the top-aligned cell's line).
await open('exact-row-centred-space-after.docx');
const erc = await page.evaluate(() => { const tr = document.querySelector('.odx-content table tr'); const tops = [...tr.cells].map((td) => { const r = document.createRange(); r.selectNodeContents(document.createTreeWalker(td, NodeFilter.SHOW_TEXT).nextNode()); return r.getBoundingClientRect().top; }); return [+tr.getBoundingClientRect().height.toFixed(2), +(tops[1] - tops[0]).toFixed(2)]; });
check(Math.abs(erc[0] - 26.67) < 0.4 && Math.abs(erc[1] - 2.1) < 0.3, `space after counts when an exact row's cell centres its text (${JSON.stringify(erc)}, want a 26.67px row, 2.1px down)`);
// A list number wider than its room pushes its text to the next tab stop (gen2.py lbltab; Word: "Appendix A." from
// 1in, a half-inch hanging indent, its text 85.3px after the number's start, at the stop at 2000 twips).
await open('list-wide-label-tab.docx');
const lwl = await page.evaluate(() => { const p = document.querySelector('.odx-content p.odx-li'); const r = document.createRange(); r.selectNodeContents(document.createTreeWalker(p, NodeFilter.SHOW_TEXT).nextNode()); return +(r.getBoundingClientRect().left - p.getBoundingClientRect().left - parseFloat(getComputedStyle(p).textIndent)).toFixed(2); });
check(Math.abs(lwl - 85.33) < 0.6, `text after a wide list number starts at the next tab stop (${lwl}px after the number's start, want 85.3)`);
// A restarted list restarts at its first paragraph only, and a paragraph holding only a page break is counted and shows
// no number (gen2.py numbrk; Word: A, B, then E to H after an override to 5; A, a page break, C and D on page 2).
const liLabels = () => page.evaluate(() => [...document.querySelectorAll('.odx-content p.odx-li')].map((p) => (p.dataset.label ?? '').replace('Item ', '').replace('.', '')).join(''));
await open('list-restart-first-only.docx');
const lrf = await liLabels();
await open('list-break-only-item.docx');
await page.waitForTimeout(300);
const lbo = [await liLabels(), await page.evaluate(() => document.querySelectorAll('.odx-sheets > .odx-sheet').length)];
check(lrf === 'ABEFGH' && lbo[0] === 'ACD' && lbo[1] === 2, `a list restarts once, and counts a page-break-only item without numbering it (${lrf}, ${JSON.stringify(lbo)}, want ABEFGH and ACD on 2 pages)`);
// A list number with no font of its own takes its paragraph mark's (gen2.py markfont; Word: the number in the mark's
// Courier New, the text in its run's Arial).
await open('list-number-mark-font.docx');
const lmf = await page.evaluate(() => { const p = document.querySelector('.odx-content p.odx-li'); return [getComputedStyle(p, '::before').fontFamily, getComputedStyle(p.querySelector('.odx-t')).fontFamily]; });
check(lmf[0].replace(/"/g, '').startsWith('Courier New') && lmf[1].replace(/"/g, '').startsWith('Arial'), `a list number in its paragraph mark's font (${JSON.stringify(lmf)})`);
// Tabs in a centred paragraph go to their stops, and a right stop past the paragraph's right indent, on the margin, is
// a stop (gen2.py tabjc, tabind; Word: left, centre and right text at 96-123, 395-421 and 685-720px; with a 1in right
// indent the text ends at the indent and its number sits at 687-720px on the next line).
const textSpans = (n) => page.evaluate((k) => { const sh = document.querySelector('.odx-sheets > .odx-sheet').getBoundingClientRect(); const out = []; const w = document.createTreeWalker(document.querySelectorAll('.odx-content > p')[k], NodeFilter.SHOW_TEXT); for (let t = w.nextNode(); t; t = w.nextNode()) { if (!t.data.trim()) continue; const r = document.createRange(); r.selectNodeContents(t); const b = r.getBoundingClientRect(); out.push([Math.round(b.left - sh.left), Math.round(b.right - sh.left), Math.round(b.top - sh.top)]); } return out; }, n);
await open('tabs-centred-paragraph.docx');
const tcp = await textSpans(1);
check(tcp.length === 3 && Math.abs(tcp[1][0] - 395) < 2 && Math.abs(tcp[1][1] - 421) < 2 && Math.abs(tcp[2][0] - 685) < 2 && Math.abs(tcp[2][1] - 720) < 2, `tabs in a centred paragraph go to their stops (${JSON.stringify(tcp)})`);
await open('tab-stop-past-indent.docx');
const tpi = await textSpans(1);
check(tpi.length === 3 && tpi[1][1] <= 624.5 && Math.abs(tpi[2][0] - 687) < 2 && Math.abs(tpi[2][1] - 720) < 2 && tpi[2][2] > tpi[1][2] + 10, `a right stop on the margin past the right indent (${JSON.stringify(tpi)}, want the number at 687-720px on the next line)`);

// A cell's text keeps clear of the lines at its sides: its side margin in, or half the line's width where that is more
// (why-tests gen2.py cellin; Word: no side margins and 3pt lines, the two cells' text starting 100.03 and 366.79px in,
// 2px inside each cell's edge).
await open('cell-margin-under-line.docx');
const cml = await cellStarts();
check(Math.abs(cml[0] - 100.03) < 0.4 && Math.abs(cml[1] - 366.79) < 0.4, `a cell's text half its side line's width in where its margin is less (${JSON.stringify(cml)}, want 100.03 and 366.79px)`);
// A line may end after a hyphen whatever follows, as in Word (gen2.py brkhy; Word: "ab-" stays at the first line's end
// before an opening quote, a full stop or a dollar sign, where the browser took it down with what followed).
await open('break-after-hyphen.docx');
const bah = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].map((p) => {
  const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  let at = 0;
  const tops = [];
  const i = p.textContent.indexOf('-');
  for (let t = w.nextNode(); t; t = w.nextNode()) {
    for (const k of [i, i + 1]) if (k >= at && k < at + t.data.length) { const r = document.createRange(); r.setStart(t, k - at); r.setEnd(t, k - at + 1); tops[k - i] = r.getBoundingClientRect().top; }
    at += t.data.length;
  }
  return tops[1] > tops[0] + 5;
}));
check(bah.length === 14 && bah.every(Boolean), `a line ends after a hyphen whatever follows it (${JSON.stringify(bah)})`);

// A page number whose field's opening mark is in a shaded run is drawn in that run's colour, on its shading, over what
// its saved result is set in (why-tests gen2.py fldfmt2, fldfmt4; Word: the mark's run red on grey, the saved result
// green: the number red on grey).
await open('footer-number-shaded-mark.docx');
const fns = await page.evaluate(() => { const t = document.querySelector('.odx-hf.odx-footer .odx-t'); const cs = t && getComputedStyle(t); return cs ? [t.textContent, cs.color, cs.backgroundColor] : null; });
check(fns && fns[0] === '1' && fns[1] === 'rgb(255, 0, 0)' && fns[2] === 'rgb(230, 230, 230)', `a page number in its shaded opening mark's colour and shading (${JSON.stringify(fns)})`);

// An underline is set where its font puts it, as Word draws it (why-tests gen2.py uline; Word: under 12pt Courier New
// the line's top 3.5px below the baseline, under Calibri 1.7px; the browser's own place is 1px for both).
await open('underline-from-font.docx');
const uff = await page.evaluate(() => { const ps = [...document.querySelectorAll('.odx-content > p')]; const u = ps[3].querySelector('u') || ps[3].querySelector('.odx-t'); const cs = getComputedStyle(u); return [ps.length, cs.textUnderlinePosition, cs.textDecorationThickness, cs.textDecorationLine]; });
check(uff[0] === 7 && uff[1] === 'from-font' && uff[2] === 'from-font' && /underline/.test(uff[3]), `an underline placed and sized from its font (${JSON.stringify(uff)})`);

// A table is set in by half the left line of its first row's first cell, the cell's own line over the table's (why-tests
// gen2.py cellx2; Word: the table's lines 0.5pt, the first row's first cell with a 3pt left line of its own: its text
// 105.15px in, 2px further than the table's own line would put it).
await open('table-first-cell-line.docx');
const tfc = await cellStarts();
check(Math.abs(tfc[0] - 105.15) < 0.4, `a table set in by half its first cell's own left line (${JSON.stringify(tfc)}, want 105.15px)`);

// Tab stops are measured from the section's own margin, not the last section's (why-tests gen2.py secttab; Word: a
// first section with 1in margins before a last one with half-inch ones: its centre stop's text at 367-401px, its right
// stop's ending at 672px).
await open('tabs-section-margin.docx');
const tsm = await textSpans(0);
check(tsm.length === 3 && Math.abs(tsm[1][0] - 367) < 2 && Math.abs(tsm[2][1] - 672) < 2, `tab stops from the section's own margin (${JSON.stringify(tsm)}, want the middle text from 367px, the right one's end at 672)`);

// A list number takes its level's letter-spacing, and the level's size where larger than its style's (gen2.py lblsp;
// Word: "Item 1." condensed by 1.1pt at the level's 14pt over unsized text in an 11pt document).
await open('list-number-spacing.docx');
const lns = await page.evaluate(() => { const cs = getComputedStyle(document.querySelector('.odx-content p.odx-li'), '::before'); return [parseFloat(cs.letterSpacing), parseFloat(cs.fontSize)]; });
check(Math.abs(lns[0] + 1.467) < 0.05 && Math.abs(lns[1] - 18.67) < 0.1, `a list number in its level's letter-spacing and larger size (${JSON.stringify(lns)}, want -1.47px and 18.67px)`);

// The text after a list number starts at the paragraph's own tab stop when one lies between the number's end and the
// hanging indent (gen2.py lbltab2; Word: "1." at 100 twips, the indent at 1540, a stop at 819: the text at 150.6px,
// not the indent's 198.8).
await open('list-own-stop-before-indent.docx');
const los = await textSpans(1);
check(los.length >= 1 && Math.abs(los[0][0] - 151) < 2, `text after a list number at the paragraph's own stop before its indent (${JSON.stringify(los[0])}, want 151px)`);

// No line break at punctuation inside unspaced text where Word has none (gen2.py brkpn; Word: a token of letters round
// "](", ".(", "?", an ellipsis and eleven more moves whole to the next line; the browser broke it at four of them).
await open('no-break-at-brackets.docx');
const nbb = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].map((p) => {
  const text = p.textContent;
  const at = [text.indexOf('abcd'), text.indexOf('wxyz')];
  const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  const tops = [];
  for (let t = w.nextNode(), from = 0; t; from += t.data.length, t = w.nextNode()) {
    at.forEach((k, n) => { if (k >= from && k < from + t.data.length) { const r = document.createRange(); r.setStart(t, k - from); r.setEnd(t, k - from + 1); tops[n] = r.getBoundingClientRect().top; } });
  }
  return Math.abs(tops[0] - tops[1]) < 2;
}));
check(nbb.length === 15 && nbb.every(Boolean), `no line break at punctuation inside unspaced text (${JSON.stringify(nbb)})`);

// A link the document came with, its runs in no Hyperlink style, is drawn as its runs: Word gives a link no colour or
// line of its own (gen2.py lnk; Word: black, no line under it).
await open('link-unstyled.docx');
const lus = await page.evaluate(() => { const a = document.querySelector('.odx-content a.odx-link'); const cs = a && getComputedStyle(a); return cs ? [cs.color, cs.textDecorationLine] : null; });
check(lus && lus[0] === 'rgb(0, 0, 0)' && lus[1] === 'none', `a link without a Hyperlink style drawn as its runs (${JSON.stringify(lus)})`);

// A text box in a footer, set 1in from the page's left edge and 9.5in down it, holding a page number: at that place
// whatever the page's margin (2in here), each page showing its own number (gen2.py hfbox; Word: its text from 105.6px
// on both pages, "1" and "2").
await open('footer-textbox-page-number.docx');
await page.waitForTimeout(300);
const ftb = await page.evaluate(() => [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((sh) => {
  const img = sh.querySelector('.odx-footer img');
  if (!img) return null;
  const s = sh.getBoundingClientRect(), r = img.getBoundingClientRect();
  let svg = '';
  try { svg = decodeURIComponent(img.getAttribute('src').split(',').slice(1).join(',')); } catch { /* not one of ours */ }
  const m = /data-odx-fld="PAGE"[^>]*>([^<]*)</.exec(svg);
  return [m && m[1], +(r.left - s.left).toFixed(1), +(r.top - s.top).toFixed(1)];
}));
check(ftb.length === 2 && ftb.every((f) => f && Math.abs(f[1] - 96) < 1 && Math.abs(f[2] - 912) < 1) && ftb[0][0] === '1' && ftb[1][0] === '2', `a footer's text box set from the page's edge, each page's number in it (${JSON.stringify(ftb)}, want 96px across, 912 down, "1" and "2")`);

// Justified text whose runs carry letter-spacing of their own is squeezed as any other (gen2.py justsp; Word's lines
// of this document, 71 of them, break where these do).
await open('justified-letter-spacing.docx');
await page.waitForTimeout(300);
const jls = await page.evaluate(() => [document.querySelectorAll('.odx-content [style*="word-spacing:-"], .odx-content [style*="word-spacing: -"]').length, document.querySelectorAll('.odx-content [style*="letter-spacing"]').length]);
check(jls[0] >= 5 && jls[1] >= 50, `justified lines of letter-spaced runs squeezed (${JSON.stringify(jls)}, want 5 or more squeezed pieces)`);

// An underlined tab has its line across its width, through the letters that reach below it, and the spaces that end a
// paragraph have none (why-tests gen2.py utab; Word: one unbroken line from "Alpha" across the tab to "Bravo"'s end;
// none under forty underlined spaces at a paragraph's end).
await open('underlined-tab.docx');
const utb = await page.evaluate(() => { const t = document.querySelector('.odx-content .odx-tab'); return [getComputedStyle(t).textDecorationLine, getComputedStyle(t, '::after').content !== 'none', getComputedStyle(t.closest('p')).textDecorationSkipInk]; });
await open('underlined-closing-spaces.docx');
const ucs = await page.evaluate(() => { const t = document.querySelector('.odx-content .odx-trail'); return t ? getComputedStyle(t).fontSize : null; });
check(/underline/.test(utb[0]) && utb[1] && utb[2] === 'none' && ucs === '0px', `an underlined tab's line across it, none under a paragraph's closing spaces (${JSON.stringify([utb, ucs])})`);

// An underline with a colour of its own is drawn, in that colour (gen2.py ucol; Word: a red line under black text).
await open('underline-own-colour.docx');
const uoc = await page.evaluate(() => { const e = document.querySelector('.odx-content [style*="text-decoration-color"]'); const cs = e && getComputedStyle(e); return cs ? [cs.textDecorationLine, cs.textDecorationColor, cs.color] : null; });
check(uoc && /underline/.test(uoc[0]) && uoc[1] === 'rgb(255, 0, 0)' && uoc[2] === 'rgb(0, 0, 0)', `an underline in its own colour (${JSON.stringify(uoc)})`);

// Justified paragraphs that hold a tab are squeezed after it, as Word squeezes them (why-tests gen2.py justtab; Word's
// lines of this document, "(a)" and a tab to the hanging indent before each paragraph's text, break where these do).
await open('justified-after-tab.docx');
await page.waitForTimeout(300);
const jat = await page.evaluate(() => { const ps = [...document.querySelectorAll('.odx-content > p')]; return [ps.length, ps.filter((p) => p.querySelector('.odx-tab') && p.querySelector('[style*="word-spacing:-"], [style*="word-spacing: -"]')).length]; });
check(jat[0] === 12 && jat[1] >= 8, `justified lines after a tab squeezed (${JSON.stringify(jat)}, want 8 or more of the 12 paragraphs with a squeezed line)`);

// A continuous section whose text starts a page (its first paragraph set to break before) has that page as its first:
// the first-page header there, the ordinary one after (gen2.py ctitle; Word: "FIRST" on page 2, the first section's
// three-line header on pages 1, 3 and 4).
await open('continuous-section-first-page.docx');
await page.waitForTimeout(300);
const csf = await page.evaluate(() => [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((sh) => (sh.querySelector('.odx-hf.odx-header')?.textContent || '').slice(0, 5)));
check(csf.length === 4 && csf[0] === 'HDRAH' && csf[1] === 'FIRST' && csf[2] === 'HDRAH' && csf[3] === 'HDRAH', `a continuous section's first-page header on the page its text starts (${JSON.stringify(csf)})`);

// A table row of exact height never breaks across pages, its text cut to the row or not: it moves whole (why-tests
// gen2.py exrow; Word: 0.5in rows of three lines, the fourth reaching the page's foot: three rows on the first page,
// three on the second, none over the bottom margin).
await open('exact-row-whole.docx');
await page.waitForTimeout(300);
const erw = await page.evaluate(() => { const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect()); return [...document.querySelectorAll('.odx-content table tr')].map((tr) => { const ps = [...tr.cells[0].querySelectorAll('p')]; const r = document.createRange(); r.selectNodeContents(ps[0]); const a = r.getBoundingClientRect(); r.selectNodeContents(ps[ps.length - 1]); const b = r.getBoundingClientRect(); const p = sheets.findIndex((s) => a.top >= s.top - 0.5 && a.top < s.bottom); return [p, sheets.findIndex((s) => b.bottom > s.top && b.bottom <= s.bottom + 0.5), Math.round(sheets[p].bottom - b.bottom)]; }); });
check(erw.length === 6 && erw.map((r) => r[0]).join('') === '000111' && erw.every((r) => r[0] === r[1] && r[2] >= 90), `rows of exact height moved whole to the next page (${JSON.stringify(erw)}, want three rows a page, each row's text on one page and clear of the bottom margin)`);

// A right-aligned list number ends at its place whatever its width, and its text starts at the paragraph's own tab
// stop before the indent (gen2.py lblright; Word: "i." to "xviii." ending at 180px, their text at 186.1px).
await open('list-number-right-aligned.docx');
const lra = [await textSpans(1), await textSpans(8), await textSpans(9)].map((t) => t[0][0]);
check(lra.every((x) => Math.abs(x - 186) < 1.5), `text after right-aligned list numbers, narrow and wide, at one place (${JSON.stringify(lra)}, want 186px)`);

// A list letter's width is taken as it is drawn, without kerning, when its text is sent to the next tab stop
// (gen2.py lblkern; Word: "I." before the stop at 12px has its text at 108px, "P.", 12.9px wide, at the next, 120px;
// kerned, the "P." measured 11.6px and kept its text against it).
await open('list-letter-next-stop.docx');
const lns2 = [await textSpans(9), await textSpans(16)].map((t) => t[0][0]);
check(Math.abs(lns2[0] - 108) < 1.5 && Math.abs(lns2[1] - 120) < 1.5, `text after a list letter at the first stop past its unkerned width (${JSON.stringify(lns2)}, want 108 and 120px)`);

// A list number under a highlighted paragraph mark is highlighted, with its room up to the text (why-tests gen2.py
// lblhl; Word: yellow from the number's place at 144px to the text at 192, under a mark highlighted yellow).
await open('list-number-mark-highlight.docx');
const lmh = await page.evaluate(() => { const p = document.querySelector('.odx-content p.odx-li'); const cs = getComputedStyle(p, '::before'); return [cs.backgroundColor, parseFloat(cs.paddingTop) > 4 && parseFloat(cs.marginTop) < -4, Math.round(parseFloat(cs.minWidth))]; });
check(lmh[0] === 'rgb(255, 255, 0)' && lmh[1] && lmh[2] === 48, `a list number highlighted with its paragraph mark (${JSON.stringify(lmh)}, want yellow over 48px of room)`);

// A tab stop written after one further along is not a stop: Word takes them in rising order (gen2.py tabdup; Word:
// stops written 880 then 900 twips take the tab to 154.7px, written 900 then 880 to 156).
await open('tab-stops-out-of-order.docx');
const tso = [await textSpans(1), await textSpans(2)].map((t) => t[1][0]);
check(Math.abs(tso[0] - 155) < 1.01 && Math.abs(tso[1] - 156) < 0.6 && tso[1] > tso[0], `a tab stop written after one further along ignored (${JSON.stringify(tso)}, want 155 and 156px)`);

// A line that crosses the foot of an exact-height row is cut there in print, not drawn whole (gen2.py exclip; Word:
// under a 20pt row holding two 12pt lines, the second line's ink ends 26.7px below the row's top, 4.5px of it showing).
await open('exact-row-cut-line.docx');
await page.emulateMedia({ media: 'print' });
await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
await page.waitForTimeout(300);
const ecl = await page.evaluate(() => {
  const foot = document.querySelector('.odx-content:not(.odx-pl-host) tr[data-exact] > td').getBoundingClientRect().bottom;
  return [...document.querySelectorAll('.odx-printlines tr[data-exact] p')].map((p) => {
    const b = p.getBoundingClientRect();
    const m = /^inset\(\S+ \S+ ([-\d.]+)px/.exec(p.style.clipPath || '');
    return [+(b.top - foot).toFixed(1), +(b.bottom - foot).toFixed(1), m ? +(b.bottom - Number(m[1]) - foot).toFixed(1) : null];
  });
});
await page.evaluate(() => window.dispatchEvent(new Event('afterprint')));
await page.emulateMedia({ media: 'screen' });
check(ecl.length === 2 && ecl[0][1] <= 0 && ecl[0][2] === null && ecl[1][0] < -3 && ecl[1][1] > 3 && ecl[1][2] !== null && Math.abs(ecl[1][2]) <= 0.5,
  `a line crossing an exact row's foot cut at the foot in print (${JSON.stringify(ecl)}, want two lines, the second from above the foot to below it and cut at 0)`);

// Garamond's single line is its own 1.125em, and a paragraph's floor is no taller than its font's line (why-tests
// gen2.py garaline and floor; Word: 12pt Garamond paragraphs 18px apart, 12pt Courier New ones 18.12px, empty or not;
// at the 1.15em floor both stood 18.39px).
await open('garamond-single-lines.docx');
const gsl = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].slice(0, 20).map((p) => +p.getBoundingClientRect().height.toFixed(2)));
check(gsl.length === 20 && gsl.every((h) => Math.abs(h - 18) < 0.06), `single-spaced 12pt Garamond paragraphs one line tall, 1.125em (${JSON.stringify(gsl.slice(0, 4))}, want 18px)`);
await open('single-lines-under-floor.docx');
const sluf = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].map((p) => +p.getBoundingClientRect().height.toFixed(2)));
check(sluf.length === 42 && sluf.slice(0, 20).every((h) => Math.abs(h - 18.125) < 0.07) && sluf.slice(21, 41).every((h) => Math.abs(h - 18.125) < 0.07),
  `single-spaced 12pt Courier New paragraphs, with text and empty, as tall as the font's line (${JSON.stringify([sluf[0], sluf[25]])}, want 18.12px, not the 18.4px floor)`);

// A run in another font makes a Calibri line taller only past Calibri's own Windows ascent and descent (why-tests
// gen2.py mixdesc; Word: 11pt Calibri lines at 1.08 holding a Garamond run are Calibri's own 19.32px; measured from
// hhea's shallower descent they stood 19.7px).
await open('calibri-line-garamond-run.docx');
await page.waitForTimeout(300);
const clg = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].slice(0, 20).map((p) => +p.getBoundingClientRect().height.toFixed(2)));
check(clg.length === 20 && clg.every((h) => Math.abs(h - 19.32) < 0.08), `Calibri lines holding a Garamond run no taller than Calibri's own (${JSON.stringify(clg.slice(0, 4))}, want 19.32px)`);

// A row at a page's top starts under its own top line only: where the line over it is the upper row's bottom line, its
// text starts at the page's top, the row moved whole or split (why-tests gen2.py rowcont; Word: under 3pt bottom-only
// lines the second page's first line is where a plain first line is; the room for the line had it 4.3px low).
const rowAtTop = async (file) => {
  await open(file);
  await page.waitForTimeout(400);
  return page.evaluate(() => {
    const sh = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
    // (Each line of text, not its paragraph's box: one carried over the page break has its box, and the spacer that
    // carries it, start on the page before.)
    const r = document.createRange();
    const first = [...document.querySelectorAll('.odx-content td p')].flatMap((p) => { r.selectNodeContents(p); return [...r.getClientRects()].filter((q) => q.width > 1 && q.height > 5).map((q) => q.top); }).filter((t) => sh[1] && t >= sh[1].top + 50).sort((x, y) => x - y)[0];
    return [sh.length, first == null ? null : +(first - sh[1].top).toFixed(1)];
  });
};
const ram = await rowAtTop('row-moved-under-upper-line.docx');
const ras = await rowAtTop('row-split-under-upper-line.docx');
check(ram[0] === 2 && Math.abs(ram[1] - 96.3) < 1 && ras[0] === 2 && Math.abs(ras[1] - 96.3) < 1,
  `a row at a page's top with no top line of its own starts at the page's top (${JSON.stringify([ram, ras])}, want 96px on page 2, moved whole and split)`);

// ... but beside a merged cell whose lower part sets a top line of its own, the row keeps that line's room (gen2.py
// rowmerge; Word: 3pt lines, the second page's first line 4px lower than with no such line).
const rbm = await rowAtTop('row-beside-merged-top-line.docx');
check(Math.abs(rbm[1] - 100.3) < 1, `a row at a page's top keeps room for a merged cell's own top line (${JSON.stringify(rbm)}, want 100px on page 2)`);

// A page of exact lines is fitted by their set height, not the font's near-equal line (why-tests gen2.py foottie; Word: 55
// lines of 11.5pt in 10pt Times New Roman, the page at 12960 twips keeps its last line, at 12961 it goes on).
const onSheetOf = async (f, n) => { await open(f); await page.waitForTimeout(400); return page.evaluate((k) => { const q = document.createRange(); q.selectNodeContents(document.querySelectorAll('.odx-content > p')[k]); const t = q.getClientRects()[0].top; return [...document.querySelectorAll('.odx-sheets > .odx-sheet')].findIndex((s) => { const r = s.getBoundingClientRect(); return t >= r.top && t < r.bottom; }); }, n); };
const ftie = [await onSheetOf('page-foot-tie-fits.docx', 110), await onSheetOf('page-foot-tie-passes.docx', 110)];
check(ftie[0] === 1 && ftie[1] === 2, `exact lines fitted to the page by the twip (${JSON.stringify(ftie)}, want the last line on sheets 1 and 2)`);

// ... and lines of multiple spacing by the font's exact units and the exact multiple (gen2.py foottie2; Word: 12pt Times New
// Roman at 259/240, the page's 42nd line stays with 473 twips of space before the second paragraph and goes on with 474).
const ftm = [await onSheetOf('page-foot-multiple-fits.docx', 41), await onSheetOf('page-foot-multiple-passes.docx', 41)];
check(ftm[0] === 0 && ftm[1] === 1, `multiple-spaced lines fitted to the page by the twip (${JSON.stringify(ftm)}, want the 42nd line on sheets 0 and 1)`);

// ... and bulleted lines with the bullet's exact room (gen2.py foottie3; Word: 12pt Times New Roman under Symbol bullets, the
// page's 44th line stays with 59 twips of space before the second paragraph and goes on with 60).
const footBul = [await onSheetOf('page-foot-bullets-fits.docx', 43), await onSheetOf('page-foot-bullets-passes.docx', 43)];
check(footBul[0] === 0 && footBul[1] === 1, `bulleted lines fitted to the page by the twip (${JSON.stringify(footBul)}, want the 44th line on sheets 0 and 1)`);

// ... and empty paragraphs at their exact line (gen2.py footempty; Word: 12pt Calibri with 16 empty paragraphs among the lines,
// the page's 44th paragraph stays with 69 twips of space before the second and goes on with 70).
const footEmpty = [await onSheetOf('page-foot-empties-fits.docx', 43), await onSheetOf('page-foot-empties-passes.docx', 43)];
check(footEmpty[0] === 0 && footEmpty[1] === 1, `empty paragraphs fitted to the page by the twip (${JSON.stringify(footEmpty)}, want the 44th paragraph on sheets 0 and 1)`);

// Tabs in a row each go on to the next default stop, from an indent that isn't a whole 64th of a px too (why-tests
// gen-lb.py lbtabs; Word: 56 documents, indents of 0 to 437 twips, every piece of text on its stop).
await open('tabs-in-a-row-from-an-indent.docx');
const tabRow = await page.evaluate(() => [...document.querySelectorAll('.odx-content p')].filter((p) => p.querySelector('.odx-tab')).map((p) => {
  const edge = p.getBoundingClientRect().left - parseFloat(getComputedStyle(p).marginLeft);
  const ends = [...p.querySelectorAll('.odx-tab')].map((t) => t.getBoundingClientRect().right - edge);
  return ends.every((e, i) => Math.abs(e / 48 - Math.round(e / 48)) < 0.02 && (!i || e > ends[i - 1] + 1));
}));
check(tabRow.length === 18 && tabRow.every(Boolean), `tabs in a row each reach their own default stop (${tabRow.filter(Boolean).length} of ${tabRow.length} paragraphs, want 18)`);

// A justified paragraph that opens with spaces: they take room on its first line, so Word squeezes no word on there that
// doesn't fit with them (why-tests gen-lb.py lbleadsp; Word: 54 documents, spaces at the start, after a line break and
// after a tab; this one 11pt Calibri, two spaces, its paragraphs' lines as Word's print).
await open('justified-opening-spaces.docx');
const leadLines = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].filter((p) => p.textContent.trim()).map((p) => {
  const r = document.createRange(); r.selectNodeContents(p);
  return new Set([...r.getClientRects()].filter((x) => x.width > 0).map((x) => Math.round(x.bottom))).size;
}));
check(JSON.stringify(leadLines) === '[3,4,4,5,4,4,4,4,4,4,4,4]', `justified paragraphs opening with spaces break as Word's (${JSON.stringify(leadLines)})`);

// A header paragraph holding only floating shapes is a full line, however many it holds (why-tests gen2.py hfshapes; Word:
// 17.9px in 11pt Calibri with one text box, two or three; with two it stood at the 1.15em floor, 16.86px).
await open('header-two-floating-shapes.docx');
const hfs = await page.evaluate(() => +document.querySelector('.odx-hf.odx-header p').getBoundingClientRect().height.toFixed(2));
check(Math.abs(hfs - 17.9) < 0.1, `a header line of two floating shapes a full line (${hfs}, want 17.9px)`);

// A Symbol bullet's taller first line holds when its paragraph opens with a tracked deletion, hidden in this view (why-tests
// gen2.py lbldel; Word: the line as tall as one opening with plain text; the room was put on the hidden deletion).
await open('bullet-opens-with-deletion.docx');
await page.waitForTimeout(300);
const bod = await page.evaluate(() => { const p = document.querySelectorAll('.odx-content > p')[1]; return [+p.getBoundingClientRect().height.toFixed(2), parseFloat(p.style.getPropertyValue('--odx-lbl-own')) || 0]; });
check(bod[1] > 0.3 && Math.abs(bod[0] - 17.9 - bod[1]) < 0.12, `a bullet's room kept over a paragraph opening with a hidden deletion (${JSON.stringify(bod)}, want the 17.9px line plus the room)`);

// A Symbol bullet over runs in fonts and sizes of their own makes the line its top over their deepest foot (why-tests
// gen2.py lblmix; Word: 11pt Calibri and 10pt Arial in a 12pt Times New Roman style, 18.63px; it stood 17.84px).
await open('bullet-over-mixed-runs.docx');
await page.waitForTimeout(300);
const bmr = await page.evaluate(() => +document.querySelectorAll('.odx-content > p')[1].getBoundingClientRect().height.toFixed(2));
check(Math.abs(bmr - 18.65) < 0.1, `a bullet's line over runs of mixed fonts (${bmr}, want 18.65px)`);

// A number whose level sets a larger size but no font makes its line taller by what its top reaches above the text's,
// in the font it's drawn in (why-tests gen-vs.py lblsize; Word: a 12pt number over 10pt Arial, a 17.83px line; it was
// the text's 15.33px).
await open('list-number-larger.docx');
await page.waitForTimeout(300);
const nlg = await page.evaluate(() => +document.querySelectorAll('.odx-content > p')[1].getBoundingClientRect().height.toFixed(2));
check(Math.abs(nlg - 17.83) < 0.1, `a number larger than its text (${nlg}, want 17.83px)`);

// The empty line after a paragraph's closing line break is the paragraph mark's size (why-tests gen-vs.py trailbr; Word:
// 14pt Arial text and break over an unsized mark in a 12pt style, 21.47 plus 18.40px; it was two 14pt lines).
await open('closing-break-mark-line.docx');
await page.waitForTimeout(300);
const cbm = await page.evaluate(() => +document.querySelectorAll('.odx-content > p')[1].getBoundingClientRect().height.toFixed(2));
check(Math.abs(cbm - 39.87) < 0.15, `the line after a closing break in the mark's size (${cbm}, want 39.87px)`);

// A number adds no room where text early in its paragraph reaches as high (why-tests gen-m108.py lblrun; Word: a 12pt
// number in the style's Times New Roman over a 10pt Arial lead-in and then the style's own 12pt text, a 36.7px item of
// two lines; with room over the lead-in it was 37.25px).
await open('list-number-early-text.docx');
await page.waitForTimeout(300);
const net = await page.evaluate(() => +document.querySelectorAll('.odx-content > p')[1].getBoundingClientRect().height.toFixed(2));
check(Math.abs(net - 36.81) < 0.1, `a number under text as tall early on its line (${net}, want 36.81px)`);

// A closing break's own size adds nothing, and the empty line after it is in the mark's font where the text's runs set
// none (why-tests gen-m108.py endbr; Word: 10pt Times New Roman, a 12pt break over a 12pt mark, 33.3px; 10pt Segoe UI
// from the style over an 11pt Calibri mark, 36px; they were 36.81 and 37.25px).
const cbs = [];
for (const f of ['closing-break-own-size.docx', 'closing-break-mark-font.docx']) {
  await open(f);
  await page.waitForTimeout(300);
  cbs.push(await page.evaluate(() => +document.querySelectorAll('.odx-content > p')[1].getBoundingClientRect().height.toFixed(2)));
}
check(Math.abs(cbs[0] - 33.73) < 0.1 && Math.abs(cbs[1] - 35.64) < 0.1, `a closing break's own size and its mark's font (${JSON.stringify(cbs)}, want [33.73,35.64])`);

// A floating table placed down from its text and moved to the next page keeps that drop there, and the text after it
// goes with it (why-tests gen2.py fltmove; Word: 182 twips down, the table 12.13px under the page's top, the paragraph
// after it at the page's top beside it; the table had lost the drop and the paragraph stayed past the old page's foot).
await open('float-moved-keeps-drop.docx');
await page.waitForTimeout(400);
const fmd = await page.evaluate(() => {
  const sh = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
  const t = document.querySelector('.odx-content .tableWrapper').getBoundingClientRect();
  const r = document.createRange(); r.selectNodeContents(document.querySelector('.odx-content .tableWrapper').nextElementSibling);
  const a = r.getClientRects()[0];
  return [sh.length, +(t.top - sh[1].top).toFixed(1), +(a.top - sh[1].top).toFixed(1)];
});
check(fmd[0] === 2 && Math.abs(fmd[1] - 108.13) < 0.6 && Math.abs(fmd[2] - 96) < 1.5, `a floating table moved on keeps its drop, the text after it with it (${JSON.stringify(fmd)}, want the table at 108.13px and the text at 96px on page 2)`);

// A table set down on its page (vertAnchor page) with text before it in the file: past its place, it and the text after
// it go on to the next page, the text before staying; at its place, the text before it on that page goes under it
// (mksetdown.mjs; Word's prints: row 1 at 612.7px on page 2 and the paragraph after it at 111.3px there; row 1 at
// 114.7px on page 2, the caption before it in the file at 472.7px under it).
const sdAt = async (f, words) => { await open(f); await page.waitForTimeout(400); return page.evaluate((ws) => {
  const sh = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
  return ws.map((w) => {
    const walk = document.createTreeWalker(document.querySelector('.odx-content'), NodeFilter.SHOW_TEXT);
    for (let n; (n = walk.nextNode());) {
      if (!n.data.startsWith(w)) continue;
      const r = document.createRange(); r.selectNodeContents(n); const a = r.getClientRects()[0];
      const s = sh.findIndex((x) => a.top >= x.top && a.top < x.bottom);
      return [s + 1, +(a.top - sh[s].top).toFixed(1)];
    }
    return null;
  });
}, words); };
const sdB = await sdAt('float-setdown-before.docx', ['Row 1 label', 'Text after the table.']);
const sdC = await sdAt('float-setdown-caption.docx', ['Row 1 label', 'Table 2 A caption']);
check(sdB[0]?.[0] === 2 && sdB[1]?.[0] === 2 && sdB[1][1] < 110 && sdC[0]?.[0] === 2 && sdC[0][1] < 115 && sdC[1]?.[0] === 2 && sdC[1][1] > 400,
  `a table set down on its page: past its place on to the next page, the text before it under it (${JSON.stringify([sdB, sdC])})`);

// A floating table placed up from its text (tblpY -59) doesn't break where it starts: past the page's foot it goes on to
// the next page whole, 59tw above the top margin; one longer than a page breaks there, and the empty paragraphs after it
// go beside its last rows (mkfloatup.mjs; Word's prints: row 1 on page 2, 3.9px above where a row at the top margin
// would be; the 70-row table's text after it on page 3 at the third line, beside its rows).
const fuAt = async (f, words) => { await open(f); await page.waitForTimeout(400); return page.evaluate((ws) => {
  const sh = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
  const t = document.querySelector('.odx-content table').getBoundingClientRect();
  const at = (y) => sh.findIndex((x) => y >= x.top && y < x.bottom);
  return [[at(t.top + 1) + 1, +(t.top - sh[at(t.top + 1)].top).toFixed(1)], ...ws.map((w) => {
    const walk = document.createTreeWalker(document.querySelector('.odx-content'), NodeFilter.SHOW_TEXT);
    for (let n; (n = walk.nextNode());) {
      if (n.data !== w) continue;
      const r = document.createRange(); r.selectNodeContents(n); const a = r.getClientRects()[0];
      const s = at(a.top);
      return [s + 1, +(a.top - sh[s].top).toFixed(1)];
    }
    return null;
  })];
}, words); };
const fuM = await fuAt('float-up-moves.docx', ['Row 30']);
const fuL = await fuAt('float-up-long-beside.docx', ['Row 70', 'After the table.']);
check(fuM[0][0] === 2 && Math.abs(fuM[0][1] - 92.07) < 0.6 && fuM[1]?.[0] === 2 && fuL[0][0] === 2 && fuL[1]?.[0] === 3 && fuL[2]?.[0] === 3 && fuL[2][1] < fuL[1][1],
  `a floating table placed up from its text moves on whole, drawn up, and the empty paragraphs after a long one go beside it (${JSON.stringify([fuM, fuL])}, want the table on page 2 at 92.07px, the text after the long one on page 3 above its last row)`);

// A floating table goes where its tblpXSpec puts it, against the margin or the page (why-tests gen2.py fltx3; Word: a 4in
// table centred in a 6.5in column starts 216px across; at the page's right its 0.5pt line ends at the page's edge).
const fltBox = async (f) => { await open(f); await page.waitForTimeout(300); return page.evaluate(() => { const s = document.querySelector('.odx-sheets > .odx-sheet').getBoundingClientRect(), t = document.querySelector('.odx-content table').getBoundingClientRect(); return [+(t.left - s.left).toFixed(2), +(t.right - s.left).toFixed(2)]; }); };
const fltC = await fltBox('float-table-centred.docx'), fltR = await fltBox('float-table-page-right.docx');
check(Math.abs(fltC[0] - 216) < 0.6 && Math.abs(fltR[1] - 815.67) < 0.6, `floating tables centred and at the page's right (${JSON.stringify([fltC, fltR])}, want 216 across and 815.67 at the right)`);

// A page a fraction of a pixel wider than whole pixels prints in a box of whole pixels, its text laid out at its own
// width: wider than the printed page, the browser cut each page a pixel late (a sample's 47 pages of 12242 by 15842
// twips, each page's text a pixel higher than the last).
await open('page-fraction-wide.docx');
await page.emulateMedia({ media: 'print' });
await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
await page.waitForTimeout(300);
const pfw = await page.evaluate(() => [document.querySelector('.odx-page').getBoundingClientRect().width, +document.querySelector('.odx-page > .odx-content').getBoundingClientRect().width.toFixed(2)]);
await page.evaluate(() => window.dispatchEvent(new Event('afterprint')));
await page.emulateMedia({ media: 'screen' });
check(pfw[0] === 816 && Math.abs(pfw[1] - 816.13) < 0.02, `a page a fraction wider prints in whole pixels, its text at its own width (${JSON.stringify(pfw)}, want 816 and 816.13)`);

// Text with an outer shadow (w14:shadow) has a taller line in Word: on the run, from its paragraph style, on an empty
// paragraph of that style (texteffects.js shadowRoom; mkspacing.mjs, Word's PDF: 3.6px for its preset at 15pt, 5.33px
// for an 8pt blur at 10pt).
await open('shadow-lines.docx');
const shH = await page.evaluate(() => [...document.querySelectorAll('.odx-content > p')].map((p) => +p.getBoundingClientRect().height.toFixed(2)));
const shOk = shH.length >= 11 && [2, 4, 6].every((i) => Math.abs(shH[i] - shH[0] - 3.6) < 0.4) && Math.abs(shH[9] - shH[8] - 5.33) < 0.5 && shH.length >= 13 && Math.abs(shH[11] - shH[0]) < 0.3;
check(shOk, `shadowed lines as much taller as Word makes them (${JSON.stringify(shH)})`);

// A long document's contents entries, laid out on screen (content-visibility, editor.css .odx-large): each entry's page
// number on its own line at the right stop, and the number hanging in its indent and the page number past the end
// indent drawn, not cut off at the paragraph's box until the caret came to it (a 122-page document's contents, both
// numbers missing on screen in Edge and Firefox).
const longTocEntries = async (pg) => pg.evaluate(() => [...document.querySelectorAll('.odx-content > p')].filter((p) => p.querySelector('.odx-tab')).map((p) => {
  const t = [...p.querySelectorAll('span.odx-t')];
  const a = t[0].getBoundingClientRect(), z = t[t.length - 1].getBoundingClientRect();
  return { line: Math.abs(a.top - z.top) < 2, clip: parseFloat(getComputedStyle(p).overflowClipMargin) || 0, large: !!p.closest('.odx-large') };
}));
await open('contents-long-document.docx');
await page.waitForTimeout(800);
const tocE = await longTocEntries(page);
check(tocE.length === 9 && tocE.every((e) => e.line && (!e.large || e.clip >= 96)), `a long document's contents entries: page numbers on their lines, nothing cut off at the box (${JSON.stringify(tocE.map((e) => [e.line, e.large, e.clip]))})`);
// And, once drawn and the browser idle, laid out in full as any document (pages.js): paragraphs split across pages on
// screen as in print (a 122-page document's had gone to the next page whole on screen, 125 pages).
const tocFull = await page.waitForFunction(() => !document.querySelector('.odx-large'), null, { timeout: 15000 }).then(() => true, () => false);
check(tocFull, 'a long document is laid out in full once it is drawn');

// A fit-to-contents table whose first column's merged cells hold text turned on its side keeps that column as given:
// laid out at its narrowest, the sideways text set the column 490px wide and the text beside it 134px.
await open('table-sideways-column.docx');
await page.waitForTimeout(500);
const sideCols = await page.evaluate(() => [...document.querySelector('.odx-content table > colgroup').children].map((c) => Math.round(parseFloat(c.style.width))));
check(JSON.stringify(sideCols) === '[79,545]', `a sideways-text column keeps its width (${JSON.stringify(sideCols)}, want [79,545])`);
// Its rows as tall as their other cells make them, the sideways text wrapping to that (rowfit.js): its length had been
// taken as what the merged cell needs, and it asked for the height it already had, a row 351px in one print and 666px in
// another.
const sideRows = await page.evaluate(() => [...document.querySelector('.odx-content table').rows].map((tr) => {
  const c = [...tr.cells].find((td) => /^horizontal/.test(getComputedStyle(td).writingMode));
  const cs = getComputedStyle(c);
  const ps = [...c.children];
  const need = ps[ps.length - 1].getBoundingClientRect().bottom - ps[0].getBoundingClientRect().top + parseFloat(cs.paddingTop) + parseFloat(cs.paddingBottom);
  return Math.round(tr.getBoundingClientRect().height - need);
}));
check(sideRows.every((d) => Math.abs(d) <= 1), `sideways merged cells don't stretch their rows (each row's height less its text cell's: ${JSON.stringify(sideRows)})`);

// Tables at a page's foot, as Word's prints of these generated documents break them (one-column tables of Times New
// Roman 12pt rows under lines of text). For each text: the sheet its first line is on, its top on the sheet, how often
// it is drawn (header rows repeat as copies) and whether its last line ends above the bottom margin.
const rowsAt = (texts) => page.evaluate((ts) => {
  const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
  const out = {};
  const w = document.createTreeWalker(document.querySelector('.odx-canvas') || document.body, NodeFilter.SHOW_TEXT);
  const r = document.createRange();
  for (let n = w.nextNode(); n; n = w.nextNode()) {
    const t = ts.find((x) => (x.endsWith('*') ? n.data.startsWith(x.slice(0, -1)) : n.data === x));
    if (!t) continue;
    r.selectNodeContents(n);
    const rs = [...r.getClientRects()].filter((q) => q.height > 0);
    if (!rs.length) continue;
    const k = sheets.findIndex((s) => rs[0].top >= s.top - 0.5 && rs[0].top < s.bottom);
    const last = rs[rs.length - 1];
    const kl = sheets.findIndex((s) => last.top >= s.top - 0.5 && last.top < s.bottom);
    const o = out[t] || (out[t] = { sheet: k, top: Math.round(rs[0].top - sheets[k].top), n: 0, inside: true });
    o.n++;
    if (kl < 0 || last.bottom > sheets[kl].bottom - 96 + 1) o.inside = false; // (1in margins)
  }
  return out;
}, texts);
// A keep-with-next heading row at the foot stays there when the next row can split, two of its lines with it (Word);
// with room for one line only (widow control) both go on; a row whose second paragraph alone keeps with the next
// doesn't keep (why-tests\tblbrk kn-f43, kn-f44, knml-f43).
await open('table-keep-row-split.docx');
let tb = await rowsAt(['Heading row', 'Row one*', 'Row two']);
check(tb['Heading row']?.sheet === 0 && tb['Row one*']?.sheet === 0 && tb['Row two']?.sheet === 1, `a keep-with-next row stays at the page's foot with the next row split under it (${JSON.stringify(tb)})`);
await open('table-keep-row-widow.docx');
tb = await rowsAt(['Heading row', 'Row one*']);
check(tb['Heading row']?.sheet === 1 && tb['Row one*']?.sheet === 1, `a keep-with-next row goes on with a next row that can leave only one line (${JSON.stringify(tb)})`);
await open('table-keep-row-second-paragraph.docx');
tb = await rowsAt(['Heading row line two', 'Row one*']);
check(tb['Heading row line two']?.sheet === 0 && tb['Row one*']?.sheet === 1, `a row whose second paragraph keeps with the next doesn't keep (${JSON.stringify(tb)})`);
// A row of 60 lines that can't split, under 20 lines of text: on to page 2's top and broken across pages 2 and 3
// (cstall-f20); it had stayed on page 1 and run past the bottom margin.
await open('table-tall-cantsplit-row.docx');
tb = await rowsAt(['Row zero', 'Row one*', 'Row two']);
check(tb['Row zero']?.sheet === 0 && tb['Row one*']?.sheet === 1 && tb['Row one*'].top < 110 && tb['Row two']?.sheet === 2, `a row taller than a page that can't split goes to the next page's top and breaks there (${JSON.stringify(tb)})`);
// Thirty rows, every one a header row, under 30 lines of text: the table starts page 2 (hdrall-f30-n30). In compatibility
// mode 14, three header rows with room for two: they break there and none repeats (hdr314-f44-n5).
await open('table-all-header-rows.docx');
tb = await rowsAt(['Row 0', 'Row 29']);
check(tb['Row 0']?.sheet === 1 && tb['Row 0'].n === 1 && tb['Row 29']?.sheet === 1, `a table of header rows that doesn't fit moves to the next page (${JSON.stringify(tb)})`);
await open('table-header-rows-mode14.docx');
tb = await rowsAt(['Row 0', 'Row 1', 'Row 2', 'Row 4']);
check(tb['Row 1']?.sheet === 0 && tb['Row 0'].n === 1 && tb['Row 2']?.sheet === 1 && tb['Row 4']?.sheet === 1, `mode-14 header rows at the foot break there, none repeating (${JSON.stringify(tb)})`);
// A table with a tracked-deleted row (shown without markup, as Word prints it): it breaks at the foot like any other, its
// rows inside the margins (delrow-f44: rows 0 and 2 on page 1, 3 to 9 on page 2).
await open('table-deleted-row.docx');
tb = await rowsAt(['Row 0', 'Row 2', 'Row 3', 'Row 9']);
check(tb['Row 2']?.sheet === 0 && tb['Row 3']?.sheet === 1 && tb['Row 9']?.sheet === 1 && Object.values(tb).every((x) => x.inside), `a table with a deleted row breaks at the page's foot (${JSON.stringify(tb)})`);
// In mode 14 a cell's three-line paragraph with room for two at the foot leaves two there, widow control or not (Word's
// print of ..\cellwidow-word two-def-m14-f42: "Leftaa" and "Leftbb" on page 1, "Leftcc" on page 2; mode 15 moves it).
// (The sheet each word of the cell's paragraph "Leftaa Leftbb Leftcc", one to a line, is on.)
const cellLines = () => page.evaluate(() => {
  const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
  const w = document.createTreeWalker(document.querySelector('.odx-content'), NodeFilter.SHOW_TEXT);
  const r = document.createRange();
  const out = [];
  for (let n = w.nextNode(); n; n = w.nextNode()) {
    for (const word of ['Leftaa', 'Leftbb', 'Leftcc']) {
      const i = n.data.indexOf(word);
      if (i < 0) continue;
      r.setStart(n, i);
      r.setEnd(n, i + word.length);
      const q = r.getBoundingClientRect();
      out.push(sheets.findIndex((s) => q.top >= s.top - 0.5 && q.top < s.bottom));
    }
  }
  return out;
});
await open('table-cell-widow-mode14.docx');
let cl = await cellLines();
check(JSON.stringify(cl) === '[0,0,1]', `mode 14: no widow control inside a cell at the page's foot (${JSON.stringify(cl)})`);
// The same row as a header row (w:tblHeader, under another): it doesn't split in mode 14, all of it on page 2 (Word's
// print of ..\cellwidow-word hdr2-m14-f42; sample 2133's table of header rows).
await open('table-header-row-whole-mode14.docx');
tb = await rowsAt(['Row zero']);
cl = await cellLines();
check(tb['Row zero']?.sheet === 0 && JSON.stringify(cl) === '[1,1,1]', `mode 14: a header row at the page's foot moves whole (${JSON.stringify([tb, cl])})`);
// In mode 14 a repeated header row keeps its own 2.25pt bottom line: on page 2 the next row's text is as far under the
// header's as on page 1 (..\hdrline-word h18-style-whole-m14: Word 33.4px both; it had been 2px less on page 2).
await open('table-header-line-mode14.docx');
{
  const d = await page.evaluate(() => {
    const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => s.getBoundingClientRect());
    const at = {};
    const w = document.createTreeWalker(document.querySelector('.odx-canvas') || document.body, NodeFilter.SHOW_TEXT);
    const r = document.createRange();
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      if (!/^(Head one|Row \d+)$/.test(n.data)) continue;
      r.selectNodeContents(n);
      const q = r.getClientRects()[0];
      if (!q) continue;
      const k = sheets.findIndex((s) => q.top >= s.top - 0.5 && q.top < s.bottom);
      (at[k] ||= []).push([n.data, q.top]);
    }
    const gap = (k) => { const l = (at[k] || []).sort((a, b) => a[1] - b[1]); const h = l.find((x) => x[0] === 'Head one'); const row = l.find((x) => x[0] !== 'Head one'); return h && row ? +(row[1] - h[1]).toFixed(2) : null; };
    return [gap(0), gap(1)];
  });
  check(d[0] != null && d[1] != null && Math.abs(d[1] - d[0]) < 0.75, `mode 14: a repeated header row keeps its wide bottom line (${JSON.stringify(d)})`);
}
// A paragraph style that switches its bulleted base style's list off (w:numId 0) starts its indents from 0: its own
// kept, its base's right indent too (style-numid0.docx, mkstylenum0.mjs; Word's print of it, 38 lines as Word's).
await open('style-numid0.docx');
{
  const m = await page.evaluate(() => Object.fromEntries([...document.querySelectorAll('.odx-content p')].map((p) => {
    const s = getComputedStyle(p);
    return [p.textContent.split(' ')[0], [+parseFloat(s.marginLeft).toFixed(1), +parseFloat(s.marginRight).toFixed(1), +parseFloat(s.textIndent).toFixed(1)]];
  })));
  const want = { dLx: [0, 0, 0], ddLx: [0, 0, 0], dPx: [0, 0, 0], dL540x: [36, 0, 0], dLrx: [0, 24.3, 0], dLfx: [0, 0, 13.3], dQx: [0, 96, 0], dQhx: [0, 96, -20], Px: [24, 0, 0] };
  const bad = Object.entries(want).filter(([k, v]) => !m[k] || m[k].some((x, i) => Math.abs(x - v[i]) > 0.2));
  check(!bad.length, `a style's own numId 0 drops the indents it inherits (${JSON.stringify(bad.map(([k]) => [k, m[k]]))})`);
}

// A picture wrapped top and bottom, placed down from its paragraph past the page's foot: Word keeps it on that page,
// drawn up to end at the page's edge (mode 12) or its bottom margin (mode 15), the text after it above it, a page's
// width wide across the margins (coverpic-word m12-a1, m15-b2; Word's prints: the picture at 387-792pt with the eight
// lines after it above, one page; at 390-720pt, "After 1" above it, "After 2" at page 2's top).
const pastFoot = (n) => page.evaluate((k) => {
  const sh = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((x) => x.getBoundingClientRect());
  const on = (r) => { const s = sh.findIndex((x) => r.top >= x.top - 1 && r.top < x.bottom); return [s + 1, s < 0 ? null : +(r.top - sh[s].top).toFixed(1)]; };
  const img = document.querySelector('.odx-content img.odx-wrap-block').getBoundingClientRect();
  const ps = [...document.querySelectorAll('.odx-content > p')];
  return { sheets: sh.length, img: [...on(img), +(img.bottom - img.top).toFixed(1), +img.width.toFixed(1)], after: [on(ps[k].getBoundingClientRect()), on(ps[k + 1].getBoundingClientRect())] };
}, n);
await open('picture-past-foot-mode12.docx');
await page.waitForTimeout(800);
const pf12 = await pastFoot(10);
check(pf12.sheets === 1 && pf12.img[0] === 1 && Math.abs(pf12.img[1] - 516) < 1.5 && Math.abs(pf12.img[3] - 816) < 1 && pf12.after.every((a) => a[0] === 1 && a[1] < 516), `a picture past the page's edge drawn up to it, the text after it above it (${JSON.stringify(pf12)}, want one page, the picture 516px down, 816px wide)`);
await open('picture-past-foot-mode15.docx');
await page.waitForTimeout(800);
const pf15 = await pastFoot(22);
check(pf15.sheets === 2 && pf15.img[0] === 1 && Math.abs(pf15.img[1] - 520) < 1.5 && pf15.after[0][0] === 1 && pf15.after[1][0] === 2 && Math.abs(pf15.after[1][1] - 96) < 1.5, `a picture past the bottom margin (mode 15) drawn up to it, the page's text ending at its top (${JSON.stringify(pf15)}, want it 520px down on page 1, "After 2" at page 2's top)`);
// Spaces alone on a centred paragraph's last line, after a break, keep that line as tall as a left-aligned one's
// (coverpic-word sp-c, sp-l; Word: both 24pt from the text's line to the spaces'; centred, the line had come out a sliver,
// sample 1387's text after its centred heading 15px high).
await open('centred-break-spaces.docx');
const cbsp = await page.evaluate(() => { const ps = [...document.querySelectorAll('.odx-content > p')]; return [ps[1], ps[3]].map((p) => +p.getBoundingClientRect().height.toFixed(1)); });
check(Math.abs(cbsp[0] - cbsp[1]) < 0.5 && cbsp[0] > 40, `a centred line of spaces after a break as tall as a left-aligned one (${JSON.stringify(cbsp)})`);

// Gaps between body paragraphs (test/fixtures/mkgaps.mjs; Word's prints of the same documents): how far below the
// first plain line each plain line after the feature starts, in Word's PDF (baselines, all 12pt Times New Roman).
const plainTops = () => page.evaluate(() => {
  const at = {};
  const w = document.createTreeWalker(document.querySelector('.odx-content'), NodeFilter.SHOW_TEXT);
  const r = document.createRange();
  for (let n = w.nextNode(); n; n = w.nextNode()) {
    const m = /^Plain line (\d+)/.exec(n.data);
    if (!m || at[m[1]] != null) continue;
    r.setStart(n, 0); r.setEnd(n, 1);
    at[m[1]] = r.getBoundingClientRect().top;
  }
  return Object.fromEntries(Object.entries(at).map(([k, v]) => [k, +(v - at[1]).toFixed(2)]));
});
for (const [f, want, what] of [
  ['gaps-theme-bidi.docx', { 4: 110.4 }, 'text in the theme\'s majorBidi font is Times New Roman'],
  ['gaps-deleted-neighbour.docx', { 2: 74.24 }, 'a paragraph deleted whole is no neighbour: the heading above keeps its space after'],
  ['gaps-bullet-missing-font.docx', { 2: 116 }, 'a 14pt bullet in a missing font makes its line Calibri\'s'],
  ['gaps-closing-tabs.docx', { 2: 128.8 }, 'tabs alone on a paragraph\'s last line make a line'],
  ['gaps-contextual-before.docx', { 2: 118.4 }, 'contextual spacing above leaves the space before less the space after'],
  ['gaps-border-groups.docx', { 2: 98.56, 3: 222.28 }, 'bordered paragraphs group where their boxes line up'],
  ['gaps-trailing-space-exact.docx', { 2: 90.08 }, 'a centred line\'s closing space adds no height on an exact line'],
]) {
  await open(f);
  const got = await plainTops();
  check(Object.entries(want).every(([k, v]) => got[k] != null && Math.abs(got[k] - v) < 1), `${what} (${JSON.stringify(got)}, Word ${JSON.stringify(want)})`);
}
// (A picture wrapped top and bottom 0.7in below a one-line paragraph: the next paragraph's first line above it, its second
// under it. Under it a line comes 0.9px lower than Word's, the browser setting lines past a float's shape on its pixels.)
{
  await open('gaps-band-below-lines.docx');
  const got = await plainTops();
  check(Math.abs(got[2] - 55.2) < 1 && Math.abs(got[3] - 146.08) < 1.5, `a top-and-bottom picture below a paragraph's lines leaves the lines above it there (${JSON.stringify(got)}, Word {"2":55.2,"3":146.08})`);
}
{
  await open('gaps-theme-bidi.docx');
  const ff = await page.evaluate(() => { const w = document.createTreeWalker(document.querySelector('.odx-content'), NodeFilter.SHOW_TEXT); for (let n = w.nextNode(); n; n = w.nextNode()) if (n.data.startsWith('Themed line')) return getComputedStyle(n.parentElement).fontFamily; return ''; });
  check(/^"?Times New Roman/.test(ff), `a majorBidi run is drawn in Times New Roman (${ff})`);
}

// The same contents in Firefox: its first answer for where a tab sits, read just after the tab before it on the line was
// sized, was where it had stood before; the leader came out 6px long and every page number went to a second line.
// (Where Playwright's Firefox is installed: CI installs only Chromium.)
{
  let ff = null;
  try { ff = await firefox.launch(); } catch { console.log('skip contents entries in Firefox (not installed)'); }
  if (ff) {
    const fp = await ff.newPage({ viewport: { width: 1300, height: 950 } });
    await fp.goto(`http://localhost:${server.address().port}/`);
    await fp.waitForSelector('.odx-content p');
    await fp.setInputFiles('input[type=file]', 'test/fixtures/contents-long-document.docx');
    await fp.waitForFunction((n) => window.odx?.fileName === n, 'contents-long-document.docx');
    await fp.waitForTimeout(2000);
    const ffE = await longTocEntries(fp);
    check(ffE.length === 9 && ffE.every((e) => e.line && (!e.large || e.clip >= 96)), `contents entries in Firefox: page numbers on their lines (${JSON.stringify(ffE.map((e) => e.line))})`);
    await ff.close();
  }
}

// Row and line heights (mkrowparts.mjs; Word's prints of the same documents).
{
  // A paragraph style of no base and no size takes its table style's size (14pt); Normal's own 12pt, a style's own
  // 10pt and one based on Normal stay.
  await open('table-style-size.docx');
  const sz = await page.evaluate(() => Object.fromEntries(['Normal0', 'Free0', 'FreeSz0', 'OnNormal0'].map((k) => {
    const p = [...document.querySelectorAll('.odx-content td > p')].find((q) => q.textContent.startsWith(k + ' '));
    const t = p && p.querySelector('span.odx-t');
    return [k, t ? +parseFloat(getComputedStyle(t).fontSize).toFixed(2) : null];
  })));
  check(sz.Normal0 === 16 && sz.Free0 === 18.67 && sz.FreeSz0 === 13.33 && sz.OnNormal0 === 16, `a style of no base and no size takes its table style's size in a cell (${JSON.stringify(sz)})`);
  // Where lines of text are, from their sheet's top: [sheet, top] of each line of the paragraph starting with `t`.
  const linesOf = (t) => page.evaluate((s) => {
    const sheets = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((q) => q.getBoundingClientRect());
    const p = [...document.querySelectorAll('.odx-content p')].find((q) => q.textContent.replace(/\s+/g, ' ').trim().startsWith(s));
    if (!p) return [];
    const tops = [];
    for (const n of p.querySelectorAll('span.odx-t')) {
      const r = document.createRange();
      r.selectNodeContents(n);
      for (const q of r.getClientRects()) if (q.height > 0 && !tops.some((y) => Math.abs(y - q.bottom) < 2)) tops.push(q.bottom);
    }
    return tops.sort((a, b) => a - b).map((y) => { const k = sheets.findIndex((q) => y > q.top && y <= q.bottom); return [k, +(y - sheets[k].top).toFixed(2)]; });
  }, t);
  // A row's empty last paragraph carried to page 2 alone: the row's part there is its whole minimum height (485tw), the
  // next row's text where Word has it, 15.4px lower than under that paragraph alone. (Its line's foot 145.67px down page
  // 2 on screen; printed, its baseline 143.5px, Word's 143.39; 130.3 under the empty paragraph alone.)
  await open('table-row-empty-part.docx');
  const nx = (await linesOf('NextRow marker'))[0];
  check(nx && nx[0] === 1 && Math.abs(nx[1] - 145.67) < 1, `a row's part holding only its empty last paragraph is its minimum height (${JSON.stringify(nx)})`);
  // Bullets typed by hand in Symbol before Calibri text in a Times New Roman paragraph: Calibri's lines, 17.92px apart.
  await open('symbol-bullet-lines.docx');
  const sl = (await page.evaluate(() => {
    const p = [...document.querySelectorAll('.odx-content p')].find((q) => q.textContent.includes('alpha bravo'));
    const r = document.createRange();
    r.selectNodeContents(p);
    const ys = [];
    for (const q of r.getClientRects()) if (q.height > 10 && !ys.some((y) => Math.abs(y - q.bottom) < 2)) ys.push(q.bottom);
    return ys.sort((a, b) => a - b);
  }));
  const pitch = sl.slice(1).map((y, i) => +(y - sl[i]).toFixed(2));
  check(pitch.length >= 3 && pitch.slice(1).every((d) => Math.abs(d - 17.92) < 0.25), `a hand-typed Symbol bullet's lines are the text font's (${JSON.stringify(pitch)})`);
  // Spaces and a closing tab under a larger mark: one line, the mark's (Times New Roman 12pt, 18.4px).
  await open('blank-tab-line.docx');
  const bh = await page.evaluate(() => {
    const p = [...document.querySelectorAll('.odx-content td > p')].find((q) => !q.textContent.trim());
    return p ? +p.getBoundingClientRect().height.toFixed(2) : null;
  });
  check(bh != null && Math.abs(bh - 18.4) < 0.3, `a line of spaces and a tab is the paragraph mark's one line (${bh})`);
}

check(!errors.length, `no page errors ${errors.join(' | ')}`);
await browser.close();
server.close();
console.log(failed ? `features: ${failed} failed` : 'features: all passed');
process.exit(failed ? 1 : 0);
