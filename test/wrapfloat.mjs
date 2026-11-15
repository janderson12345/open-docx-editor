// Browser check of floating shapes wrapped square as Word wraps them (test/fixtures/mkwrapfloat.mjs, checked against
// Word's PDF of it): text on both sides of one in the middle of the column (each line beside it split, whole words
// before it, none in its room), still so after typing; text on the left only of one wrapped "left"; one aligned to the
// column's centre set there; and one near the page's foot reaching past its edge kept on its page, drawn up to the edge.
import { chromium, firefox, webkit } from 'playwright';
import { serveDemo, launchOptions } from './browser.mjs';

const { server } = serveDemo();
// (BROWSER=firefox or webkit: in those engines, Playwright's own builds.)
const engine = { firefox, webkit }[process.env.BROWSER];
const browser = engine ? await engine.launch() : await chromium.launch(launchOptions());
const page = await browser.newPage({ viewport: { width: 1300, height: 950 } });
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
let failed = 0;
const check = (ok, what) => { console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}`); if (!ok) failed++; };

await page.goto(`http://localhost:${server.address().port}/`);
await page.waitForSelector('.odx-content p');
await page.setInputFiles('input[type=file]', 'test/fixtures/wrap-float.docx');
await page.waitForFunction(() => window.odx?.fileName === 'wrap-float.docx');
await page.waitForTimeout(2500);

// The lines of the paragraph holding picture `n` (and the next paragraph): words left of it, right of it, and in its
// room, for the lines beside it.
const beside = (n) => page.evaluate((k) => {
  const img = [...document.querySelectorAll('.odx-content img:not(.ProseMirror-separator)')][k];
  const r = img.getBoundingClientRect();
  const cs = getComputedStyle(img);
  const room = cs.getPropertyValue('--odx-both').trim().split(/\s+/).map(parseFloat);
  const [, dr, , dl] = room.length === 4 ? room : [0, parseFloat(cs.marginRight) || 0, 0, parseFloat(cs.marginLeft) || 0];
  const zone = { left: r.left - (room.length === 4 ? dl : 12), right: r.right + (room.length === 4 ? dr : 12), top: r.top, bottom: r.bottom };
  const p = img.closest('p');
  const out = { left: 0, right: 0, inside: [], lines: 0, float: cs.float, centre: (r.left + r.right) / 2 };
  const range = document.createRange();
  for (const el of [p, p.nextElementSibling]) {
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      for (const m of t.data.matchAll(/\S+/g)) {
        range.setStart(t, m.index); range.setEnd(t, m.index + m[0].length);
        const w = range.getBoundingClientRect();
        if (w.bottom <= zone.top || w.top >= zone.bottom) continue;
        if (w.right <= zone.left + 0.5) out.left++;
        else if (w.left >= zone.right - 0.5) out.right++;
        else out.inside.push(m[0]);
      }
    }
  }
  return out;
}, n);

const both = await beside(0);
check(both.left > 4 && both.right > 4 && !both.inside.length, `text on both sides of a shape in the middle of the column (${both.left} words left, ${both.right} right, in its room: ${both.inside.join(' ') || 'none'})`);
const left = await beside(1);
check(left.float === 'right' && left.left > 4 && left.right === 0 && !left.inside.length, `wrapped "left": text on its left only (${left.left} left, ${left.right} right, ${left.float} float)`);
const centre = await beside(2);
const colCentre = await page.evaluate(() => { const p = document.querySelector('.odx-content > p'); const r = p.getBoundingClientRect(); return (r.left + r.right) / 2; });
check(Math.abs(centre.centre - colCentre) < 1.5 && centre.left > 2 && centre.right > 2 && !centre.inside.length, `aligned to the column's centre, set there, text both sides (${(centre.centre - colCentre).toFixed(1)}px off)`);

// Typing before the shape in its paragraph: the lines settle again round it.
await page.click('.odx-content p:nth-of-type(2)', { position: { x: 5, y: 5 } });
await page.keyboard.press('Home');
await page.keyboard.type('Extra words typed here push the text along. ');
await page.waitForTimeout(1500);
const again = await beside(0);
check(!again.inside.length && again.left > 4 && again.right > 4, `after typing, still none in its room (${again.inside.join(' ') || 'none'})`);
await page.keyboard.press('Control+z');
await page.waitForTimeout(800);

// The foot: the paragraph on its page, its float drawn up to end at the page's edge.
const foot = await page.evaluate(() => {
  const img = [...document.querySelectorAll('.odx-content img:not(.ProseMirror-separator)')][3];
  const r = img.getBoundingClientRect();
  const para = img.closest('p').getBoundingClientRect();
  // (The page the paragraph is on.)
  const sheet = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((e) => e.getBoundingClientRect()).find((q) => para.top >= q.top && para.top < q.bottom) || { bottom: NaN };
  return { bottom: r.bottom, edge: sheet.bottom, paraTop: para.top, translate: img.style.translate };
});
check(foot.paraTop < foot.edge && Math.abs(foot.bottom - foot.edge) < 1.5 && /-/.test(foot.translate), `near the foot, kept on its page and drawn up to its edge (${(foot.bottom - foot.edge).toFixed(1)}px past; ${foot.translate})`);

// Rectangles in a footer or header the body's text wraps round, Word-made (pagefloat-word pf-footer, footwrap-word
// fw-right and fw-header): the body's lines beside each on the first page go on the sides Word's print has them, none
// in its room. The header's, 250pt below its paragraph, leaves the header's own text and the body's start where they
// are without it.
const hfBeside = async (file, sel) => {
  await page.setInputFiles('input[type=file]', `test/fixtures/${file}`);
  await page.waitForFunction((f) => window.odx?.fileName === f, file);
  await page.waitForTimeout(2500);
  return page.evaluate((s) => {
    const img = document.querySelector(s);
    const r = img.getBoundingClientRect();
    const out = { left: 0, right: 0, inside: [], hfText: null, bodyTop: document.querySelector('.odx-content > p').getBoundingClientRect().top - document.querySelector('.odx-sheet').getBoundingClientRect().top };
    const hp = img.closest('.odx-hf').querySelector('p');
    const hr = document.createRange();
    hr.selectNodeContents(hp);
    out.hfText = hr.getBoundingClientRect().left - img.closest('.odx-hf').getBoundingClientRect().left;
    const range = document.createRange();
    const walker = document.createTreeWalker(document.querySelector('.odx-content'), NodeFilter.SHOW_TEXT);
    for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      for (const m of t.data.matchAll(/\S+/g)) {
        range.setStart(t, m.index); range.setEnd(t, m.index + m[0].length);
        const w = range.getBoundingClientRect();
        if (w.bottom <= r.top || w.top >= r.bottom) continue;
        if (w.right <= r.left - 11.5) out.left++;
        else if (w.left >= r.right + 11.5) out.right++;
        else out.inside.push(m[0]);
      }
    }
    return out;
  }, sel);
};
const hfBoth = await hfBeside('footer-wrap.docx', '.odx-footer img');
check(hfBoth.left > 8 && hfBoth.right > 4 && !hfBoth.inside.length, `a footer's square-wrapped shape: the body's text both sides of it (${hfBoth.left} words left, ${hfBoth.right} right, in its room: ${hfBoth.inside.join(' ') || 'none'})`);
const onRight = await hfBeside('footer-wrap-right.docx', '.odx-footer img');
check(!onRight.left && onRight.right > 4 && !onRight.inside.length, `one wrapped "right": the body's text on its right only (${onRight.left} left, ${onRight.right} right, in its room: ${onRight.inside.join(' ') || 'none'})`);
const head = await hfBeside('header-wrap.docx', '.odx-header img');
check(head.left > 8 && !head.right && !head.inside.length && head.hfText < 1 && Math.abs(head.bodyTop - 96) < 2, `a header's tight-wrapped shape below its paragraph: the body's text on its left, the header's text and the body's start as without it (${head.left} left, ${head.right} right; header text at ${head.hfText.toFixed(1)}px, body at ${head.bodyTop.toFixed(1)}px)`);
check(!errors.length, `no page errors (${errors.join(' | ')})`);
await browser.close();
server.close();
console.log(failed ? `wrapfloat: ${failed} failed` : 'wrapfloat: all passed');
process.exit(failed ? 1 : 0);
