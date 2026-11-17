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
  const [, dr, , dl] = room.length >= 4 ? room : [0, parseFloat(cs.marginRight) || 0, 0, parseFloat(cs.marginLeft) || 0];
  const zone = { left: r.left - (room.length >= 4 ? dl : 12), right: r.right + (room.length >= 4 ? dr : 12), top: r.top, bottom: r.bottom };
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
// A float placed from the page's top, anchored in a paragraph that would end page 1 (pagefloat-word pf-early, made in
// Word): lifted there it wraps the text above until its paragraph moves on, so it goes on with it, as Word lays it out:
// the paragraph and its picture start page 2, page 1 keeping no picture.
await page.setInputFiles('input[type=file]', 'test/fixtures/float-page-moved.docx');
await page.waitForFunction(() => window.odx?.fileName === 'float-page-moved.docx');
await page.waitForTimeout(3000);
const moved = await page.evaluate(() => {
  const sheets = [...document.querySelectorAll('.odx-sheet')].map((s) => s.getBoundingClientRect());
  const sheetOf = (y) => sheets.findIndex((r) => y >= r.top && y < r.bottom);
  const lift = document.querySelector('.odx-lift img');
  const anchor = [...document.querySelectorAll('.odx-content > p')].find((p) => p.querySelector('.odx-lifted'));
  const rg = document.createRange();
  rg.selectNodeContents(anchor);
  const first = [...rg.getClientRects()].find((r) => r.height > 0 && r.width > 1);
  return { lift: lift ? sheetOf(lift.getBoundingClientRect().top + 2) : null, anchor: first ? sheetOf(first.top) : null };
});
check(moved.lift === 1 && moved.anchor === 1, `a page-placed float whose paragraph it pushes on goes on with it, as Word (picture on sheet ${moved.lift}, its paragraph on sheet ${moved.anchor})`);

// Justified lines beside floats, Word-made (wrap-word2 wj-both2, wl-upj, footwrap-word fw-justify), checked against
// Word's print: each side of a line split beside a both-sides shape justified on its own (its words before the shape
// ending at its room, those after it at the margin); the lines above a float drawn up to the page's edge and wrapped on
// its left only kept to its left, justified to its room; and printed, the lines beside a footer's shape so too.
const load = async (file) => {
  await page.setInputFiles('input[type=file]', `test/fixtures/${file}`);
  await page.waitForFunction((f) => window.odx?.fileName === f, file);
  await page.waitForTimeout(3500);
};
// The lines of the paragraphs under `root` (or the print layer's copies) beside the room [left, right] x [top, bottom]:
// each with its words' left and right edges, the paragraph's right edge and whether it's the paragraph's last line.
const linesBeside = (sel, room) => page.evaluate(([s, z]) => {
  const out = [];
  const range = document.createRange();
  for (const p of document.querySelectorAll(s)) {
    const pr = p.getBoundingClientRect();
    const words = [];
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      for (const m of t.data.matchAll(/\S+/g)) {
        range.setStart(t, m.index); range.setEnd(t, m.index + m[0].length);
        const w = range.getBoundingClientRect();
        if (w.width) words.push({ l: w.left, r: w.right, top: w.top, bottom: w.bottom });
      }
    }
    const lines = [];
    for (const w of words) {
      const ln = lines.find((x) => Math.abs(x.top - w.top) < 3);
      if (ln) ln.words.push(w); else lines.push({ top: w.top, bottom: w.bottom, words: [w] });
    }
    lines.forEach((ln, i) => {
      if (ln.bottom <= z.top || ln.top >= z.bottom) return;
      out.push({ left: ln.words.filter((w) => w.r <= z.left + 0.5).map((w) => w.r), right: ln.words.filter((w) => w.l >= z.right - 0.5).map((w) => w.r), inside: ln.words.filter((w) => w.r > z.left + 0.5 && w.l < z.right - 0.5).length, end: pr.right, last: i === lines.length - 1 });
    });
  }
  return out;
}, [sel, room]);

await load('wrap-justified.docx');
const rooms = await page.evaluate(() => [...document.querySelectorAll('.odx-content img[style*="--odx-both"]')].map((img) => {
  const r = img.getBoundingClientRect();
  const [, dr, , dl] = getComputedStyle(img).getPropertyValue('--odx-both').trim().split(/\s+/).map(parseFloat);
  return { left: r.left - dl, right: r.right + dr, top: r.top, bottom: r.bottom };
}));
let split = 0, off = 0, inRoom = 0;
for (const z of rooms) {
  for (const ln of await linesBeside('.odx-content > p', z)) {
    inRoom += ln.inside;
    if (ln.last || ln.left.length < 2 || ln.right.length < 2) continue;
    split++;
    // (The left side's last word at the room's left side, the right side's at the margin.)
    off = Math.max(off, Math.abs(ln.left[ln.left.length - 1] - z.left), Math.abs(ln.right[ln.right.length - 1] - ln.end));
  }
}
check(rooms.length === 2 && split > 6 && off < 2 && !inRoom, `justified beside both-sides shapes, each side justified on its own (${split} lines split, sides' ends ${off.toFixed(1)}px off, ${inRoom} words in a room)`);

// (wrap-squeeze.docx, Word's own: a rectangle beside each of 24 justified paragraphs, stepped across the column; some
// leave a side too narrow for a word, or one word, and some sit above their paragraph's text, under the one before.
// The spacers had come and gone until their passes ran out, words left in the rooms and the pages never laid out.)
await load('wrap-squeeze.docx');
await page.waitForFunction(() => document.querySelectorAll('.odx-sheet').length > 1, null, { timeout: 30000 }).catch(() => {});
const squeeze = await page.evaluate(() => {
  const v = window.odx.view;
  const st = v.state.plugins.find((pl) => pl.key.startsWith('odx-wrapboth'))?.getState(v.state);
  const rooms = [...document.querySelectorAll('.odx-content img[style*="--odx-both"]')].map((img) => {
    const r = img.getBoundingClientRect();
    const [, dr, , dl] = getComputedStyle(img).getPropertyValue('--odx-both').trim().split(/\s+/).map(parseFloat);
    return { left: r.left - dl, right: r.right + dr, top: r.top, bottom: r.bottom };
  });
  let inside = 0;
  const range = document.createRange();
  for (const p of document.querySelectorAll('.odx-content > p')) {
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      for (const m of t.data.matchAll(/\S+/g)) {
        range.setStart(t, m.index); range.setEnd(t, m.index + m[0].length);
        const w = range.getBoundingClientRect();
        const y = (w.top + w.bottom) / 2;
        if (rooms.some((z) => w.right > z.left + 1 && w.left < z.right - 1 && y > z.top && y < z.bottom)) inside++;
      }
    }
  }
  return { passes: st?.passes, spacers: st?.list.length, rooms: rooms.length, inside, sheets: document.querySelectorAll('.odx-sheet').length };
});
check(squeeze.rooms > 18 && squeeze.passes < 40 && squeeze.spacers > 60 && !squeeze.inside && squeeze.sheets > 1, `justified paragraphs beside 24 rectangles across the column (${squeeze.rooms} with text on both sides): the spacers settle (${squeeze.passes} passes, ${squeeze.spacers} spacers), no word in a room (${squeeze.inside}), the pages laid out (${squeeze.sheets})`);

await load('wrap-up-left.docx');
const upRoom = await page.evaluate(() => {
  const img = [...document.querySelectorAll('.odx-content img')].find((i) => getComputedStyle(i).float === 'right');
  const r = img.getBoundingClientRect();
  return { left: r.left - (parseFloat(getComputedStyle(img).marginLeft) || 0), right: r.right, top: r.top, bottom: r.bottom, translate: img.style.translate };
});
const above = await linesBeside('.odx-content > p:has(~ p img[style*="translate"])', upRoom);
const upOff = Math.max(0, ...above.filter((ln) => !ln.last && ln.left.length > 1).map((ln) => Math.abs(ln.left[ln.left.length - 1] - upRoom.left)));
check(/-/.test(upRoom.translate) && above.length > 2 && above.every((ln) => !ln.inside && !ln.right.length) && upOff < 2, `a float drawn up to the page's edge, wrapped on its left only: the lines above it kept to its left, justified to it (${above.length} lines, ${above.reduce((a, ln) => a + ln.inside + ln.right.length, 0)} words in or past its room, ends ${upOff.toFixed(1)}px off; ${upRoom.translate})`);

await load('footer-wrap-justified.docx');
await page.emulateMedia({ media: 'print' });
await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
const footRoom = await page.evaluate(() => {
  const r = document.querySelector('.odx-footer img').getBoundingClientRect();
  return { left: r.left - 12, right: r.right + 12, top: r.top, bottom: r.bottom };
});
const printed = await linesBeside('.odx-printlines p', footRoom);
await page.evaluate(() => window.dispatchEvent(new Event('afterprint')));
await page.emulateMedia({ media: 'screen' });
const footOff = Math.max(0, ...printed.filter((ln) => !ln.last && ln.left.length > 1).map((ln) => Math.abs(ln.left[ln.left.length - 1] - footRoom.left)));
check(printed.length > 2 && printed.every((ln) => !ln.inside && !ln.right.length) && footOff < 2, `printed, justified lines beside a footer's shape wrapped on its left only end at its room (${printed.length} lines, ${printed.reduce((a, ln) => a + ln.inside + ln.right.length, 0)} words in or past it, ends ${footOff.toFixed(1)}px off)`);
// Shapes placed from their paragraph's top under space before (sb-word.docx, Word-made: wrap-word5 mksbfix.ps1): Word
// measures the offset from the paragraph's top above that space, the space after the paragraph above not part of it,
// and from the page's top for one starting a page. In Word's PDF: under the filler line (no space after) at 0, a square
// one under a paragraph spaced 18pt before; 24pt under it (12 after, offset 12) one in front of the text under 18pt
// before; at 0 a tight one under 12pt before; at the page's top margin one spaced 18pt before starting page 2.
await page.setInputFiles('input[type=file]', 'test/fixtures/sb-word.docx');
await page.waitForFunction(() => window.odx?.fileName === 'sb-word.docx');
await page.waitForTimeout(2500);
const sbTops = await page.evaluate(() => [...document.querySelectorAll('.odx-content img:not(.ProseMirror-separator)')].map((img) => {
  const pad = parseFloat(getComputedStyle(img).getPropertyValue('--odx-pad')) || 0;
  const top = img.getBoundingClientRect().top + pad;
  const p = img.closest('p');
  const above = p.previousElementSibling;
  const sheet = [...document.querySelectorAll('.odx-sheets > .odx-sheet')].map((e) => e.getBoundingClientRect()).find((q) => p.getBoundingClientRect().top >= q.top && p.getBoundingClientRect().top < q.bottom);
  return { fromAbove: top - above.getBoundingClientRect().bottom, fromPage: sheet ? top - sheet.top - 96 : NaN };
}));
const sbWant = [['fromAbove', 0], ['fromAbove', 32], ['fromAbove', 0], ['fromPage', 0]];
check(sbTops.length === 4 && sbWant.every(([k, v], i) => Math.abs(sbTops[i][k] - v) < 1.5), `shapes under space before placed from their paragraph's top above it, as Word's (${sbTops.map((t, i) => t[sbWant[i]?.[0]]?.toFixed(1)).join(', ')}px; Word 0, 32, 0, 0)`);
check(!errors.length, `no page errors (${errors.join(' | ')})`);
await browser.close();
server.close();
console.log(failed ? `wrapfloat: ${failed} failed` : 'wrapfloat: all passed');
process.exit(failed ? 1 : 0);
