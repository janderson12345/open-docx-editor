// Browser check of the arrow keys at a right-to-left table's edges (test/fixtures/rtl-caret.docx, made in Word; the
// keys checked against Word's own, pressed for real): Left goes on through the cells, an English one too, and Right
// back, never stuck at a cell's end; Left and Right into the table from the paragraphs over and under it, not over
// the whole table; Up and Down to the cell above or below, out of the table, and into it from under it, the caret kept
// near its column, not put at the cell's start. A place is p:o, the paragraph's index (the cells in order) and the
// offset in its text. (The cells, not exact offsets, where fonts would move them.)
import { chromium, firefox, webkit } from 'playwright';
import { serveDemo, launchOptions } from './browser.mjs';

const { server } = serveDemo();
const engine = { firefox, webkit }[process.env.BROWSER];
const browser = engine ? await engine.launch() : await chromium.launch(launchOptions());
const page = await browser.newPage({ viewport: { width: 1300, height: 1100 } });
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
let failed = 0;
const check = (ok, what) => { console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}`); if (!ok) failed++; };

await page.goto(`http://localhost:${server.address().port}/`);
await page.waitForSelector('.odx-content p');
await page.setInputFiles('input[type=file]', 'test/fixtures/rtl-caret.docx');
await page.waitForFunction(() => window.odx?.fileName === 'rtl-caret.docx');
await page.waitForTimeout(2000);
await page.click('.odx-content p');
await page.evaluate(() => {
  const view = window.odx.view;
  const blocks = () => { const out = []; view.state.doc.descendants((n, pos) => { if (n.isTextblock) { out.push(pos + 1); return false; } return true; }); return out; };
  window.__at = {
    set: (p, o) => { const pos = blocks()[p] + o; const S = view.state.selection.constructor; view.focus(); view.dispatch(view.state.tr.setSelection(S.create(view.state.doc, pos))); },
    get: () => { view.domObserver?.flush?.(); const h = view.state.selection.head; const bs = blocks(); let i = bs.length - 1; while (i > 0 && bs[i] > h) i--; return { p: i, o: h - bs[i] }; },
  };
});
const at = (p, o) => page.evaluate(([a, b]) => window.__at.set(a, b), [p, o]);
const here = () => page.evaluate(() => window.__at.get());
const press = async (key, times = 1) => { const out = []; for (let i = 0; i < times; i++) { await page.keyboard.press(key); await page.waitForTimeout(30); out.push(await here()); } return out; };
const s = (x) => `${x.p}:${x.o}`;

// The table's cells are paragraphs 7-12 (row 1: 7 "תא אחד", 8 "cell two", 9 "תא 3 עם 45"; row 2: 10-12); 6 is the
// paragraph over it, 13 the one under it.
check((await page.evaluate(() => window.odx.view.state.doc.child(7).type.name)) === 'table', 'the table is where the check expects it');
await at(7, 0);
let steps = await press('ArrowLeft', 30);
check(steps[5].p === 7 && steps[5].o === 6 && s(steps[6]) === '8:0', `Left at the first cell's end goes on into the next cell (${steps.slice(4, 8).map(s).join(' ')})`);
check(s(steps[14]) === '8:8' && s(steps[15]) === '9:0', `Left through the English cell to the third (${steps.slice(13, 17).map(s).join(' ')})`);
check(steps.at(-1).p >= 10, `Left goes on into the second row (${s(steps.at(-1))})`);
steps = await press('ArrowRight', 40);
const back = steps.findIndex((x) => s(x) === '7:0');
check(back > 0 && s(steps[back + 1]) === '6:296', `Right back through the cells and out over the table (${steps.slice(back - 1, back + 2).map(s).join(' ')})`);
await at(13, 0);
steps = await press('ArrowRight', 2);
check(steps[0].p === 12 && steps[0].o === 3, `Right from the paragraph under the table goes into its last cell, not over the table (${steps.map(s).join(' ')})`);
await at(6, 296);
steps = await press('ArrowLeft');
check(s(steps[0]) === '7:0', `Left at the end of the paragraph over the table goes into its first cell (${s(steps[0])})`);
await at(8, 3);
steps = await press('ArrowDown', 2);
check(steps[0].p === 11 && steps[0].o > 0, `Down from the second cell to the cell under it, near the caret's column (${s(steps[0])})`);
check(steps[1].p === 13 && steps[1].o > 0, `Down out of the table, near the caret's column (${s(steps[1])})`);
steps = await press('ArrowUp', 3);
check(steps[0].p === 11 && steps[1].p === 8 && steps[2].p === 6, `Up into the table, the row above, and out over it (${steps.map(s).join(' ')})`);
await at(13, 0);
steps = await press('ArrowUp');
check(steps[0].p === 10, `Up from the start of the paragraph under the table goes to the cell over it on the right (${s(steps[0])})`);

check(!errors.length, `no page errors${errors.length ? ': ' + errors.join(' | ') : ''}`);
await browser.close();
server.close();
if (failed) { console.log(`${failed} failed`); process.exit(1); }
