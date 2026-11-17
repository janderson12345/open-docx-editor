// Browser check of moving and resizing text boxes: a handle drag resizes one (its frame and shape, its text wrapped
// again, not stretched), dragging a floating one moves it (its offsets), and the keyboard does both (arrows move a
// floating box, Ctrl for 1px; Shift+arrows resize; Esc goes back to the text). Undo, save, and the saved XML.
import { chromium } from 'playwright';
import fs from 'node:fs';
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
await page.setInputFiles('input[type=file]', 'test/fixtures/textbox.docx');
await page.waitForFunction(() => window.odx?.fileName === 'textbox.docx');
await page.waitForTimeout(800);
// The two boxes' nodes: positions, sizes and XML, in document order.
const boxes = () => page.evaluate(() => {
  const out = [];
  window.odx.view.state.doc.descendants((n, pos) => { if (n.type.name === 'image' && /txbxContent/.test(n.attrs.xml)) out.push({ pos, w: n.attrs.width, h: n.attrs.height, xml: n.attrs.xml, src: n.attrs.src }); });
  return out;
});
const off = (xml, axis) => Number(new RegExp(`<wp:position${axis}\\b[^>]*>\\s*<wp:posOffset>(-?\\d+)`).exec(xml)?.[1]);
const imgs = page.locator('.odx-content img[src^="data:image/svg"]');
const [inline0, float0] = await boxes();
check(inline0 && float0, 'two text boxes, one floating');

// 1. Resize the inline box by its right handle.
await imgs.nth(0).click();
const handle = page.locator('.odx-imgsize-e');
const hb = await handle.boundingBox();
await page.mouse.move(hb.x + hb.width / 2, hb.y + hb.height / 2);
await page.mouse.down();
await page.mouse.move(hb.x + 120, hb.y + hb.height / 2, { steps: 5 });
await page.mouse.up();
await page.waitForTimeout(400);
let [inline1] = await boxes();
check(Math.abs(inline1.w - (inline0.w + 120 - hb.width / 2)) < 6 && inline1.h === inline0.h, `the handle drag makes it wider (${inline0.w} -> ${inline1.w})`);
const cx = (xml) => [...xml.matchAll(/<wp:extent cx="(\d+)"|<a:ext cx="(\d+)"/g)].map((m) => Number(m[1] || m[2]));
check(cx(inline1.xml).every((v) => Math.abs(v - inline1.w * 9525) < 9525), 'its frame and its shape take the new width');
check(/width:([\d.]+)pt/.exec(inline1.xml)[1] !== /width:([\d.]+)pt/.exec(inline0.xml)[1], 'and its VML copy');
check(inline1.src !== inline0.src, 'it is drawn again (its text wrapped to the width, not stretched)');

// 2. Drag the floating box by its body.
await imgs.nth(1).click();
await page.waitForTimeout(200);
const fb = await imgs.nth(1).boundingBox();
const before = fb;
await page.mouse.move(fb.x + fb.width / 2, fb.y + fb.height / 2);
await page.mouse.down();
await page.mouse.move(fb.x + fb.width / 2 - 60, fb.y + fb.height / 2 + 40, { steps: 6 });
await page.mouse.up();
await page.waitForTimeout(600);
let [, float1] = await boxes();
check(Math.abs(off(float1.xml, 'H') - (off(float0.xml, 'H') - 60 * 9525)) < 9525 * 2 && Math.abs(off(float1.xml, 'V') - (off(float0.xml, 'V') + 40 * 9525)) < 9525 * 2, 'dragging moves its offsets');
const after = await imgs.nth(1).boundingBox();
check(Math.abs(after.x - (before.x - 60)) < 3 && Math.abs(after.y - (before.y + 40)) < 3, `and the box on the page (${Math.round(after.x - before.x)}, ${Math.round(after.y - before.y)})`);

// 3. The keyboard: arrows move it, Ctrl for 1px; Shift+arrows resize; Esc back to the text.
await page.keyboard.press('ArrowRight');
await page.keyboard.press('Control+ArrowDown');
let [, float2] = await boxes();
check(off(float2.xml, 'H') === off(float1.xml, 'H') + 8 * 9525 && off(float2.xml, 'V') === off(float1.xml, 'V') + 9525, 'ArrowRight moves it 8px, Ctrl+ArrowDown 1px');
await page.keyboard.press('Shift+ArrowRight');
await page.keyboard.press('Shift+ArrowDown');
let [, float3] = await boxes();
check(float3.w === float2.w + 8 && float3.h === float2.h + 8, 'Shift+arrows make it 8px wider and taller');
check(await page.evaluate(() => window.odx.view.state.selection.node?.type.name === 'image'), 'it stays selected');
// The selected floating box still opens its menu on a right-click and its text on a double-click (the move overlay is
// over it).
const fb2 = await imgs.nth(1).boundingBox();
await page.mouse.click(fb2.x + fb2.width / 2, fb2.y + fb2.height / 2, { button: 'right' });
await page.waitForTimeout(300);
check(await page.evaluate(() => { const m = document.querySelector('.odx-ctx'); return !!m && !m.hidden; }), 'a right-click on it opens the menu');
await page.keyboard.press('Escape');
await page.evaluate(() => window.odx.view.focus());
await page.mouse.dblclick(fb2.x + fb2.width / 2, fb2.y + fb2.height / 2);
await page.waitForTimeout(400);
check(!!(await page.$('.odx-hfedit-content')), 'a double-click on it edits its text');
await page.keyboard.press('Escape');
await page.waitForTimeout(300);
await page.mouse.click(fb2.x + fb2.width / 2, fb2.y + fb2.height / 2); // (through the move overlay, as a person clicks)
await page.keyboard.press('Escape');
check(await page.evaluate(() => !window.odx.view.state.selection.node && window.odx.view.state.selection.empty), 'Esc goes back to the text');

// 4. Undo all of it, step by step, then redo; save.
for (let i = 0; i < 6; i++) await page.keyboard.press('Control+z');
await page.waitForTimeout(300);
const undone = await boxes();
check(undone[0].w === inline0.w && off(undone[1].xml, 'H') === off(float0.xml, 'H'), 'undo puts both back');
for (let i = 0; i < 6; i++) await page.keyboard.press('Control+y');
await page.waitForTimeout(300);
const redone = await boxes();
check(redone[1].w === float3.w && off(redone[1].xml, 'V') === off(float3.xml, 'V'), 'redo the moves and sizes');
const bytes = await page.evaluate(async () => Array.from(await window.odx.save()));
fs.mkdirSync('test/out', { recursive: true });
fs.writeFileSync('test/out/textbox-moved.docx', Buffer.from(bytes));
const xml = strFromU8(unzipSync(new Uint8Array(bytes))['word/document.xml']);
check(xml.includes(`<wp:posOffset>${off(float3.xml, 'H')}</wp:posOffset>`) && xml.includes(`cx="${Math.round(float3.w * 9525)}"`), 'the saved file has the new place and size');

// 5. Text measured before its stand-in font loaded: smartart.docx (Aptos, a web stand-in here) read in a fresh page,
// where the boxes' words are measured in the fallback. Once the faces load, each is drawn again on the page, as a
// reload with the fonts in draws it (lines wrapped at the stand-in's widths), and that's not an edit.
const fresh = await browser.newPage({ viewport: { width: 1300, height: 950 } });
fresh.on('pageerror', (e) => errors.push(e.message));
await fresh.goto(`http://localhost:${server.address().port}/`);
await fresh.waitForSelector('.odx-content p');
// Each shape's picture with text: as the node has it, and as the page shows it (its embedded fonts left out).
const pictures = () => fresh.evaluate(() => {
  const svg = (s) => decodeURIComponent(s.slice(s.indexOf(',') + 1)).replace(/<style>[\s\S]*?<\/style>/, '');
  const v = window.odx.view, out = [];
  v.state.doc.descendants((n, pos) => {
    if (n.type.name !== 'image' || !/^data:image\/svg\+xml/.test(n.attrs.src || '') || !/%3Ctspan/.test(n.attrs.src)) return;
    out.push({ node: svg(n.attrs.src), shown: svg(v.nodeDOM(pos)?.getAttribute?.('src') || '') });
  });
  return out;
});
const openIn = async (file) => {
  await fresh.setInputFiles('input[type=file]', []);
  await fresh.setInputFiles('input[type=file]', file);
  await fresh.waitForFunction(() => window.odx?.fileName === 'smartart.docx' && window.odx.view.state.doc.textContent.length > 0);
  await fresh.evaluate(() => { window.__opened = window.odx.view.state.doc; });
  await fresh.waitForTimeout(3000);
};
await openIn('test/fixtures/smartart.docx');
const first = await pictures();
const quiet = await fresh.evaluate(() => {
  const st = window.odx.view.state;
  const history = st.plugins.map((p) => p.getState?.(st)).find((s) => s && s.done && s.undone);
  return { same: st.doc === window.__opened, dirty: window.odx.dirty, undo: history ? history.done.eventCount : 0 };
});
await fresh.evaluate(() => window.odx.newDocument());
await fresh.waitForTimeout(300);
await openIn('test/fixtures/smartart.docx');
const again = await pictures();
check(first.length > 0 && first.length === again.length, `smartart.docx: ${first.length} boxes with text`);
check(first.some((p) => p.shown !== p.node), `boxes measured before the font loaded are drawn again (${first.filter((p) => p.shown !== p.node).length} of ${first.length})`);
check(first.every((p, i) => p.shown === again[i]?.node), 'each as a reload with the fonts loaded draws it');
check(again.every((p) => p.shown === p.node), 'and not again after that reload (measured right the first time)');
check(quiet.same && !quiet.dirty && quiet.undo === 0, `the redraw is no edit: the document unchanged, not dirty, no undo step (${JSON.stringify(quiet)})`);
await fresh.close();

// A text box with an 8pt outline (shapeclip-word.docx, made in Word): its text, opened for editing, where its picture
// draws it, inset from the inside of the outline (shapes.js textBox), so it doesn't move as editing starts. Each first
// letter's box on the screen: the drawn one (its picture's SVG put in the page over the picture, getExtentOfChar), the
// editor's (a range round its first letter).
const ob = await browser.newPage({ viewport: { width: 1300, height: 950 } });
ob.on('pageerror', (e) => errors.push(e.message));
await ob.goto(`http://localhost:${server.address().port}/`);
await ob.waitForSelector('.odx-content p');
await ob.setInputFiles('input[type=file]', 'test/fixtures/shapeclip-word.docx');
await ob.waitForFunction(() => window.odx?.fileName === 'shapeclip-word.docx');
await ob.waitForTimeout(1500);
const outlined = () => ob.evaluate(() => {
  let pos = null;
  window.odx.view.state.doc.descendants((n, p) => { if (pos == null && n.type.name === 'image' && /8pt outline/.test(n.attrs.xml)) pos = p; });
  return pos;
});
const boxPos = await outlined();
const drawn = await ob.evaluate((pos) => {
  const img = window.odx.view.nodeDOM(pos);
  img.scrollIntoView({ block: 'center' });
  const r = img.getBoundingClientRect();
  const src = img.getAttribute('src');
  const holder = document.createElement('div');
  holder.id = 'odx-test-drawn';
  holder.style.cssText = `position:fixed;left:${r.left}px;top:${r.top}px;width:${r.width}px;height:${r.height}px;pointer-events:none;opacity:0`;
  holder.innerHTML = decodeURIComponent(src.slice(src.indexOf(',') + 1));
  document.body.append(holder);
  const text = [...holder.querySelectorAll('text')].find((t) => t.textContent.trim());
  const e = text.getExtentOfChar(0);
  const m = text.getScreenCTM();
  const at = (x, y) => ({ x: m.a * x + m.c * y + m.e, y: m.b * x + m.d * y + m.f });
  const p0 = at(e.x, e.y), p1 = at(e.x + e.width, e.y + e.height);
  holder.remove();
  return { left: p0.x, top: p0.y, bottom: p1.y, ch: text.textContent.trim()[0], img: { x: r.left, y: r.top, w: r.width, h: r.height } };
}, boxPos);
await ob.mouse.dblclick(drawn.img.x + drawn.img.w / 2, drawn.img.y + drawn.img.h / 2);
await ob.waitForSelector('.odx-txbxedit .ProseMirror');
await ob.waitForTimeout(300);
const edited = await ob.evaluate(() => {
  const pm = document.querySelector('.odx-txbxedit .ProseMirror');
  const walk = document.createTreeWalker(pm, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.data.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP) });
  const t = walk.nextNode();
  const i = t.data.search(/\S/);
  const range = document.createRange();
  range.setStart(t, i);
  range.setEnd(t, i + 1);
  const r = range.getBoundingClientRect();
  return { left: r.left, top: r.top, bottom: r.bottom, ch: t.data[i] };
});
await ob.keyboard.press('Escape');
const dx = edited.left - drawn.left, dy = (edited.top + edited.bottom) / 2 - (drawn.top + drawn.bottom) / 2;
check(drawn.ch === edited.ch && Math.abs(dx) < 1 && Math.abs(dy) < 1.5,
  `a text box with an 8pt outline: its text opened for editing where its picture draws it (first letter ${dx.toFixed(2)}px across, ${dy.toFixed(2)}px down; ${JSON.stringify({ drawn, edited })})`);
await ob.close();
check(errors.length === 0, `no page errors ${errors.join('; ')}`);

await browser.close();
server.close();
console.log(failed ? `${failed} failures` : 'textboxes: all passed');
process.exit(failed ? 1 : 0);
