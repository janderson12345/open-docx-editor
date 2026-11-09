// Browser check of moving and resizing text boxes: a handle drag resizes one (its frame and shape, its text wrapped
// again, not stretched), dragging a floating one moves it (its offsets), and the keyboard does both (arrows move a
// floating box, Ctrl for 1px; Shift+arrows resize; Esc goes back to the text). Undo, save, and the saved XML.
import { chromium } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';

const root = path.resolve('demo');
const server = http.createServer((req, res) => {
  const p = path.join(root, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
  if (!p.startsWith(root) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'content-type': { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' }[path.extname(p)] || 'application/octet-stream' });
  fs.createReadStream(p).pipe(res);
}).listen(0);
const launch = process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : process.env.BROWSER_CHANNEL ? { channel: process.env.BROWSER_CHANNEL } : {};
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
check(errors.length === 0, `no page errors ${errors.join('; ')}`);

await browser.close();
server.close();
console.log(failed ? `${failed} failures` : 'textboxes: all passed');
process.exit(failed ? 1 : 0);
