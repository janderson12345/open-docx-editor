// Browser check of resizing shapes with the mouse (test/fixtures/mkshapeedit.mjs): a shape's corner resizes it freely,
// Shift keeping its proportions; its yellow handles change its adjust values (all of them written, as Word needs); a
// line has a handle at each end, which moves that end (flipping the line where it crosses the other, Shift at 15
// degree steps), a floating line's offsets following; a callout's pointer is drawn past its box. Undo, and the saved
// file (written to test/out/shape-edit-edited.docx, to open in Word).
import { chromium, firefox, webkit } from 'playwright';
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
// (BROWSER=firefox or webkit: in those engines, Playwright's own builds.)
const engine = { firefox, webkit }[process.env.BROWSER];
const browser = engine ? await engine.launch() : await chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : process.env.BROWSER_CHANNEL ? { channel: process.env.BROWSER_CHANNEL } : {});
const page = await browser.newPage({ viewport: { width: 1300, height: 950 } });
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
let failed = 0;
const check = (ok, what) => { console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}`); if (!ok) failed++; };

await page.goto(`http://localhost:${server.address().port}/`);
await page.waitForSelector('.odx-content p');
await page.setInputFiles('input[type=file]', 'test/fixtures/shape-edit.docx');
await page.waitForFunction(() => window.odx?.fileName === 'shape-edit.docx');
await page.waitForTimeout(800);

// The shapes' nodes, in document order: size, place and XML.
const shapes = () => page.evaluate(() => {
  const out = [];
  window.odx.view.state.doc.descendants((n, pos) => { if (n.type.name === 'image' && /prstGeom/.test(n.attrs.xml)) out.push({ pos, w: n.attrs.width, h: n.attrs.height, xml: n.attrs.xml, place: n.attrs.place || '' }); });
  return out;
});
const imgs = page.locator('.odx-content img[src^="data:image/svg"]');
const select = async (i) => { await imgs.nth(i).click(); await page.waitForTimeout(250); };
const drag = async (sel, dx, dy, mod) => {
  const b = await page.locator(sel).boundingBox();
  await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2);
  if (mod) await page.keyboard.down(mod);
  await page.mouse.down();
  await page.mouse.move(b.x + b.width / 2 + dx, b.y + b.height / 2 + dy, { steps: 6 });
  await page.waitForTimeout(120);
  await page.mouse.up();
  if (mod) await page.keyboard.up(mod);
  await page.waitForTimeout(400);
};
const avOf = (xml) => [...(/<a:avLst>([\s\S]*?)<\/a:avLst>/.exec(xml)?.[1] || '').matchAll(/name="(\w+)" fmla="val (-?\d+)"/g)].map((m) => `${m[1]}=${m[2]}`).join(' ');
const xfrmOf = (xml) => /<wps:spPr><a:xfrm([^>]*)>/.exec(xml)?.[1].trim() || '';
const off = (xml, axis) => Number(new RegExp(`<wp:position${axis}\\b[^>]*>\\s*<wp:posOffset>(-?\\d+)`).exec(xml)?.[1]);
const s0 = await shapes();
check(s0.length === 5, `five shapes (${s0.length})`);

// 1. The arrow: a corner resizes it freely, Shift keeps its proportions; its yellow handles; undo.
await select(0);
check(await page.locator('.odx-imgsize-adj').count() === 2, 'an arrow has two yellow handles');
await drag('.odx-imgsize-se', 60, 0);
let [a1] = await shapes();
check(a1.w === s0[0].w + 60 && a1.h === s0[0].h, `its corner makes it wider only (${s0[0].w}x${s0[0].h} -> ${a1.w}x${a1.h})`);
await drag('.odx-imgsize-se', 39, 0, 'Shift');
const [a2] = await shapes();
check(Math.abs(a2.w / a2.h - a1.w / a1.h) < 0.03 && a2.w > a1.w + 30, `Shift keeps its proportions (${a1.w}x${a1.h} -> ${a2.w}x${a2.h})`);
await drag('.odx-imgsize-adj >> nth=1', -40, 0);
const [a3] = await shapes();
// (Its head: adj2 of the shorter side; 40px more of it.)
const head = (x) => Number(/name="adj2" fmla="val (\d+)"/.exec(x)?.[1] ?? 50000) * Math.min(a2.w, a2.h) / 100000;
check(Math.abs(head(a3.xml) - head(a2.xml) - 40) < 2, `its head handle makes the head 40px longer (${avOf(a3.xml)})`);
check(avOf(a3.xml).startsWith('adj1=50000 adj2='), 'every adjust value written, in order (Word won\'t open a file with only some)');
await page.keyboard.press('Control+z');
await page.waitForTimeout(300);
check(avOf((await shapes())[0].xml) === '', 'undo takes the adjust back');

// 2. The callout: its pointer drawn past its box (the picture larger round it, the box where it was).
await select(1);
const c0 = (await shapes())[1];
check(/--odx-pad:\d+px/.test(c0.place), `its pointer, below its box, has room drawn for it (${/--odx-pad:\d+px/.exec(c0.place)?.[0]})`);
await drag('.odx-imgsize-adj', 40, 30);
const c1 = (await shapes())[1];
check(avOf(c1.xml) === 'adj1=20833 adj2=104167 adj3=16667', `its pointer handle moves the pointer as far as the mouse (${avOf(c1.xml)})`);

// 3. The inline line (144x1): two end handles, no box; its end dragged past its start flips it.
await select(2);
check(await page.locator('.odx-imgsize-end').count() === 2 && await page.locator('.odx-imgsize-h:visible').count() === 0, 'a line has a handle at each end and no others');
await drag('.odx-imgsize-end >> nth=1', -40, 80);
const l1 = (await shapes())[2];
check(l1.w === 104 && l1.h === 81 && xfrmOf(l1.xml) === '', `its end dragged down and in: ${l1.w}x${l1.h}, unflipped (${xfrmOf(l1.xml)})`);
await drag('.odx-imgsize-end >> nth=0', 200, 0);
const l2 = (await shapes())[2];
check(l2.w === 96 && l2.h === 81 && xfrmOf(l2.xml) === 'flipH="1"', `its start dragged past its end: ${l2.w}x${l2.h}, ${xfrmOf(l2.xml)}`);
await drag('.odx-imgsize-end >> nth=0', 3, -24, 'Shift');
const l3 = (await shapes())[2];
const ang = Math.atan2(l3.h, l3.w) * 180 / Math.PI;
check(Math.abs(ang / 15 - Math.round(ang / 15)) < 0.05, `Shift: at a multiple of 15 degrees (${ang.toFixed(1)})`);

// 4. The floating line: its start dragged up and left moves its offsets with it, its end staying.
await select(3);
const f0 = (await shapes())[3];
const e0 = await page.locator('.odx-imgsize-end >> nth=1').boundingBox();
await drag('.odx-imgsize-end >> nth=0', -30, -20);
const f1 = (await shapes())[3];
check(f1.w === f0.w + 30 && f1.h === f0.h + 20, `the floating line's start moved: ${f0.w}x${f0.h} -> ${f1.w}x${f1.h}`);
check(Math.abs(off(f1.xml, 'H') - (off(f0.xml, 'H') - 30 * 9525)) <= 9525 && Math.abs(off(f1.xml, 'V') - (off(f0.xml, 'V') - 20 * 9525)) <= 9525, 'and its offsets with it');
const e1 = await page.locator('.odx-imgsize-end >> nth=1').boundingBox();
check(Math.abs(e1.x - e0.x) < 2 && Math.abs(e1.y - e0.y) < 2, `its end stays where it was (${Math.round(e1.x - e0.x)}, ${Math.round(e1.y - e0.y)})`);

// 5. The floating rounded rectangle: its corner handle.
await select(4);
const r0 = (await shapes())[4];
await drag('.odx-imgsize-adj', 20, 0);
const r1 = (await shapes())[4];
check(/adj=\d+/.test(avOf(r1.xml)) && Number(/adj=(\d+)/.exec(avOf(r1.xml))[1]) > 16667, `a rounded rectangle's handle rounds its corners more (${avOf(r0.xml) || 'default'} -> ${avOf(r1.xml)})`);

// 6. Turning: the arrow's rotate handle dragged to the right of its centre turns it to 90 degrees; the picture turned
// about its centre, the room round it on its line written (effectExtent) as Word writes it for a turned shape.
const rotOf = (xml) => Number(/<wps:spPr><a:xfrm[^>]*\brot="(\d+)"/.exec(xml)?.[1] || 0) / 60000;
const roomOf = (xml) => /<wp:effectExtent l="(\d+)" t="(\d+)" r="(\d+)" b="(\d+)"/.exec(xml).slice(1).map((v) => Math.round(v / 9525)).join(' ');
await select(0);
const t0 = (await shapes())[0];
const centre = await imgs.nth(0).boundingBox();
const rh = await page.locator('.odx-imgsize-rot').boundingBox();
check(rh && rh.y + rh.height < centre.y, 'a shape has a rotate handle above it');
await page.mouse.move(rh.x + rh.width / 2, rh.y + rh.height / 2);
await page.mouse.down();
await page.mouse.move(centre.x + centre.width / 2 + 150, centre.y + centre.height / 2 + 3, { steps: 8 });
await page.mouse.up();
await page.waitForTimeout(500);
const t1 = (await shapes())[0];
check(Math.abs(rotOf(t1.xml) - 90) <= 2, `dragged to its right, it turns to 90 degrees (${rotOf(t1.xml)})`);
// (Drawn its left and top room along from its box, as Word draws a turned shape in line with the text.)
check(/transform:translate\(0px, (\d+)px\) rotate\(\d+(\.\d+)?deg\)/.test(t1.place), `its picture turned by the page (${/transform:[^;]*/.exec(t1.place)?.[0]})`);
// (A 96 x 72 shape at 90 degrees: taller than its box by 24px, so 12px (and half its outline) top and bottom, none at its sides.)
const [rl1, tt1, rr1, b1] = roomOf(t1.xml).split(' ').map(Number);
check(rl1 === 0 && rr1 === 0 && tt1 > 0 && tt1 === b1, `room round it for its turned outline, as Word writes it (${roomOf(t1.xml)}; was ${roomOf(t0.xml)})`);
// Shift: 15-degree steps; the keyboard: Alt+Right 15 more, Ctrl+Alt+Left 1 back.
const rh2 = await page.locator('.odx-imgsize-rot').boundingBox();
const c2 = await page.locator('.odx-imgsize').boundingBox();
await page.mouse.move(rh2.x + rh2.width / 2, rh2.y + rh2.height / 2);
await page.keyboard.down('Shift');
await page.mouse.down();
await page.mouse.move(c2.x + c2.width / 2 + 100, c2.y + c2.height / 2 + 63, { steps: 8 }); // (about 122 degrees)
await page.mouse.up();
await page.keyboard.up('Shift');
await page.waitForTimeout(400);
const t2 = (await shapes())[0];
check(rotOf(t2.xml) % 15 === 0 && Math.abs(rotOf(t2.xml) - 120) <= 15, `with Shift, a multiple of 15 degrees (${rotOf(t2.xml)})`);
await page.keyboard.press('Alt+ArrowRight');
await page.keyboard.press('Control+Alt+ArrowLeft');
const t3 = (await shapes())[0];
check(rotOf(t3.xml) === rotOf(t2.xml) + 14, `Alt+Right turns it 15 more, Ctrl+Alt+Left 1 back (${rotOf(t2.xml)} -> ${rotOf(t3.xml)})`);
// A turned shape's handles turn with it: its side handle makes it longer along its own length.
// (Undo, step by step, back to the first turn: two quick keys may be one step.)
for (let k = 0; k < 4 && rotOf((await shapes())[0].xml) !== rotOf(t1.xml); k++) { await page.keyboard.press('Control+z'); await page.waitForTimeout(200); }
const t4 = (await shapes())[0];
check(rotOf(t4.xml) === rotOf(t1.xml) && roomOf(t4.xml) === roomOf(t1.xml), `undo takes the turns back (${rotOf(t4.xml)})`);
await select(0);
await drag('.odx-imgsize-e', 0, 50); // (its right side, turned to face down: dragged down)
const t5 = (await shapes())[0];
check(Math.abs(t5.w - t4.w - 50) <= 2 && t5.h === t4.h, `turned to 90 degrees, its side handle dragged down makes it longer (${t4.w}x${t4.h} -> ${t5.w}x${t5.h})`);
const [, tt5] = roomOf(t5.xml).split(' ').map(Number);
check(Math.abs(tt5 - tt1 - 25) <= 1, `and its room round it follows (${roomOf(t5.xml)})`);

// 7. The floating rectangle turned: its offsets as they were (it turns about its centre), its picture turned.
await select(4);
const g0 = (await shapes())[4];
await page.keyboard.press('Alt+ArrowRight');
await page.keyboard.press('Alt+ArrowRight');
const g1 = (await shapes())[4];
check(rotOf(g1.xml) === 30 && off(g1.xml, 'H') === off(g0.xml, 'H') && off(g1.xml, 'V') === off(g0.xml, 'V'), `a floating shape turns about its centre, its place as it was (${rotOf(g1.xml)})`);
check(/rotate\(30deg\)/.test(g1.place), 'its picture turned on the page');

// The saved file: the shapes as changed.
const bytes = await page.evaluate(async () => Array.from(await window.odx.save()));
fs.mkdirSync('test/out', { recursive: true });
fs.writeFileSync('test/out/shape-edit-edited.docx', Buffer.from(bytes));
const saved = strFromU8(unzipSync(new Uint8Array(bytes))['word/document.xml']);
check(saved.includes('<a:gd name="adj1" fmla="val 20833"/><a:gd name="adj2" fmla="val 104167"/><a:gd name="adj3" fmla="val 16667"/>') && saved.includes('flipH="1"'), 'saved with the callout\'s pointer and the flipped line');
check(!errors.length, `no page errors (${errors.join(' | ')})`);
await browser.close();
server.close();
console.log(failed ? `shapeedit: ${failed} failed` : 'shapeedit: all passed');
process.exit(failed ? 1 : 0);
