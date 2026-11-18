// Co-editing in the browser (run with the e2e tests): two editors in one room through a local y-websocket server.
// A opens a document, B gets it; each one's typing reaches the other; each one's undo takes back only its own edit;
// A sees B's cursor; and what B saves has both editors' text.
import { chromium } from 'playwright';
import { build } from 'esbuild';
import { createRequire } from 'node:module';
import { WebSocketServer } from 'ws';
import { unzipSync, strFromU8 } from 'fflate';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { launchOptions } from './browser.mjs';

const require = createRequire(import.meta.url);
const { setupWSConnection } = require('y-websocket/bin/utils');

const out = path.resolve('test/out/collab');
fs.mkdirSync(out, { recursive: true });
fs.copyFileSync('src/editor/editor.css', path.join(out, 'editor.css'));
await build({ entryPoints: ['test/collab/app.js'], bundle: true, format: 'esm', outfile: path.join(out, 'app.js'), logLevel: 'error' });
fs.writeFileSync(path.join(out, 'index.html'), '<!doctype html><meta charset="utf-8"><link rel="icon" href="data:,"><link rel="stylesheet" href="editor.css"><div id="app"></div><script type="module" src="app.js"></script>');

const server = http.createServer((req, res) => {
  const p = path.join(out, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
  if (!p.startsWith(out) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'content-type': { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' }[path.extname(p)] || 'application/octet-stream' });
  fs.createReadStream(p).pipe(res);
}).listen(0);
const wss = new WebSocketServer({ server });
wss.on('connection', (conn, req) => setupWSConnection(conn, req));
const port = server.address().port;
const launch = launchOptions();
const browser = await chromium.launch(launch);
let failed = 0;
const check = (ok, msg) => { if (!ok) { failed++; console.log('FAIL', msg); } else console.log('ok  ', msg); };
const errors = [];
const open = async (name) => {
  const page = await browser.newPage({ viewport: { width: 1300, height: 900 } });
  page.on('pageerror', (e) => errors.push(`${name}: ${e.message}`));
  await page.goto(`http://localhost:${port}/?port=${port}&room=test-${Date.now()}&name=${name}`.replace(/test-\d+/, room));
  await page.waitForFunction(() => window.odx?.view, null, { timeout: 15000 });
  return page;
};
const room = `room-${Date.now()}`;
const text = (page) => page.evaluate(() => window.odx.view.state.doc.textContent);
const until = (page, fn, arg) => page.waitForFunction(fn, arg, { timeout: 10000 }).then(() => true, () => false);

const a = await open('A');
// A opens a document: it becomes the room's.
await a.evaluate((bytes) => window.odx.open(new Uint8Array(bytes)), [...fs.readFileSync('test/fixtures/form.docx')]);
const docText = await text(a);
const b = await open('B');
check(await until(b, (t) => window.odx.view.state.doc.textContent === t, docText), 'B joins with the document A opened');
// Each one's typing reaches the other.
await a.evaluate(() => { const v = window.odx.view; v.focus(); v.dispatch(v.state.tr.insertText('ALPHA ', 1)); });
check(await until(b, () => window.odx.view.state.doc.textContent.startsWith('ALPHA ')), "A's typing shows in B");
await b.evaluate(() => { const v = window.odx.view; v.focus(); const end = v.state.doc.firstChild.nodeSize - 1; v.dispatch(v.state.tr.insertText(' BETA', end)); });
check(await until(a, () => window.odx.view.state.doc.textContent.includes(' BETA')), "B's typing shows in A");
// A sees B's cursor.
check(await until(a, () => !!document.querySelector('.ProseMirror-yjs-cursor')), "A shows B's cursor");
// B saves: both editors' text is in the file.
const saved = Buffer.from(await b.evaluate(() => [...window.odx.save()]));
const xml = strFromU8(unzipSync(new Uint8Array(saved))['word/document.xml']);
check(xml.includes('ALPHA') && xml.includes('BETA'), "B's saved file has both editors' text");
// A's undo takes back A's own edit, not B's.
await a.waitForTimeout(600); // (y-prosemirror groups edits made within half a second into one undo step)
await a.evaluate(() => { const v = window.odx.view; v.odx.undo(v.state, v.dispatch, v); });
check(await until(b, () => !window.odx.view.state.doc.textContent.includes('ALPHA') && window.odx.view.state.doc.textContent.includes(' BETA')), "A's undo removes only A's edit, everywhere");
// Comments: B adds one and writes it; A sees it and answers; B sees the answer; A resolves the thread, B sees that;
// B deletes it, and it's gone from A too, and from what A saves.
const thread = (page) => page.evaluate(() => JSON.stringify(window.odx.view.odx.session.comments.map((c) => [c.text, c.parentId != null, c.done])));
const cid = await b.evaluate(() => {
  const v = window.odx.view;
  v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.create(v.state.doc, 2, 6)));
  const c = window.cm.addComment(v);
  window.cm.setText(v, c.id, 'Please check this');
  return c.id;
});
check(await until(a, (id) => window.odx.view.odx.session.comments.some((c) => c.id === id && c.text === 'Please check this'), cid), "B's comment shows in A");
check(await until(a, () => !!document.querySelector('.odx-balloons, .odx-markup') && document.body.textContent.includes('Please check this')), "... in A's comment pane");
await a.evaluate((id) => window.cm.replyTo(window.odx.view, id, 'Checked, fine'), cid);
check(await until(b, () => window.odx.view.odx.session.comments.some((c) => c.parentId != null && c.text === 'Checked, fine')), "A's reply shows in B");
await a.evaluate((id) => window.cm.setResolved(window.odx.view, id, true), cid);
check(await until(b, (id) => window.odx.view.odx.session.comments.find((c) => c.id === id)?.done === true, cid), "A resolving the thread shows in B");
await b.evaluate((id) => window.cm.deleteComment(window.odx.view, id), cid);
check(await until(a, () => window.odx.view.odx.session.comments.length === 0), `B deleting the thread removes it in A ${await thread(a)}`);
const savedA = strFromU8(unzipSync(new Uint8Array(Buffer.from(await a.evaluate(() => [...window.odx.save()]))))['word/document.xml']);
check(!savedA.includes('commentReference'), "A's saved file has no comment left");
// Headers and notes: A opens a document with a header (it becomes the room's); A edits the header and adds a footnote;
// B shows both once A closes them.
await a.evaluate((bytes) => window.odx.open(new Uint8Array(bytes), 'long.docx'), [...fs.readFileSync('test/fixtures/long.docx')]);
check(await until(b, () => window.odx.fileName === 'long.docx' || window.odx.view.state.doc.childCount > 20), 'B follows A to the new document');
await a.waitForTimeout(800);
await a.evaluate(() => { const v = window.odx.view; window.hf.editStory(v, 'header'); const h = v.odx.hfView; h.dispatch(h.state.tr.insertText('SHARED HEADER ', 1)); window.hf.closeStory(); });
check(await until(b, () => [...document.querySelectorAll('.odx-hf')].some((e) => e.textContent.includes('SHARED HEADER'))), "A's header edit shows in B's pages");
await a.evaluate(() => { const v = window.odx.view; v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.create(v.state.doc, 3))); window.notes.insertNote(v, 'footnote'); });
await a.waitForTimeout(900); // (the note opens for typing once drawn)
await a.evaluate(() => { const n = window.odx.view.odx.hfView; if (n) n.dispatch(n.state.tr.insertText('A shared note', n.state.doc.content.size - 1)); window.notes.closeNote(); });
check(await until(b, () => [...document.querySelectorAll('.odx-note')].some((e) => e.textContent.includes('A shared note'))), "A's footnote shows in B's pages");
// A header part A adds (the document had no header): B gets the part as well as its text, and saves it.
await a.evaluate((bytes) => window.odx.open(new Uint8Array(bytes), 'form.docx'), [...fs.readFileSync('test/fixtures/form.docx')]);
check(await until(b, () => window.odx.view.state.doc.textContent.includes('Customer name')), 'B follows A to a document with no header');
await a.waitForTimeout(800);
await a.evaluate(() => { const v = window.odx.view; window.hf.editStory(v, 'header'); const h = v.odx.hfView; h.dispatch(h.state.tr.insertText('NEW HEADER PART ', 1)); window.hf.closeStory(); });
check(await until(b, () => [...document.querySelectorAll('.odx-hf')].some((e) => e.textContent.includes('NEW HEADER PART'))), "the header part A added shows in B's pages");
const bFiles = unzipSync(new Uint8Array(Buffer.from(await b.evaluate(() => [...window.odx.save()]))));
check(Object.entries(bFiles).some(([n, d]) => /^word\/header/.test(n) && strFromU8(d).includes('NEW HEADER PART')), "B's saved file has the header part");
// A picture A inserts: B shows it, and B's saved file has it (its file, under the same relationship).
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAJUlEQVR4nGP8z4AKmBgYGBgYGJgZGBj+M2Ao+M+AoQBZDWYGAAA0HwIKC8bLQwAAAABJRU5ErkJggg==', 'base64');
fs.mkdirSync('test/out', { recursive: true });
fs.writeFileSync('test/out/pic.png', PNG);
await a.evaluate(() => { const v = window.odx.view; v.focus(); v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.create(v.state.doc, 3))); });
await a.setInputFiles('input[type=file][accept^="image/"]', 'test/out/pic.png');
check(await until(b, () => [...document.querySelectorAll('.odx-content img')].some((i) => i.src.startsWith('data:image/png') && i.complete && i.naturalWidth === 16)), "A's picture shows in B");
const bPics = unzipSync(new Uint8Array(Buffer.from(await b.evaluate(() => [...window.odx.save()]))));
const media = Object.keys(bPics).filter((n) => /^word\/media\//.test(n) && bPics[n].length === PNG.length);
const rels = strFromU8(bPics['word/_rels/document.xml.rels']);
check(media.length === 1 && rels.includes(media[0].replace('word/', '')), `B's saved file has the picture (${media.join(', ')})`);
// Styles and lists A adds: a new list (I., II., which form.docx has no numbering for) on the first paragraph, and a
// table of contents (its TOC Heading and TOC 1-3 styles, which form.docx lacks). B draws them as A does without
// reopening, and B's saved styles.xml and numbering.xml have them; so does an editor joining after.
await a.evaluate(() => {
  const v = window.odx.view;
  v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.create(v.state.doc, 1)));
  window.lists.setList(v.odx.session, window.lists.LIST_LIBRARY.numbering.find((d) => d.label === 'I.').levels)(v.state, v.dispatch);
});
const numId = await a.evaluate(() => window.odx.view.state.doc.firstChild.attrs.numId);
const label = (page) => page.evaluate(() => { const v = window.odx.view; return v.nodeDOM(0)?.getAttribute('data-label') ?? null; });
const labelA = await label(a);
check(labelA === 'I.', `A's new list numbers its paragraph (${labelA})`);
check(await until(b, (id) => window.odx.view.state.doc.firstChild.attrs.numId === id && window.odx.view.nodeDOM(0)?.getAttribute('data-label') === 'I.', numId), `B numbers the paragraph in A's new list (${await label(b)})`);
await a.evaluate(() => { const v = window.odx.view; v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.create(v.state.doc, 1))); window.toc.insertToc(v); });
// The TOC Heading paragraph as drawn: its style's font size, colour and weight.
const heading = (page) => page.evaluate(() => {
  const v = window.odx.view;
  let pos = -1;
  v.state.doc.forEach((n, p) => { if (pos < 0 && n.attrs.styleId === 'TOCHeading') pos = p; });
  const el = pos >= 0 && v.nodeDOM(pos);
  if (!el) return null;
  const cs = getComputedStyle(el);
  return [cs.fontSize, cs.color, cs.fontWeight, cs.marginTop].join(' ');
});
// (And the document's last paragraph, in its Normal style, to tell them apart.)
const normal = await a.evaluate(() => { const d = window.odx.view.state.doc; const cs = getComputedStyle(window.odx.view.nodeDOM(d.content.size - d.lastChild.nodeSize)); return [cs.fontSize, cs.color, cs.fontWeight, cs.marginTop].join(' '); });
const headA = await heading(a);
check(!!headA && headA !== normal, `A draws its TOC Heading in its style (${headA} | Normal ${normal})`);
check(await until(b, (h) => {
  const v = window.odx.view;
  let pos = -1;
  v.state.doc.forEach((n, p) => { if (pos < 0 && n.attrs.styleId === 'TOCHeading') pos = p; });
  const el = pos >= 0 && v.nodeDOM(pos);
  if (!el) return false;
  const cs = getComputedStyle(el);
  return [cs.fontSize, cs.color, cs.fontWeight, cs.marginTop].join(' ') === h;
}, headA), `B draws the TOC Heading as A does (${await heading(b)})`);
const defsSaved = (files) => {
  const styles = strFromU8(files['word/styles.xml']);
  const numbering = strFromU8(files['word/numbering.xml']);
  const num = new RegExp(`<w:num w:numId="${numId}"[^>]*>\\s*<w:abstractNumId w:val="(\\d+)"`).exec(numbering);
  return ['TOCHeading', 'TOC1', 'TOC2', 'TOC3'].every((id) => styles.includes(`w:styleId="${id}"`)) && !!num &&
    new RegExp(`<w:abstractNum [^>]*w:abstractNumId="${num[1]}"[\\s\\S]*?<w:numFmt w:val="upperRoman"`).test(numbering);
};
const bDefs = unzipSync(new Uint8Array(Buffer.from(await b.evaluate(() => [...window.odx.save()]))));
check(defsSaved(bDefs), "B's saved styles.xml and numbering.xml have A's TOC styles and new list");
const c = await open('C');
check(await until(c, (id) => { const v = window.odx.view; let pos = -1; v.state.doc.forEach((n, p) => { if (pos < 0 && n.attrs.numId === id) pos = p; }); return pos >= 0 && v.nodeDOM(pos)?.getAttribute('data-label') === 'I.'; }, numId), 'an editor joining after numbers the paragraph in the new list');
check(await heading(c) === headA, `... and draws the TOC Heading as A does (${await heading(c)})`);
check(defsSaved(unzipSync(new Uint8Array(Buffer.from(await c.evaluate(() => [...window.odx.save()]))))), "... and its saved file has them");
// Controls bound to custom XML (test/fixtures/mkbinding.mjs): A types into one; in B the other control bound to the
// same node shows the value once (not twice: B only follows the steps A's editor made), B's header shows it, and
// B's saved custom XML part has it; so does what an editor joining after saves.
await a.evaluate((bytes) => window.odx.open(new Uint8Array(bytes), 'binding.docx'), [...fs.readFileSync('test/fixtures/binding.docx')]);
check(await until(b, () => window.odx.view.state.doc.textContent.includes('Order no again: first order')), 'B follows A to the document with bound controls');
await a.waitForTimeout(800);
const orderRange = (page) => page.evaluate(() => {
  let r = null;
  window.odx.view.state.doc.descendants((n, pos) => { const m = n.isText && n.marks.find((k) => k.type.name === 'sdt' && k.attrs.tag === 'Order'); if (m) r = r ? { from: r.from, to: pos + n.nodeSize } : { from: pos, to: pos + n.nodeSize }; });
  return r;
});
const ord = await orderRange(a);
await a.evaluate(({ from, to }) => { const v = window.odx.view; v.focus(); v.dispatch(v.state.tr.insertText('shared order', from, to)); }, ord);
check(await until(b, () => window.odx.view.state.doc.textContent.includes('Order no again: shared order')), "A's typing in a bound control shows in B's other control bound to it");
await b.waitForTimeout(800);
const bText = await text(b);
check((bText.match(/shared order/g) || []).length === 2 && !bText.includes('shared ordershared order'), `... once each (${bText.slice(0, 120)})`);
check(await until(b, () => [...document.querySelectorAll('.odx-hf')].some((e) => e.textContent.includes('Header order: shared order'))), "... and in B's header");
const bBind = unzipSync(new Uint8Array(Buffer.from(await b.evaluate(() => [...window.odx.save()]))));
check(/<order>shared order<\/order>/.test(strFromU8(bBind['customXml/item1.xml'])), "B's saved custom XML part has A's value");
const d = await open('D');
check(await until(d, () => window.odx.view.state.doc.textContent.includes('Order no again: shared order')), 'an editor joining after shows the value');
const dBind = unzipSync(new Uint8Array(Buffer.from(await d.evaluate(() => [...window.odx.save()]))));
check(/<order>shared order<\/order>/.test(strFromU8(dBind['customXml/item1.xml'])) && (strFromU8(dBind['word/document.xml']).match(/shared order/g) || []).length === 2, "... and saves it, in the part and the controls");
check(!errors.length, `no page errors ${errors.join(' | ')}`);

await browser.close();
wss.close();
server.close();
console.log(failed ? `${failed} failures` : 'collab: all passed');
process.exit(failed ? 1 : 0);
