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
const launch = process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : process.env.BROWSER_CHANNEL ? { channel: process.env.BROWSER_CHANNEL } : {};
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
check(!errors.length, `no page errors ${errors.join(' | ')}`);

await browser.close();
wss.close();
server.close();
console.log(failed ? `${failed} failures` : 'collab: all passed');
process.exit(failed ? 1 : 0);
