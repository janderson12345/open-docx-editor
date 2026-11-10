// Browser check of WOPI (src/wopi.js) against a mock host on another origin (test/wopi-host.mjs): a writable file
// opens locked, saves on its own after an edit and with Ctrl+S under its lock, keeps the lock fresh and releases it on
// closing and when the page goes; a read-only user and a file someone else has locked open read-only; a refused token
// and a lost lock are reported, and nothing is overwritten.
import { chromium } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { startWopiHost } from './wopi-host.mjs';

const root = path.resolve('demo');
const server = http.createServer((req, res) => {
  const p = path.join(root, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
  if (!p.startsWith(root) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'content-type': { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' }[path.extname(p)] || 'application/octet-stream' });
  fs.createReadStream(p).pipe(res);
}).listen(0);
const sample = fs.readFileSync('test/fixtures/rich.docx');
const host = await startWopiHost({
  a: { name: 'Report.docx', bytes: sample, token: 'tok-a', user: 'Ann Editor' },
  b: { name: 'Viewer.docx', bytes: sample, token: 'tok-b', canWrite: false },
  c: { name: 'Taken.docx', bytes: sample, token: 'tok-c' },
  d: { name: 'Lost.docx', bytes: sample, token: 'tok-d' },
  e: { name: 'Keys.docx', bytes: sample, token: 'tok-e' },
});
host.setLock('c', 'someone-elses-lock');
const launch = process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : process.env.BROWSER_CHANNEL ? { channel: process.env.BROWSER_CHANNEL } : {};
const browser = await chromium.launch(launch);
let failed = 0;
const check = (ok, what) => { console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}`); if (!ok) failed++; };
const text = (id) => strFromU8(unzipSync(new Uint8Array(host.files[id].bytes))['word/document.xml']);
const ops = (id) => host.log.filter((l) => l.id === id).map((l) => l.op);
const errors = [];
const open = async (id, token, extra = '') => {
  const page = await browser.newPage({ viewport: { width: 1200, height: 900 } });
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto(`http://localhost:${server.address().port}/wopi.html#src=${encodeURIComponent(host.url(id))}&access_token=${token}${extra}`);
  await page.waitForFunction(() => window.wopiApp || window.wopiError, null, { timeout: 15000 });
  await page.waitForTimeout(500);
  return page;
};
const status = (page) => page.textContent('#status');
const typeAtStart = async (page, s) => { await page.click('.odx-content p >> nth=0'); await page.keyboard.press('Home'); await page.keyboard.type(s); };

// 1. A writable file: open, locked, edited, saved on its own, the lock refreshed, then closed and unlocked.
let page = await open('a', 'tok-a', '&autosave=500&refresh=1000');
check((await page.evaluate(() => window.wopiApp.editor.fileName)) === 'Report.docx' && !(await page.evaluate(() => window.wopiApp.editor.readOnly)), 'a writable file opens for editing under its name');
check(/^odx-[0-9a-f]{32}$/.test(host.lockOf('a') || ''), 'it is locked while open');
check(host.log.every((l) => !l.id || l.token), 'every call carried the access token');
await typeAtStart(page, 'WOPISAVED ');
await page.waitForTimeout(1500);
check(text('a').includes('WOPISAVED') && host.files.a.version >= 2, 'an edit is saved to the host on its own');
check(host.log.some((l) => l.id === 'a' && l.op === 'PUT' && l.lock === host.lockOf('a')), 'saved under its lock');
check(ops('a').includes('REFRESH_LOCK'), 'the lock is refreshed while open');
check(/Saved/.test(await status(page)), 'the page says it saved');
await page.evaluate(() => window.wopiApp.close());
check(host.lockOf('a') === null && ops('a').includes('UNLOCK'), 'closing releases the lock');
await page.close();

// 2. Ctrl+S saves at once (no autosave); leaving the page releases the lock.
page = await open('e', 'tok-e', '&autosave=0');
// One file: File > New and Open aren't offered (nothing else could be saved over the host's file).
await page.click('.odx-rtab:text-is("File")');
const items = await page.$$eval('.odx-fileitem', (els) => els.map((e) => e.textContent));
check(items.length > 0 && !items.some((t) => /^\W*(New|Open)$/.test(t.trim())), `File menu has no New or Open (${items.map((t) => t.trim()).join(', ')})`);
await page.keyboard.press('Escape');
await typeAtStart(page, 'KEYSAVE ');
check(!text('e').includes('KEYSAVE'), 'nothing saved before Ctrl+S with autosave off');
await page.keyboard.press('Control+s');
await page.waitForTimeout(800);
check(text('e').includes('KEYSAVE'), 'Ctrl+S saves to the host');
await page.goto('about:blank');
await page.waitForTimeout(800);
check(host.lockOf('e') === null, 'leaving the page releases the lock');
await page.close();

// 3. A user who may only read: read-only, no lock taken.
page = await open('b', 'tok-b');
check(await page.evaluate(() => window.wopiApp.editor.readOnly), 'a read-only user gets a read-only editor');
check(!ops('b').includes('LOCK'), 'and no lock is taken');
await page.close();

// 4. A file someone else has locked: read-only, told why, their lock untouched.
page = await open('c', 'tok-c');
check(await page.evaluate(() => window.wopiApp.editor.readOnly), 'a file locked by someone else opens read-only');
check(/someone else/.test(await status(page)), 'the page says someone else is editing it');
check(host.lockOf('c') === 'someone-elses-lock' && !ops('c').includes('PUT'), 'their lock stays and nothing is written');
await page.close();

// 5. A refused token.
page = await open('a', 'wrong-token');
check(/access token was refused/.test(await status(page)), 'a refused token is reported');
await page.close();

// 6. The lock lost while editing (the host gave it to someone else): the save is refused and reported; nothing overwritten.
page = await open('d', 'tok-d', '&autosave=500');
host.dropLock('d');
host.setLock('d', 'taken-over');
const before = host.files.d.version || 1;
await typeAtStart(page, 'SHOULDNOTSAVE ');
await page.waitForTimeout(1200);
check((host.files.d.version || 1) === before && !text('d').includes('SHOULDNOTSAVE'), 'a save under a lost lock is refused, the file untouched');
check(/locked by someone else/.test(await status(page)), 'and the page says why');
check(await page.evaluate(() => window.wopiApp.editor.dirty), 'the edit stays unsaved (dirty)');
await page.close();

check(!errors.length, `no page errors ${errors.join(' | ')}`);
await browser.close();
await host.close();
server.close();
console.log(failed ? `wopi: ${failed} failed` : 'wopi: all passed');
process.exit(failed ? 1 : 0);
