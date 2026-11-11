// Browser check of the typing comforts: AutoCorrect as you type (and Ctrl+Z giving back what was typed, and turning it
// off), Find with patterns (a regular expression, $1 in its replacement, an invalid one reported), and drafts (unsaved
// changes kept in the browser, offered back on opening the file again: Restore and Discard).
import { chromium } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve('demo');
const server = http.createServer((req, res) => {
  const p = path.join(root, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
  if (!p.startsWith(root) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'content-type': { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' }[path.extname(p)] || 'application/octet-stream' });
  fs.createReadStream(p).pipe(res);
}).listen(0);
const launch = process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : process.env.BROWSER_CHANNEL ? { channel: process.env.BROWSER_CHANNEL } : {};
const browser = await chromium.launch(launch);
const context = await browser.newContext({ viewport: { width: 1300, height: 950 } });
const page = await context.newPage();
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
let failed = 0;
const check = (ok, what) => { console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}`); if (!ok) failed++; };
const url = `http://localhost:${server.address().port}/`;
const open = async (f) => {
  await page.setInputFiles('input[type=file]', `test/fixtures/${f}`);
  await page.waitForFunction((n) => window.odx?.fileName === n, f);
  await page.waitForTimeout(800);
};
const firstPara = () => page.evaluate(() => window.odx.view.state.doc.firstChild.textContent);
await page.goto(url);
await page.waitForSelector('.odx-content p');

// AutoCorrect.
await open('word-template.docx');
await page.click('.odx-content p >> nth=0');
await page.keyboard.press('Control+Home');
await page.keyboard.press('Enter');
await page.keyboard.press('ArrowUp');
await page.keyboard.type('he said "it\'s fine" -- well(c)... teh end. so ');
check(await firstPara() === 'He said “it’s fine” – well©… the end. So ', `AutoCorrect as you type: ${await firstPara()}`);
await page.keyboard.type('x--');
await page.keyboard.type('y');
check((await firstPara()).endsWith('x—y'), 'a double hyphen between words becomes an em dash when the next word starts');
await page.keyboard.press('Control+z');
check((await firstPara()).endsWith('x--y'), 'Ctrl+Z gives back what was typed');
await page.evaluate(() => window.odx.setFeatures({ autocorrect: false }));
await page.keyboard.type(' "plain" teh');
check((await firstPara()).endsWith(' "plain" teh'), 'turned off, nothing is corrected');
await page.evaluate(() => window.odx.setFeatures({ autocorrect: true }));

// Find with patterns.
await open('word-template.docx');
await page.click('.odx-content p >> nth=0');
await page.keyboard.press('Control+Home');
await page.keyboard.type('Invoice 2024-03 and 2025-11 ');
await page.keyboard.press('Control+h');
await page.click('.odx-findbar label:has-text("Patterns") input');
await page.fill('.odx-findbar input[aria-label="Find"]', '(\\d{4})-(\\d{2})');
await page.keyboard.press('Enter');
check(await page.textContent('.odx-findinfo') === '1 of 2', `a pattern finds its matches (${await page.textContent('.odx-findinfo')})`);
await page.fill('.odx-findbar input[aria-label="Replace with"]', '$2/$1');
await page.click('.odx-findbar button:has-text("Replace all")');
check((await firstPara()).startsWith('Invoice 03/2024 and 11/2025'), `Replace all fills in $1 and $2: ${(await firstPara()).slice(0, 30)}`);
await page.fill('.odx-findbar input[aria-label="Find"]', '(unclosed');
await page.keyboard.press('Enter');
check(/^Not a valid pattern/.test(await page.textContent('.odx-findinfo')), 'an invalid pattern is reported');
await page.keyboard.press('Escape');

// Drafts: edit, leave the page, open the same file again.
await open('rich.docx');
await page.click('.odx-content p >> nth=0');
await page.keyboard.press('End');
await page.keyboard.type(' DRAFTKEPT');
for (let t = 0; t < 40 && !(await page.evaluate(() => window.odx.draft().then(Boolean))); t++) await page.waitForTimeout(250); // (kept 3 seconds after the last edit; waitForFunction doesn't wait on a promise)
await page.reload();
await page.waitForSelector('.odx-content p');
await open('rich.docx');
check(await page.waitForSelector('.odx-draftnote', { timeout: 8000 }).then(() => true, () => false), 'the drafts notice appears');
check(await page.isVisible('.odx-draftnote') && /rich\.docx/.test(await page.textContent('.odx-draftnote')), 'opening the file again offers its unsaved changes');
await page.click('.odx-draftnote button:has-text("Restore")');
await page.waitForTimeout(800);
check((await page.evaluate(() => window.odx.view.state.doc.textContent)).includes('DRAFTKEPT') && await page.evaluate(() => window.odx.dirty), 'Restore brings them back, still unsaved');
check(!(await page.isVisible('.odx-draftnote')), 'and the notice goes');
await page.evaluate(() => window.odx.markSaved());
await page.waitForTimeout(300);
check(await page.evaluate(() => window.odx.draft()) === null, 'saving clears the draft');
// Discard.
await page.click('.odx-content p >> nth=0');
await page.keyboard.type(' AGAIN');
for (let t = 0; t < 40 && !(await page.evaluate(() => window.odx.draft().then(Boolean))); t++) await page.waitForTimeout(250); // (kept 3 seconds after the last edit; waitForFunction doesn't wait on a promise)
check((await page.evaluate(() => window.odx.draft()))?.name === 'rich.docx', 'a new edit is kept again');
await page.reload(); // (the same file chosen again in the same input doesn't count as a change)
await page.waitForSelector('.odx-content p');
await open('rich.docx');
check(await page.waitForSelector('.odx-draftnote', { timeout: 8000 }).then(() => true, () => false), 'the drafts notice appears');
await page.click('.odx-draftnote button:has-text("Discard")');
await page.waitForTimeout(300);
check(await page.evaluate(() => window.odx.draft()) === null && !(await page.evaluate(() => window.odx.view.state.doc.textContent)).includes('AGAIN'), 'Discard forgets them');

check(!errors.length, `no page errors ${errors.join(' | ')}`);
await browser.close();
server.close();
console.log(failed ? `comforts: ${failed} failed` : 'comforts: all passed');
process.exit(failed ? 1 : 0);
