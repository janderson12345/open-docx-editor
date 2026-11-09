// Forms mode end to end in the browser (run with the e2e tests, after `npm run build`), with the keyboard and mouse as
// a person filling the form would: Tab into the first field and type; typing outside a field changes nothing; Space
// ticks the check box; Enter opens the drop-down's list and the date's calendar; Shift+Tab goes back; the rich-text
// field takes text. Then the values read back through the API, survive saving (read again from the saved bytes), and
// the saved file is sound (tools/check-ooxml.mjs). The saved file is left in test/out/form.filled.docx.
import { chromium } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import './setup.js';
import { readFieldValues } from '../src/fieldsNode.js';

const root = path.resolve('demo');
const server = http.createServer((req, res) => {
  const p = path.join(root, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
  if (!p.startsWith(root) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'content-type': { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' }[path.extname(p)] || 'application/octet-stream' });
  fs.createReadStream(p).pipe(res);
}).listen(0);
const launch = process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : process.env.BROWSER_CHANNEL ? { channel: process.env.BROWSER_CHANNEL } : {};
const browser = await chromium.launch(launch);
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } });
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
let failed = 0;
const check = (ok, msg) => { if (!ok) { failed++; console.log('FAIL', msg); } else console.log('ok  ', msg); };
const values = () => page.evaluate(() => window.odx.getFieldValues());
const inField = (tag) => page.evaluate((t) => { const v = window.odx.view; const at = v.state.selection.from; let hit = null; v.state.doc.nodesBetween(Math.max(0, at - 1), Math.min(v.state.doc.content.size, at + 1), (n) => { const m = n.marks?.find((k) => k.type.name === 'sdt'); if (m) hit = m.attrs.tag; }); return hit === t; }, tag);
const outside = () => page.evaluate(() => { const d = window.odx.view.state.doc; let t = ''; d.descendants((n) => { if (n.isText && !n.marks.some((m) => m.type.name === 'sdt')) t += n.text; }); return t; });

await page.goto(`http://localhost:${server.address().port}/`);
await page.waitForSelector('.odx-content p');
await page.setInputFiles('input[type=file]', 'test/fixtures/form.docx');
await page.waitForFunction(() => window.odx?.fileName === 'form.docx');
await page.waitForTimeout(800);
await page.evaluate(() => window.odx.setForms(true));
check(await page.evaluate(() => getComputedStyle(document.querySelector('.odx-toolbar')).display === 'none'), 'forms mode hides the ribbon');
const before = await outside();

// Click the first field's placeholder (Customer name): it's selected, as in Word, and typing replaces it.
await page.click('.odx-content .odx-field');
await page.keyboard.type('Ada Lovelace');
check((await values()).CustomerName === 'Ada Lovelace', 'typed into the text field');
// The caret stays in the fields, as in a protected Word form: Ctrl+Home goes to the first field; a click on the form's
// own text puts it in the field after, where typing goes (and Backspace takes it back out). The form's text stays.
await page.keyboard.press('Control+Home');
check(await inField('CustomerName'), 'Ctrl+Home goes to the first field');
await page.click('.odx-content p', { position: { x: 5, y: 5 } });
await page.keyboard.type('XYZ');
check((await outside()) === before, "typing after a click on the form's own text leaves that text alone");
check((await values()).CustomerName === 'Ada LovelaceXYZ', `... and goes into the field after (${(await values()).CustomerName})`);
await page.keyboard.press('Backspace');
await page.keyboard.press('Backspace');
await page.keyboard.press('Backspace');
check((await values()).CustomerName === 'Ada Lovelace', 'Backspace edits the field');
// Tab to the check box and tick it with Space.
await page.keyboard.press('Tab');
await page.keyboard.press('Space');
check((await values()).Agree === true, 'Space ticks the check box');
// Tab to the drop-down; Enter opens its list; pick the second item.
await page.keyboard.press('Tab');
await page.keyboard.press('Enter');
const pick = await page.waitForSelector('select.odx-fieldpick', { timeout: 3000 }).catch(() => null);
check(!!pick, 'Enter opens the drop-down list');
check(!!pick && await pick.evaluate((e) => getComputedStyle(e).position === 'fixed' && Number(getComputedStyle(e).zIndex) >= 30 && !!e.style.left), 'the list opens at the field, over the page');
let chosen = null;
if (pick) {
  chosen = await pick.evaluate((s) => s.options[1].value);
  await pick.selectOption({ index: 1 }); // (choosing closes the list)
}
check(chosen != null && (await values()).Color === chosen, `chose "${chosen}" from the list`);
// Tab to the date; Enter opens the calendar; set 25 December 2026.
await page.keyboard.press('Tab');
await page.keyboard.press('Enter');
const cal = await page.waitForSelector('input.odx-fieldpick[type=date]', { timeout: 3000 }).catch(() => null);
check(!!cal, 'Enter opens the date picker');
if (cal) await cal.evaluate((i) => { i.value = '2026-12-25'; i.dispatchEvent(new Event('change')); });
const due = (await values()).Due;
check(/12\/25\/2026|25|December/.test(String(due)), `date set (${due})`);
// Shift+Tab goes back to the drop-down; Tab twice more reaches the rich-text field.
await page.keyboard.press('Shift+Tab');
check(await inField('Color'), 'Shift+Tab moves back to the drop-down');
await page.keyboard.press('Tab');
await page.keyboard.press('Tab');
await page.keyboard.type('Line one of the notes');
check(String((await values()).Notes).includes('Line one of the notes'), 'typed into the rich-text field');

// Saved, the values read back from the file.
const saved = Buffer.from(await page.evaluate(() => [...window.odx.save()]));
fs.mkdirSync('test/out', { recursive: true });
fs.writeFileSync('test/out/form.filled.docx', saved);
const back = readFieldValues(new Uint8Array(saved));
const live = await values();
check(JSON.stringify(back) === JSON.stringify(live), `saved file has the same values ${JSON.stringify(back)}`);
try { execFileSync(process.execPath, ['tools/check-ooxml.mjs', 'test/out/form.filled.docx'], { stdio: 'pipe' }); check(true, 'saved file is sound'); }
catch (e) { check(false, `saved file: ${e.stdout}`); }
check(!errors.length, `no page errors ${errors.join(' | ')}`);

await browser.close();
server.close();
console.log(failed ? `${failed} failures` : 'forms: all passed');
process.exit(failed ? 1 : 0);
