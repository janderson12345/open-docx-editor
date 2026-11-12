// Browser check of AI Writing (Home > AI Writing, with a stand-in model the page gives as window.odxAi): a rewrite of
// the selection goes in as a tracked change under the AI's name; a summary is shown, then inserted below as a tracked
// insertion; a question's answer replaces the selection; what the model is asked; a model's error is said.
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
// The stand-in model: what it's asked is kept; its answers are canned.
await page.addInitScript(() => {
  window.aiAsked = [];
  window.odxAi = async (req) => {
    window.aiAsked.push({ action: req.action, instruction: req.instruction, prompt: req.prompt, text: req.text, hasDoc: !!req.document, signal: req.signal instanceof AbortSignal });
    await new Promise((r) => setTimeout(r, 50));
    if (req.action === 'shorter') return req.text.replace('quarterly results were', 'results were');
    if (req.action === 'summarize') return 'A short summary.';
    if (req.action === 'ask') { if (/fail/.test(req.prompt)) throw new Error('quota reached'); return { text: 'The answer.' }; }
    return req.text;
  };
});
await page.goto(`http://localhost:${server.address().port}/`);
await page.waitForSelector('.odx-content p');
await page.evaluate(() => window.odx.newDocument({ share: false }));
await page.click('.odx-content p');
await page.keyboard.type('The quarterly results were better than expected.');
await page.keyboard.press('Enter');
await page.keyboard.type('Second paragraph.');
const openAi = async () => { await page.click('.odx-rtab:text-is("Home")'); await page.click('button[title^="AI writing"]'); await page.waitForTimeout(150); };
const click = (label) => page.evaluate((l) => [...document.querySelectorAll('.odx-pop .odx-item')].find((b) => b.textContent.trim() === l)?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })), label);

// A rewrite of the first paragraph (the cursor in it, nothing selected): tracked under "AI", only the edit.
await page.click('.odx-content p >> nth=0');
await openAi();
await click('Make it shorter');
await page.waitForTimeout(400);
const saved1 = strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/document.xml']);
const asked1 = await page.evaluate(() => window.aiAsked[0]);
check(/<w:del [^>]*w:author="AI"[^>]*>[\s\S]*?quarterly [\s\S]*?<\/w:del>/.test(saved1) && !/<w:ins /.test(saved1) && /better than expected/.test(saved1)
  && asked1.text === 'The quarterly results were better than expected.' && asked1.instruction === 'Make it shorter' && asked1.hasDoc && asked1.signal,
  `a rewrite goes in as a tracked change under the AI's name (${JSON.stringify(asked1)})`);

// A summary of the document (nothing selected in the second paragraph? the cursor's paragraph: summarized), shown, then
// inserted below as a tracked insertion.
await page.evaluate(() => window.odx.resolveAll?.(true));
await page.click('.odx-content p >> nth=1');
await openAi();
await click('Summarize the selection');
await page.waitForTimeout(400);
const shown = await page.evaluate(() => document.querySelector('.odx-pop .odx-ai-out')?.textContent);
await click('Insert below');
await page.waitForTimeout(300);
const paras = await page.evaluate(() => { const out = []; window.odx.view.state.doc.forEach((n) => out.push(n.textContent)); return out; });
const saved2 = strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/document.xml']);
check(shown === 'A short summary.' && paras.includes('A short summary.') && /<w:ins [^>]*w:author="AI"[^>]*>[\s\S]*?A short summary\./.test(saved2),
  `a summary is shown, then inserted below as a tracked insertion (${JSON.stringify({ shown, paras })})`);

// A question about a selection: its answer replaces the selection (tracked).
await page.evaluate(() => {
  const v = window.odx.view; let a = null, b = null;
  v.state.doc.forEach((n, pos) => { if (n.textContent === 'Second paragraph.') { a = pos + 1; b = pos + 1 + 6; } });
  v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.create(v.state.doc, a, b)));
});
await openAi();
await page.fill('.odx-pop .odx-ai-q', 'What should this say?');
await page.press('.odx-pop .odx-ai-q', 'Enter');
await page.waitForTimeout(400);
if (process.env.SHOT) await page.screenshot({ path: process.env.SHOT, clip: { x: 0, y: 0, width: 1300, height: 520 } });
await click('Replace selection');
await page.waitForTimeout(300);
const asked3 = await page.evaluate(() => window.aiAsked.at(-1));
const saved3 = strFromU8(unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(await window.odx.save()))))['word/document.xml']);
check(asked3.action === 'ask' && asked3.prompt === 'What should this say?' && asked3.text === 'Second' && /<w:ins [^>]*w:author="AI"[^>]*>[\s\S]*?The answer\./.test(saved3),
  `a question's answer replaces the selection, tracked (${JSON.stringify(asked3)})`);

// Esc in the panel (from its Ask button, after an answer) closes it, back on AI Writing, collapsed.
await openAi();
await page.fill('.odx-pop .odx-ai-q', 'Anything?');
await page.press('.odx-pop .odx-ai-q', 'Enter');
await page.waitForTimeout(400);
await page.focus('.odx-pop .odx-ai-ask .odx-item');
await page.keyboard.press('Escape');
await page.waitForTimeout(200);
const esc = await page.evaluate(() => ({ hidden: document.querySelector('.odx-pop').hidden, expanded: document.querySelector('button[title^="AI writing"]').getAttribute('aria-expanded'), focus: document.activeElement?.title?.slice(0, 10) }));
check(esc.hidden && esc.expanded === 'false' && esc.focus === 'AI writing', `Esc closes the AI panel, back on its button (${JSON.stringify(esc)})`);

// A model's error is said, and nothing changes.
const before = await page.evaluate(() => window.odx.view.state.doc.textContent);
await openAi();
await page.fill('.odx-pop .odx-ai-q', 'please fail');
await page.press('.odx-pop .odx-ai-q', 'Enter');
await page.waitForTimeout(400);
const errText = await page.evaluate(() => document.querySelector('.odx-pop [role=status]')?.textContent);
const after = await page.evaluate(() => window.odx.view.state.doc.textContent);
check(/quota reached/.test(errText || '') && before === after, `a model's error is said (${errText})`);
await page.keyboard.press('Escape');

check(!errors.length, `no page errors ${errors.join(' | ')}`);
await browser.close();
server.close();
console.log(failed ? `aiwrite: ${failed} failed` : 'aiwrite: all passed');
process.exit(failed ? 1 : 0);
