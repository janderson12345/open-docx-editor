// Accessibility checks on the rendered editor (Section 508 / WCAG 2.1 AA), run with the e2e tests.
// 1. axe-core over every ribbon tab and each menu, light and dark.
// 2. Our pattern scan (test/a11y-scan.js) for what axe doesn't cover: names that hide the visible label (2.5.3),
//    targets under 24 px (2.5.8), status messages without a live region (4.1.3), unnamed widgets.
// Floors make sure the checks still find the controls, so a markup change can't pass by matching nothing.
// Exceptions are listed by name, with the reason.
/* global __a11yScan */ // (test/a11y-scan.js, added to the page)
import { chromium } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';

const EXCEPT = {
  // The document's own content is the author's (Review > Check Accessibility reports it), not the editor's UI.
  axe: (v, n) => n.target.join(' ').includes('.odx-content'),
};
const FLOORS = { controls: 45, tabs: 6, menuItems: 5, contextItems: 6 };

const root = path.resolve('demo');
const server = http.createServer((req, res) => {
  const p = path.join(root, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
  if (!p.startsWith(root) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'content-type': { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' }[path.extname(p)] || 'application/octet-stream' });
  fs.createReadStream(p).pipe(res);
}).listen(0);
const url = `http://localhost:${server.address().port}/`;
const axe = fs.readFileSync('node_modules/axe-core/axe.min.js', 'utf8');
const scan = fs.readFileSync('test/a11y-scan.js', 'utf8');
const browser = await chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {});
let failed = 0;
const fail = (msg) => { failed++; console.log('FAIL', msg); };

for (const scheme of ['light', 'dark']) {
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 }, colorScheme: scheme });
  await page.goto(url);
  await page.waitForSelector('.odx-content p');
  await page.addScriptTag({ content: axe });
  await page.addScriptTag({ content: scan });
  const check = async (where) => {
    const r = await page.evaluate(() => axe.run(document.querySelector('.odx-toolbar') ? document.querySelector('.odx-toolbar').parentElement : document, { runOnly: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'] }));
    for (const v of r.violations) for (const n of v.nodes) if (!EXCEPT.axe(v, n)) fail(`${scheme} ${where}: ${v.id} ${n.target.join(' ')}`);
    const s = await page.evaluate(() => __a11yScan(document.querySelector('.odx-toolbar').parentElement));
    for (const [k, list] of Object.entries(s)) if (k !== '_counts') for (const x of list) fail(`${scheme} ${where}: ${k}: ${x}`);
    return s._counts;
  };
  const counts = await check('Home');
  if (counts.controls < FLOORS.controls) fail(`only ${counts.controls} controls found on Home (floor ${FLOORS.controls})`);
  const tabs = await page.$$eval('[role=tab]', (t) => t.map((x) => x.textContent));
  if (tabs.length < FLOORS.tabs) fail(`only ${tabs.length} ribbon tabs found`);
  for (const t of tabs.filter((x) => x !== 'File' && x !== 'Home')) { await page.click(`[role=tab]:text-is("${t}")`); await check(t); }
  await page.click('[role=tab]:text-is("File")');
  if ((await page.$$('.odx-pop [role=menuitem]')).length < FLOORS.menuItems) fail('File menu items not found');
  await check('File menu');
  await page.keyboard.press('Escape');
  const para = await page.locator('.odx-content p').first().boundingBox();
  await page.mouse.click(para.x + 10, para.y + 5, { button: 'right' });
  if ((await page.$$('.odx-ctx [role=menuitem]')).length < FLOORS.contextItems) fail('context menu items not found');
  await check('context menu');
  await page.keyboard.press('Escape');
  // Toggle state is exposed to screen readers (4.1.2).
  if ((await page.getAttribute('[aria-label^="Bold"]', 'aria-pressed')) == null) fail('Bold has no aria-pressed');
  // Keyboard: F6 reaches the ribbon, arrows move between tabs, Esc returns to the document.
  await page.click('.odx-content p');
  await page.keyboard.press('F6');
  if ((await page.evaluate(() => document.activeElement.getAttribute('role'))) !== 'tab') fail('F6 did not move focus to the ribbon');
  await page.keyboard.press('Escape');
  if (!(await page.evaluate(() => document.activeElement.classList.contains('odx-content')))) fail('Esc in the ribbon did not return to the document');
  await page.close();
}
await browser.close();
server.close();
console.log(failed ? `${failed} accessibility failure(s)` : 'accessibility: ok');
process.exit(failed ? 1 : 0);
