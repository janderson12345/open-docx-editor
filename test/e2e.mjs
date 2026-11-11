// Browser check: open each fixture in the demo, edit, save, verify with mammoth.
// Run after `npm run build`: node test/e2e.mjs [outDir]
import { chromium, firefox, webkit } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import mammoth from 'mammoth';
import { unzipSync, strFromU8 } from 'fflate';

const root = path.resolve('demo');
const outDir = process.argv[2] || 'test/out';
fs.mkdirSync(outDir, { recursive: true });
const server = http.createServer((req, res) => {
  const p = path.join(root, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
  if (req.url === '/favicon.ico') { res.writeHead(204); return res.end(); }
  if (!p.startsWith(root) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'content-type': { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' }[path.extname(p)] || 'application/octet-stream' });
  fs.createReadStream(p).pipe(res);
}).listen(0);
const url = `http://localhost:${server.address().port}/`;

// BROWSER=firefox or webkit runs it in those engines instead (Playwright's own builds).
const engine = { firefox, webkit }[process.env.BROWSER];
const browser = engine ? await engine.launch() : await chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : process.env.BROWSER_CHANNEL ? { channel: process.env.BROWSER_CHANNEL } : {});
const page = await browser.newPage({ viewport: { width: 1200, height: 1000 } });
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
// (A file that fails to load is reported by its address, from the response; the console's message doesn't name it.
// The demo looks for a sample document beside the page, which only the published demo has.)
page.on('console', (m) => m.type() === 'error' && !/^Failed to load resource/.test(m.text()) && errors.push(m.text()));
page.on('response', (r) => r.status() >= 400 && !r.url().endsWith('/sample.docx') && errors.push(`${r.status()} ${r.url()}`));
await page.goto(url);
let failed = 0;
const fixtures = fs.readdirSync('test/fixtures').filter((f) => f.endsWith('.docx'));
for (const f of fixtures) {
  await page.setInputFiles('input[type=file]', `test/fixtures/${f}`);
  await page.waitForFunction((n) => window.odx.fileName === n, f);
  // (Not in WebKit: Playwright's screenshot waits for the page's fonts, which WebKit can leave loading for good.)
  if (process.env.BROWSER !== 'webkit') await page.screenshot({ path: `${outDir}/${f}.png`, fullPage: true });
  // Type at the end of the first paragraph, then add a bold word.
  await page.click('.odx-content p >> nth=0');
  await page.keyboard.press('End');
  await page.keyboard.type(' typed');
  await page.keyboard.press('Control+b');
  await page.keyboard.type(' bolded');
  await page.keyboard.press('Control+b');
  const bytes = await page.evaluate(() => Array.from(window.odx.save()));
  fs.writeFileSync(`${outDir}/${f.replace('.docx', '.edited.docx')}`, Buffer.from(bytes));
  const { value: html } = await mammoth.convertToHtml({ buffer: Buffer.from(bytes) });
  // (After a full stop AutoCorrect capitalises the word, as Word does.) A protected document refuses what it protects
  // against: the typing when it can only be read, the bold when formatting is limited.
  const prot = await page.evaluate(() => window.odx.protection);
  const locked = prot.readOnlyNow || (prot.enforced && (prot.edit === 'readOnly' || prot.edit === 'comments'));
  const ok = locked ? !/typed/.test(html)
    : prot.enforced && prot.formatting ? / [tT]yped bolded/.test(html) && !/<strong>/.test(html)
      : / [tT]yped<strong> bolded<\/strong>/.test(html.replace(/<\/strong><strong>/g, ''));
  console.log(`${ok ? 'ok  ' : 'FAIL'} ${f}`);
  if (!ok) { failed++; console.log(html.slice(0, 400)); }
}
// Blank document: type a heading, a list and a table.
await page.click('.odx-rtab:text-is("File")');
await page.click('.odx-fileitem:has-text("New")');
await page.click('.odx-rtab:text-is("Home")');
await page.click('.odx-content p');
await page.click('.odx-stile[title="Heading 1"]');
await page.keyboard.type('Made in the browser');
await page.keyboard.press('Enter');
await page.keyboard.type('Normal text after the heading.');
await page.keyboard.press('Enter');
await page.click('button[title="Bullets"]');
await page.keyboard.type('first');
await page.keyboard.press('Enter');
await page.keyboard.press('Tab');
await page.keyboard.type('nested');
await page.keyboard.press('Enter');
await page.keyboard.press('Enter');
await page.keyboard.press('Enter');
await page.click('.odx-rtab:text-is("Insert")');
await page.click('button[title="Insert table"]');
await page.click('.odx-pop .odx-item >> nth=2');
await page.keyboard.type('cell');
if (process.env.BROWSER !== 'webkit') await page.screenshot({ path: `${outDir}/new-document.png` }); // (as above: WebKit can leave fonts loading)
const bytes = await page.evaluate(() => Array.from(window.odx.save()));
fs.writeFileSync(`${outDir}/new-document.docx`, Buffer.from(bytes));
const { value: html } = await mammoth.convertToHtml({ buffer: Buffer.from(bytes) });
const ok = /<h1>Made in the browser<\/h1><p>Normal text after the heading.<\/p><ul><li>first<ul><li>nested<\/li><\/ul><\/li><\/ul>/.test(html) && /<table><tr><td><p>cell<\/p><\/td>/.test(html);
console.log(`${ok ? 'ok  ' : 'FAIL'} new document`);
if (!ok) { failed++; console.log(html); }

// The ribbon's other pieces on the same document: page setup (Landscape, Two columns), Find, the context menu.
const docXml = async () => {
  const b = Buffer.from(await page.evaluate(() => Array.from(window.odx.save())));
  return strFromU8(unzipSync(new Uint8Array(b))['word/document.xml']);
};
await page.click('.odx-rtab:text-is("Layout")');
await page.click('button[title="Orientation"]');
await page.click('.odx-pop .odx-item:has-text("Landscape")');
await page.waitForTimeout(1500);
await page.click('.odx-rtab:text-is("Layout")');
await page.click('button[title="Columns"]');
await page.click('.odx-pop .odx-item:has-text("Two")');
await page.waitForTimeout(1500);
const sheet = await page.evaluate(() => { const r = document.querySelector('.odx-sheet').getBoundingClientRect(); return { w: r.width, h: r.height }; });
const xml = await docXml();
const setupOk = sheet.w > sheet.h && xml.includes('w:orient="landscape"') && xml.includes('w:num="2"');
console.log(`${setupOk ? 'ok  ' : 'FAIL'} page setup: landscape, two columns`);
if (!setupOk) { failed++; console.log(sheet, xml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)?.at(-1)); }
await page.click('.odx-content p >> nth=0');
await page.keyboard.press('Control+f');
await page.fill('.odx-findbox[type=search]', 'cell');
await page.press('.odx-findbox[type=search]', 'Enter');
const found = await page.textContent('.odx-findinfo');
const findOk = found === '1 result' && await page.isVisible('.odx-findbar');
console.log(`${findOk ? 'ok  ' : 'FAIL'} find: ${found}`);
if (!findOk) failed++;
await page.press('.odx-findbox[type=search]', 'Escape');
await page.click('.odx-content p >> nth=0', { button: 'right' });
const ctxOk = await page.isVisible('.odx-ctxmenu') && await page.isVisible('.odx-mini');
await page.keyboard.press('Escape');
const ctxClosed = await page.isHidden('.odx-ctx');
console.log(`${ctxOk && ctxClosed ? 'ok  ' : 'FAIL'} context menu opens on right-click, Esc closes it`);
if (!(ctxOk && ctxClosed)) failed++;
if (errors.length) { failed++; console.log('page errors:', errors); }
await browser.close();
server.close();
process.exit(failed ? 1 : 0);
