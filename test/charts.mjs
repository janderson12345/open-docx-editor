// Browser check of chart editing: double-click the chart, change its title, type and data in the dialog, apply,
// undo and redo, save, and read the saved chart part and workbook. Also: Enter on a selected chart opens the dialog,
// Esc closes it, and the dialog keeps the keyboard (Tab wraps).
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
await page.setInputFiles('input[type=file]', 'test/fixtures/chart.docx');
await page.waitForFunction(() => window.odx?.fileName === 'chart.docx');
await page.waitForTimeout(800);
const img = page.locator('.odx-content img[src^="data:image/svg"]').first();
const src = async () => decodeURIComponent((await img.getAttribute('src')) || '');
check((await src()).includes('Sales'), 'the chart is drawn with its title');

// (Near its right edge: the click's own position is after the chart, the dialog must still open.)
const ib = await img.boundingBox();
await page.mouse.dblclick(ib.x + ib.width - 6, ib.y + ib.height / 2);
await page.waitForSelector('.odx-chartdlg', { timeout: 3000 }).catch(() => {});
check(!!(await page.$('.odx-chartdlg')), 'double-clicking near its right edge opens the dialog');
const d = await page.evaluate(() => ({
  focus: document.activeElement.getAttribute('aria-label'),
  role: document.querySelector('.odx-chartdlg').getAttribute('role'),
  labelled: !!document.getElementById(document.querySelector('.odx-chartdlg').getAttribute('aria-labelledby')),
  values: [...document.querySelectorAll('.odx-chartdlg-grid tbody input')].map((i) => i.value),
}));
check(d.role === 'dialog' && d.labelled, 'a labelled dialog');
check(d.focus === 'Chart title', 'focus starts on the title');
check(d.values.join('|') === 'North|4.2|5|South|2.5|3.4|West|3.1|2.2', 'the grid holds the categories and values');

// Tab wraps inside the dialog.
await page.locator('.odx-chartdlg button', { hasText: 'Cancel' }).focus();
await page.keyboard.press('Tab');
check(await page.evaluate(() => document.activeElement.getAttribute('aria-label') === 'Chart title'), 'Tab from the last button goes back to the title');

await page.fill('[aria-label="Chart title"]', 'Sales by region');
await page.selectOption('.odx-chartdlg select', 'bar');
await page.fill('[aria-label="2025, North"]', '9');
await page.getByRole('button', { name: 'Add category' }).click();
await page.keyboard.type('East');
await page.getByRole('button', { name: 'Add series' }).click();
await page.keyboard.press('Control+a');
await page.keyboard.type('2027');
await page.fill('[aria-label="2027, East"]', '7');
await page.getByRole('button', { name: 'Apply' }).click();
await page.waitForTimeout(500);
check(!(await page.$('.odx-chartdlg')), 'Apply closes the dialog');
check(await page.evaluate(() => !!document.activeElement.closest('.odx-content')), 'focus goes back to the document');
const after = await src();
check(after.includes('Sales by region') && after.includes('East') && after.includes('2027'), 'the chart is drawn again with the new title, category and series');

await page.keyboard.press('Control+z');
await page.waitForTimeout(300);
check((await src()).includes('>Sales<') && !(await src()).includes('Sales by region'), 'undo brings the old chart back');
await page.keyboard.press('Control+y');
await page.waitForTimeout(300);
check((await src()).includes('Sales by region'), 'redo the edit');

// Enter on the selected chart opens the dialog; Esc closes it without changing anything.
await img.click();
await page.keyboard.press('Enter');
check(!!(await page.$('.odx-chartdlg')), 'Enter on a selected chart opens the dialog');
await page.fill('[aria-label="Chart title"]', 'Not applied');
await page.keyboard.press('Escape');
check(!(await page.$('.odx-chartdlg')) && !(await src()).includes('Not applied'), 'Esc closes it, nothing applied');

const bytes = await page.evaluate(async () => Array.from(await window.odx.save()));
fs.mkdirSync('test/out', { recursive: true });
fs.writeFileSync('test/out/chart-edited.docx', Buffer.from(bytes));
const files = unzipSync(new Uint8Array(bytes));
const chart = strFromU8(files['word/charts/chart1.xml']);
check(/barDir val="bar"/.test(chart) && chart.includes('Sales by region') && chart.includes('<c:v>9</c:v>') && chart.includes('2027'), 'the saved chart part has the edit');
const wb = unzipSync(files['word/embeddings/Microsoft_Excel_Worksheet.xlsx']);
const sheet = strFromU8(wb['xl/worksheets/sheet1.xml']);
check(/<dimension ref="A1:D5"\/>/.test(sheet) && sheet.includes('<c r="D5"><v>7</v></c>') && /ref="A1:D5"/.test(strFromU8(wb['xl/tables/table1.xml'])), 'the saved workbook has the data, its table resized');
check(errors.length === 0, `no page errors ${errors.join('; ')}`);

await browser.close();
server.close();
console.log(failed ? `${failed} failures` : 'charts: all passed');
process.exit(failed ? 1 : 0);
