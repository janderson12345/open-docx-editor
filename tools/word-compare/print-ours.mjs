// Prints every .docx in a folder to PDF through the editor (its print layout, as a person printing would), for
// comparison with Word's PDFs of the same files. Several documents at a time; each one has a time limit, so one bad
// file can't hang the run; documents already printed are skipped, so a rerun picks up where it stopped.
// usage: node tools/word-compare/print-ours.mjs <docxDir> <outDir>   (from the repository, after `node build.mjs`: the demo
//   page in demo/ is served as the browser tests serve it)
//   BROWSER_CHANNEL=msedge (default) | chrome; CHROMIUM_PATH=<executable> instead; ODX_JOBS=8
// Use the browser your users print from: its fonts and font rendering are part of what's measured (Edge on Windows
// for the published numbers).
import { chromium } from 'playwright';
import http from 'node:http';
import fs, { readdirSync, mkdirSync, existsSync } from 'node:fs';
import path, { join, basename } from 'node:path';

const [dir, outDir] = process.argv.slice(2);
if (!outDir) { console.log('usage: node print-ours.mjs <docxDir> <outDir>'); process.exit(1); }
const root = path.resolve('demo');
const server = http.createServer((req, res) => {
  const p = path.join(root, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
  if (!p.startsWith(root) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'content-type': { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' }[path.extname(p)] || 'application/octet-stream' });
  fs.createReadStream(p).pipe(res);
}).listen(0);
const url = `http://localhost:${server.address().port}/`;
mkdirSync(outDir, { recursive: true });
const browser = await chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : { channel: process.env.BROWSER_CHANNEL || 'msedge' });
const todo = readdirSync(dir).filter((f) => /\.docx$/i.test(f) && !existsSync(join(outDir, basename(f, '.docx') + '.pdf')));
let done = 0;
const one = async (f) => {
  const page = await browser.newPage({ viewport: { width: 1300, height: 1000 } });
  try {
    await Promise.race([(async () => {
      await page.goto(url);
      await page.waitForTimeout(600);
      await page.setInputFiles('input[type=file]', join(dir, f));
      await page.waitForTimeout(2500);
      // Scroll through it once, so every page has been drawn (long documents draw off-screen pages lazily).
      await page.evaluate(async () => { const c = document.querySelector('.odx-canvas'); for (let y = 0; y < c.scrollHeight; y += 800) { c.scrollTop = y; await new Promise((r) => setTimeout(r, 20)); } });
      await page.waitForTimeout(600);
      await page.emulateMedia({ media: 'print' });
      await page.waitForTimeout(400);
      await page.pdf({ path: join(outDir, basename(f, '.docx') + '.pdf'), preferCSSPageSize: true, printBackground: true });
    })(), new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), 150000))]);
    done++;
  } catch (e) { console.log('FAILED', f, e.message); }
  await page.close();
};
const jobs = Number(process.env.ODX_JOBS || 8);
await Promise.all(Array.from({ length: jobs }, async () => { while (todo.length) await one(todo.shift()); }));
await browser.close();
server.close();
console.log(`printed ${done}`);
