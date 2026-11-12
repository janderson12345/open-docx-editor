// Prints every .docx in a folder to PDF through the editor, as tools/word-compare/print-ours.mjs does, but quicker for
// thousands of small documents: each browser page is reused for document after document, a document is opened with
// the editor's own open() and printed when it says it's laid out (and its fonts are loaded), instead of after fixed
// waits. Documents already printed are skipped; each has a time limit.
// usage: node tools/doc-generator/print-fast.mjs <docxDir> <outDir>   (from the repository, after `node build.mjs`)
//   BROWSER_CHANNEL=msedge (default) | chrome; CHROMIUM_PATH=<executable> instead; ODX_JOBS=16
import { chromium } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';

const [dir, outDir] = process.argv.slice(2);
if (!outDir) { console.log('usage: node print-fast.mjs <docxDir> <outDir>'); process.exit(1); }
const root = path.resolve('demo');
const server = http.createServer((req, res) => {
  const p = path.join(root, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
  if (!p.startsWith(root) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'content-type': { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' }[path.extname(p)] || 'application/octet-stream' });
  fs.createReadStream(p).pipe(res);
}).listen(0);
const url = `http://localhost:${server.address().port}/`;
fs.mkdirSync(outDir, { recursive: true });
const browser = await chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : { channel: process.env.BROWSER_CHANNEL || 'msedge' });
const todo = fs.readdirSync(dir).filter((f) => /\.docx$/i.test(f) && !fs.existsSync(path.join(outDir, path.basename(f, '.docx') + '.pdf')));
const total = todo.length;
let done = 0;

const fresh = async () => {
  const page = await browser.newPage({ viewport: { width: 1300, height: 1000 } });
  await page.goto(url);
  await page.waitForFunction(() => window.odx);
  return page;
};
const one = async (page, f) => {
  const b64 = fs.readFileSync(path.join(dir, f)).toString('base64');
  await page.emulateMedia({ media: 'screen' });
  await page.evaluate(async ([data, name]) => {
    const bytes = Uint8Array.from(atob(data), (c) => c.charCodeAt(0));
    await window.odx.open(bytes.buffer, name);
    await document.fonts.ready;
    // Scroll through it once, so every page has been drawn (long documents draw off-screen pages lazily).
    const c = document.querySelector('.odx-canvas');
    for (let y = 0; y < c.scrollHeight; y += 800) { c.scrollTop = y; await new Promise((r) => requestAnimationFrame(() => setTimeout(r, 0))); }
    c.scrollTop = 0;
    await document.fonts.ready;
    // Then until the pages stop changing for 300ms (columns and justified lines settle a pass or two later), at most 5s.
    const sig = () => { const p = document.querySelectorAll('.odx-canvas *').length; return `${p}:${c.scrollHeight}:${c.innerHTML.length}`; };
    let last = sig(), still = 0;
    for (let t = 0; t < 5000 && still < 300; t += 50) { await new Promise((r) => setTimeout(r, 50)); const now = sig(); still = now === last ? still + 50 : 0; last = now; }
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
  }, [b64, f]);
  await page.emulateMedia({ media: 'print' });
  await page.evaluate(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
  await page.pdf({ path: path.join(outDir, path.basename(f, '.docx') + '.pdf'), preferCSSPageSize: true, printBackground: true });
};

const jobs = Number(process.env.ODX_JOBS || 16);
await Promise.all(Array.from({ length: jobs }, async () => {
  let page = await fresh();
  while (todo.length) {
    const f = todo.shift();
    try {
      await Promise.race([one(page, f), new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), 120000))]);
      if (++done % 250 === 0) console.log(`printed ${done} of ${total}`);
    } catch (e) {
      console.log('FAILED', f, e.message);
      await page.close().catch(() => {});
      page = await fresh(); // a document that broke the page doesn't break the next
    }
  }
  await page.close();
}));
console.log(`printed ${done}`);
// (Without these the process waits about two minutes for the pages' idle connections to the server to time out.)
server.closeAllConnections();
server.close();
await browser.close();
process.exit(0);
