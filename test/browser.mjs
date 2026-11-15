// What the browser tests share: the demo page served from a port of its own, the browser launched as CI and local
// runs ask (CHROMIUM_PATH, BROWSER_CHANNEL), opening a fixture in the demo, and the ok/FAIL report with its exit code.
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';

const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' };

/** A static server for a folder (the built demo by default) on a free port: { server, url }. */
export function serveDemo(dir = 'demo', { headers = {} } = {}) {
  const root = path.resolve(dir);
  const server = http.createServer((req, res) => {
    const p = path.join(root, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
    if (!p.startsWith(root) || !fs.existsSync(p)) { res.writeHead(404); return res.end(); }
    res.writeHead(200, { 'content-type': TYPES[path.extname(p)] || 'application/octet-stream', ...headers });
    fs.createReadStream(p).pipe(res);
  }).listen(0);
  return { server, url: `http://localhost:${server.address().port}/` };
}

/** Playwright's launch options: a given Chromium, a channel (msedge), or its own. */
export const launchOptions = () => (process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : process.env.BROWSER_CHANNEL ? { channel: process.env.BROWSER_CHANNEL } : {});

/** Opens a test fixture in the demo page and waits for it to be laid out. */
export async function openFixture(page, file, settle = 1200) {
  const name = path.basename(file);
  await page.setInputFiles('input[type=file]', file.includes('/') ? file : `test/fixtures/${file}`);
  await page.waitForFunction((n) => window.odx?.fileName === n, name);
  await page.waitForTimeout(settle);
}

/** The ok/FAIL lines of a run, and its summary and exit code at the end. */
export function reporter(label) {
  let failed = 0;
  return {
    check(ok, what) { console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}`); if (!ok) failed++; },
    get failed() { return failed; },
    finish() { console.log(failed ? `${label}: ${failed} failed` : `${label}: all passed`); process.exit(failed ? 1 : 0); },
  };
}
