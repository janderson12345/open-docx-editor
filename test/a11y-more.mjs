// More accessibility checks on the rendered editor, for what axe doesn't measure (run with the e2e tests):
// 1. 1.4.11 Non-text contrast: ribbon icons (drawn in currentColor) and form-control borders against what's behind them,
//    at least 3:1, light and dark.
// 2. 1.4.10 Reflow / 1.4.4 Resize: at 320 CSS px wide (1280 at 400%) and at 200% zoom (700 px), the ribbon stays one row
//    with every group either shown or behind its More button, and nothing of the page but the document sheet (a
//    two-dimensional layout, exempt) scrolls sideways.
// 3. 1.4.12 Text spacing: with WCAG's text-spacing overrides on the UI, no control's text is cut off.
// Exceptions are listed by name, with the reason.
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
const url = `http://localhost:${server.address().port}/`;
const launch = process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : process.env.BROWSER_CHANNEL ? { channel: process.env.BROWSER_CHANNEL } : {};
const browser = await chromium.launch(launch);
let failed = 0;
const fail = (msg) => { failed++; console.log('FAIL', msg); };
const FLOORS = { icons: 12, fields: 3 };

for (const scheme of ['light', 'dark']) {
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 }, colorScheme: scheme });
  await page.goto(url);
  await page.waitForSelector('.odx-content p');
  // 1. Non-text contrast, on every ribbon tab.
  const tabs = await page.$$eval('.odx-rtab', (b) => b.map((x) => x.textContent).filter((t) => t !== 'File'));
  const seen = { icons: 0, fields: 0 };
  for (const tab of tabs) {
    await page.click(`.odx-rtab:text-is("${tab}")`);
    const r = await page.evaluate(() => {
      const rgb = (c) => { const m = /rgba?\(([\d.]+), ([\d.]+), ([\d.]+)(?:, ([\d.]+))?/.exec(c); return m ? [+m[1], +m[2], +m[3], m[4] == null ? 1 : +m[4]] : null; };
      const lum = ([r, g, b]) => { const f = (v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; }; return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b); };
      const ratio = (a, b) => { const [x, y] = [lum(a), lum(b)].sort((p, q) => q - p); return (x + 0.05) / (y + 0.05); };
      const bg = (e) => { for (let n = e; n; n = n.parentElement) { const c = rgb(getComputedStyle(n).backgroundColor); if (c && c[3] > 0.5) return c; } return [255, 255, 255, 1]; };
      const out = { icons: 0, fields: 0, low: [] };
      const panel = document.querySelector('.odx-panel:not([hidden])');
      for (const svg of panel.querySelectorAll('button svg')) {
        const b = svg.closest('button');
        if (b.disabled || !svg.getClientRects().length || b.closest('.odx-cut')) continue;
        out.icons++;
        const k = ratio(rgb(getComputedStyle(svg).color), bg(b));
        if (k < 3) out.low.push(`icon ${b.getAttribute('aria-label') || b.title} ${k.toFixed(2)}`);
      }
      for (const f of panel.querySelectorAll('select, input:not([type=checkbox]):not([type=radio]):not([type=color])')) {
        if (!f.getClientRects().length || f.closest('.odx-cut')) continue;
        out.fields++;
        const cs = getComputedStyle(f);
        const k = ratio(rgb(cs.borderBottomColor) || [0, 0, 0], bg(f.parentElement));
        if (parseFloat(cs.borderBottomWidth) > 0 && k < 3) out.low.push(`field ${f.getAttribute('aria-label') || f.title} border ${k.toFixed(2)}`);
      }
      return out;
    });
    seen.icons += r.icons;
    seen.fields += r.fields;
    for (const l of r.low) fail(`${scheme} ${tab}: 1.4.11 ${l}`);
  }
  for (const [k, n] of Object.entries(FLOORS)) if (seen[k] < n) fail(`${scheme}: only ${seen[k]} ${k} checked (floor ${n})`);
  console.log(`${scheme}: 1.4.11 checked ${seen.icons} icons and ${seen.fields} fields`);

  // 2. Reflow and resize: 320 px (400% zoom) and 700 px (200%).
  for (const width of [320, 700]) {
    await page.setViewportSize({ width, height: 900 });
    await page.waitForTimeout(300);
    const r = await page.evaluate(() => {
      const bar = document.querySelector('.odx-toolbar');
      const panel = bar.querySelector('.odx-panel:not([hidden])');
      const more = bar.querySelector('.odx-more');
      const rows = new Set([...panel.children].filter((g) => !g.hidden && !g.classList.contains('odx-cut')).map((g) => Math.round(g.getBoundingClientRect().top)));
      const cut = panel.querySelectorAll('.odx-cut').length;
      // What scrolls sideways, other than the document's own sheets.
      const wide = [...document.querySelectorAll('body *')].filter((e) => {
        if (e.closest('.odx-page, .odx-sheets, .odx-content, .odx-live')) return false;
        const cs = getComputedStyle(e);
        return /auto|scroll/.test(cs.overflowX) && e.scrollWidth > e.clientWidth + 1 && !e.classList.contains('odx-canvas');
      }).map((e) => e.className || e.tagName);
      return { rows: rows.size, cut, moreShown: more && !more.hidden, barWide: bar.scrollWidth > bar.clientWidth + 1, wide };
    });
    if (r.rows > 1) fail(`${scheme} ${width}px: ribbon groups on ${r.rows} rows`);
    if (r.cut && !r.moreShown) fail(`${scheme} ${width}px: ${r.cut} groups hidden with no More button`);
    if (r.barWide) fail(`${scheme} ${width}px: the ribbon scrolls sideways`);
    for (const w of r.wide) fail(`${scheme} ${width}px: ${w} scrolls sideways`);
    if (r.cut) {
      await page.click('.odx-more');
      const all = await page.evaluate(() => { const p = document.querySelector('.odx-panel:not([hidden])'); return [...p.children].every((g) => g.hidden || (!g.classList.contains('odx-cut') && g.getBoundingClientRect().right <= p.getBoundingClientRect().right + 1)); });
      if (!all) fail(`${scheme} ${width}px: More doesn't bring every group into view`);
      await page.click('.odx-more');
    }
    console.log(`${scheme} ${width}px: ${r.cut} groups behind More`);
  }
  await page.setViewportSize({ width: 1400, height: 900 });

  // 3. Text spacing (WCAG 1.4.12's values) on the UI: line height 1.5, paragraph spacing 2em, letter 0.12em, word 0.16em.
  await page.addStyleTag({ content: '.odx-toolbar *, .odx-pop *, .odx-findbar * { line-height: 1.5 !important; letter-spacing: 0.12em !important; word-spacing: 0.16em !important; } .odx-toolbar p { margin-bottom: 2em !important; }' });
  await page.waitForTimeout(300);
  const clipped = await page.evaluate(() => [...document.querySelectorAll('.odx-panel:not([hidden]) button, .odx-rtab, .odx-glabel')].filter((e) => {
    if (!e.getClientRects().length || e.closest('.odx-cut, .odx-stile')) return false; // (style tiles are previews, cut on purpose)
    const cs = getComputedStyle(e);
    const hides = /hidden|clip/.test(cs.overflowX + cs.overflowY) || cs.textOverflow === 'ellipsis';
    return hides && (e.scrollWidth > e.clientWidth + 1 || e.scrollHeight > e.clientHeight + 1);
  }).map((e) => e.getAttribute('aria-label') || e.textContent.trim().slice(0, 20)));
  for (const c of clipped) fail(`${scheme}: 1.4.12 text cut off in ${c}`);
  console.log(`${scheme}: 1.4.12 ${clipped.length} cut off`);
  await page.close();
}
await browser.close();
server.close();
console.log(failed ? `${failed} failures` : 'all passed');
process.exit(failed ? 1 : 0);
