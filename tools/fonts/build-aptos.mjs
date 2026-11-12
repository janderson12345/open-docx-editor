// Builds the Aptos stand-ins (src/fonts/files/odx-aptos-*.woff2 and their @font-face lines in src/fonts/fonts.css):
// fonts made from Source Sans 3 (Adobe; SIL Open Font License 1.1) in which every letter has the advance the editor
// gives it for Aptos (fitted to Word's PDFs, as the stand-ins have now) and Aptos's x-height. Each letter is narrowed or
// widened sideways only, so lines break where Word breaks them and the text keeps one height; Source Sans 3's letters
// (its a, t, y, G...) are closer to Aptos's than Arial's design, which these were made from before.
//
// It measures each letter's advance as the stand-ins draw it now (in Microsoft Edge), builds one font per face
// (tools/fonts/aptos_standin.py, Python with fontTools and brotli; modified versions of Source Sans 3, under its
// license, with names of their own: its Reserved Font Name "Source" is not used), and writes them with their CSS.
// Run again after a change to the fitted widths.
//
//   node tools/fonts/build-aptos.mjs <folder with SourceSans3-Regular.ttf, -Bold.ttf, -It.ttf, -BoldIt.ttf>
// (Source Sans 3's TTFs: github.com/adobe-fonts/source-sans releases, TTF-source-sans-*.zip.)
import { chromium } from 'playwright';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const SRC_DIR = process.argv[2];
if (!SRC_DIR) throw new Error('usage: node tools/fonts/build-aptos.mjs <folder with Source Sans 3 TTFs>');
const CSS = 'src/editor/editor.css';
const FONTS_CSS = 'src/fonts/fonts.css';
const FILES = 'src/fonts/files';
const BEGIN = '/* Aptos stand-ins (tools/fonts/build-aptos.mjs):';
const END = '/* end of the Aptos stand-ins */';
// Aptos's x-height in ems (its OS/2 sxHeight, 974 of 2048 units; the same in Aptos Display).
const APTOS_XHEIGHT = 0.4756;
// The line every Aptos face keeps (its ascent, descent and line gap, in ems of the font size): Arial's at 97%.
const LINE = { ascent: 87.79, descent: 20.57, gap: 3.17 };
const SOURCE = { '400-normal': 'SourceSans3-Regular.ttf', '700-normal': 'SourceSans3-Bold.ttf', '400-italic': 'SourceSans3-It.ttf', '700-italic': 'SourceSans3-BoldIt.ttf' };
const FACES = [];
for (const family of ['odx-aptos', 'odx-aptos-display']) {
  for (const [weight, style] of [[400, 'normal'], [700, 'normal'], [400, 'italic'], [700, 'italic']]) {
    FACES.push({ family, weight, style, src: path.join(SRC_DIR, SOURCE[`${weight}-${style}`]), file: `${family}-${weight}-${style}.woff2`,
      // (The letters measured: those the Arimo-based stand-ins covered.)
      was: `${FILES}/arimo-latin-${weight}-${style}.woff2` });
  }
}
const py = (...args) => execFileSync('python', ['-I', 'tools/fonts/aptos_standin.py', ...args], { encoding: 'utf8', maxBuffer: 1 << 26 });
const cmaps = {};
const cmapOf = (p) => (cmaps[p] ??= JSON.parse(py('cmap', p)));
const points = (f) => Object.keys(cmapOf(f.was)).map(Number).filter((cp) => cp >= 0x20 && !(cp >= 0x7f && cp < 0xa0) && cmapOf(f.src)[cp] != null);

// Measure what the stand-ins draw now, at 1000px, in the editor's own stack ("odx-aptos", sans-serif): a page beside
// fonts.css, so its files load as they do for the editor.
const page0 = path.resolve('src/fonts/_measure.html');
fs.writeFileSync(page0, `<link rel="stylesheet" href="../editor/editor.css"><link rel="stylesheet" href="fonts.css"><div id="m"></div>`);
const browser = await chromium.launch({ channel: process.env.BROWSER_CHANNEL || 'msedge' });
const page = await browser.newPage();
await page.goto(pathToFileURL(page0).href);
const measured = await page.evaluate(async (faces) => {
  const out = [];
  const box = document.getElementById('m');
  for (const f of faces) {
    const font = `${f.style} ${f.weight} 1000px '${f.family}'`;
    await document.fonts.load(font, f.points.map((cp) => String.fromCodePoint(cp)).join(''));
    const widths = {};
    for (const cp of f.points) {
      const s = document.createElement('span');
      s.style.cssText = `font-family:'${f.family}',sans-serif;font-weight:${f.weight};font-style:${f.style};font-size:1000px;white-space:pre;font-kerning:none;font-variant-ligatures:none`;
      s.textContent = String.fromCodePoint(cp);
      box.appendChild(s);
      widths[cp] = s.getBoundingClientRect().width / 1000;
      s.remove();
    }
    out.push(widths);
  }
  return out;
}, FACES.map((f) => ({ ...f, points: points(f) })));
await browser.close();
fs.rmSync(page0);

// One size for all faces: Source Sans 3's x-height made Aptos's (its OS/2 sxHeight; the CSS size-adjust).
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'odx-aptos-'));
const spec = { faces: FACES.map((f, i) => {
  f.scale = Math.round((APTOS_XHEIGHT / Number(py('xheight', f.src))) * 10000) / 10000;
  f.out = path.join(FILES, f.file);
  f.widths = measured[i];
  return { src: f.src, out: f.out, scale: f.scale, name: f.family === 'odx-aptos' ? 'odx Standin Sans' : 'odx Standin Display',
    style: `${f.weight === 700 ? 'Bold' : ''}${f.style === 'italic' ? ' Italic' : ''}`.trim() || 'Regular', widths: f.widths };
}) };
fs.writeFileSync(path.join(tmp, 'spec.json'), JSON.stringify(spec));
py('build', path.join(tmp, 'spec.json'));

const range = (cps) => {
  const out = []; let a = null; let b = null;
  for (const cp of [...cps].sort((x, y) => x - y)) { if (b !== null && cp === b + 1) b = cp; else { if (a !== null) out.push([a, b]); a = b = cp; } }
  if (a !== null) out.push([a, b]);
  return out.map(([x, y]) => (x === y ? `U+${x.toString(16).toUpperCase().padStart(4, '0')}` : `U+${x.toString(16).toUpperCase().padStart(4, '0')}-${y.toString(16).toUpperCase().padStart(4, '0')}`)).join(', ');
};
const pct = (v) => `${Math.round(v * 100) / 100}%`;
const block = [
  `${BEGIN} Source Sans 3's letters (Adobe; SIL Open Font License 1.1,`,
  '   LICENSE.md), each narrowed or widened sideways to the advance Aptos gives it (fitted to Word\'s PDFs), at Aptos\'s',
  '   x-height: lines break as in Word and the letters look like Aptos\'s. The editor\'s own faces (editor.css: Arial at',
  '   Aptos\'s average width) stay for letters these fonts lack. */',
  ...FACES.map((f) => `@font-face { font-family: '${f.family}'; src: url('./files/${f.file}') format('woff2'); unicode-range: ${range(Object.keys(f.widths).map(Number))}; font-weight: ${f.weight === 700 ? 700 : 'normal'}; font-style: ${f.style}; font-display: swap; size-adjust: ${pct(f.scale * 100)}; ascent-override: ${pct(LINE.ascent / f.scale)}; descent-override: ${pct(LINE.descent / f.scale)}; line-gap-override: ${pct(LINE.gap / f.scale)}; }`),
  END,
];
// Out of editor.css (the stand-ins once lived there, as data), into fonts.css (replacing an earlier block).
const strip = (text) => { const s = text.indexOf(BEGIN); return s < 0 ? text : text.slice(0, s) + text.slice(text.indexOf(END) + END.length + 1); };
fs.writeFileSync(CSS, strip(fs.readFileSync(CSS, 'utf8')));
fs.writeFileSync(FONTS_CSS, strip(fs.readFileSync(FONTS_CSS, 'utf8')).replace(/\n*$/, '\n') + block.join('\n') + '\n');
console.log(FACES.map((f) => `${f.family} ${f.weight} ${f.style}: ${pct(f.scale * 100)}, ${fs.statSync(f.out).size} bytes`).join('\n'));
