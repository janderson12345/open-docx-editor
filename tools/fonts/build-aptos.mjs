// Builds the Aptos stand-ins in src/editor/editor.css: fonts made from the bundled Arimo (Arial's design) in which every
// letter has the advance the editor gives it for Aptos (fitted to Word's PDFs) and one height. Before, Arial was drawn
// at a separate size per letter to reach those widths, and the size changed each letter's height as well: a capital
// T or J came out small, a t or an l tall. Now each letter is narrowed or widened sideways only, so lines break where
// they did and the text is even.
//
// It measures each letter's advance as editor.css draws it now (in Microsoft Edge, which has Arial), builds one font
// per face (tools/fonts/aptos_standin.py, Python with fontTools and brotli; modified versions of Arimo, under its SIL
// Open Font License 1.1, with names of their own), and writes them into editor.css between
// its markers, removing the per-letter faces they replace. Run again after a change to the fitted widths.
//
//   node tools/fonts/build-aptos.mjs
import { chromium } from 'playwright';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const CSS = 'src/editor/editor.css';
const FILES = 'src/fonts/files';
const BEGIN = '/* Aptos stand-ins (tools/fonts/build-aptos.mjs):';
const END = '/* end of the Aptos stand-ins */';
// The line every Aptos face keeps (its ascent, descent and line gap, in ems of the font size): Arial's at 97%.
const LINE = { ascent: 87.79, descent: 20.57, gap: 3.17 };
const FACES = [];
for (const family of ['odx-aptos', 'odx-aptos-display']) {
  for (const [weight, style] of [[400, 'normal'], [700, 'normal'], [400, 'italic'], [700, 'italic']]) {
    FACES.push({ family, weight, style, src: `${FILES}/arimo-latin-${weight}-${style}.woff2` });
  }
}
const py = (...args) => execFileSync('python', ['-I', 'tools/fonts/aptos_standin.py', ...args], { encoding: 'utf8', maxBuffer: 1 << 26 });

let css = fs.readFileSync(CSS, 'utf8');
const arimo = {};
for (const f of FACES) arimo[f.src] ??= JSON.parse(py('cmap', f.src));
const points = (src) => Object.keys(arimo[src]).map(Number).filter((cp) => cp >= 0x20 && !(cp >= 0x7f && cp < 0xa0));

// Measure what editor.css draws now, at 1000px, in the editor's own stack ("odx-aptos", sans-serif).
const browser = await chromium.launch({ channel: process.env.BROWSER_CHANNEL || 'msedge' });
const page = await browser.newPage();
await page.setContent(`<style>${css}</style><div id="m"></div>`);
const measured = await page.evaluate(async (faces) => {
  const out = [];
  const box = document.getElementById('m');
  for (const f of faces) {
    const font = `${f.style} ${f.weight} 1000px '${f.family}'`;
    const text = f.points.map((cp) => String.fromCodePoint(cp)).join('');
    await document.fonts.load(font, text);
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
}, FACES.map((f) => ({ ...f, points: points(f.src) })));
await browser.close();

// Each family's one size, for all its faces (bold drawn shorter than the text around it would look as uneven): the
// median of its regular lowercase letters' sizes now, so running text looks as it did.
const median = (a) => { const s = [...a].sort((x, y) => x - y); return s[s.length >> 1]; };
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'odx-aptos-'));
const sizes = {};
FACES.forEach((f, i) => {
  if (f.weight !== 400 || f.style !== 'normal') return;
  sizes[f.family] = Math.round(median([...'abcdefghijklmnopqrstuvwxyz'].map((c) => measured[i][c.codePointAt(0)] / arimo[f.src][c.codePointAt(0)])) * 10000) / 10000;
});
const spec = { faces: FACES.map((f, i) => {
  const widths = measured[i];
  const scale = sizes[f.family];
  f.scale = scale;
  f.out = path.join(tmp, `${f.family}-${f.weight}-${f.style}.woff2`);
  return { src: f.src, out: f.out, scale, name: f.family === 'odx-aptos' ? 'odx Standin Sans' : 'odx Standin Display',
    style: `${f.weight === 700 ? 'Bold' : ''}${f.style === 'italic' ? ' Italic' : ''}`.trim() || 'Regular', widths };
}) };
fs.writeFileSync(path.join(tmp, 'spec.json'), JSON.stringify(spec));
py('build', path.join(tmp, 'spec.json'));

// Remove the per-letter faces the new fonts replace: same family, weight and style, with a unicode-range inside theirs.
const range = (cps) => {
  const out = []; let a = null; let b = null;
  for (const cp of [...cps].sort((x, y) => x - y)) { if (b !== null && cp === b + 1) b = cp; else { if (a !== null) out.push([a, b]); a = b = cp; } }
  if (a !== null) out.push([a, b]);
  return out.map(([x, y]) => (x === y ? `U+${x.toString(16).toUpperCase().padStart(4, '0')}` : `U+${x.toString(16).toUpperCase().padStart(4, '0')}-${y.toString(16).toUpperCase().padStart(4, '0')}`)).join(', ');
};
const covers = (rule, f) => {
  const fam = /font-family: '([^']+)'/.exec(rule)?.[1];
  const w = /font-weight: (\w+)/.exec(rule)?.[1] || 'normal';
  const st = /font-style: (\w+)/.exec(rule)?.[1] || 'normal';
  const ur = /unicode-range: ([^;]+);/.exec(rule)?.[1];
  if (fam !== f.family || (w === 'bold' || w === '700' ? 700 : 400) !== f.weight || st !== f.style || !ur) return false;
  const set = new Set(points(f.src));
  return ur.split(',').every((part) => {
    const [x, y = x] = part.trim().replace(/^U\+/, '').split('-').map((h) => parseInt(h, 16));
    for (let cp = x; cp <= y; cp++) if (!set.has(cp)) return false;
    return true;
  });
};
const start = css.indexOf(BEGIN);
if (start >= 0) css = css.slice(0, start) + css.slice(css.indexOf(END) + END.length + 1);
const lines = css.split('\n');
const kept = lines.filter((l) => !(l.startsWith('@font-face') && FACES.some((f) => covers(l, f))));
const removed = lines.length - kept.length;
const pct = (v) => `${Math.round(v * 100) / 100}%`;
const block = [
  `${BEGIN} Arimo's letters (Arial's design; SIL Open Font License 1.1,`,
  '   src/fonts/LICENSE.md), each narrowed or widened sideways to the advance Aptos gives it (fitted to Word\'s PDFs), all at',
  '   one size: the text is even and lines break as in Word. Declared last, so they come before the faces above, which',
  '   stay for letters these fonts lack. */',
  ...FACES.map((f) => `@font-face { font-family: '${f.family}'; src: url(data:font/woff2;base64,${fs.readFileSync(f.out).toString('base64')}) format('woff2'); unicode-range: ${range(Object.keys(spec.faces[FACES.indexOf(f)].widths).map(Number))}; font-weight: ${f.weight === 700 ? 700 : 'normal'}; font-style: ${f.style}; size-adjust: ${pct(f.scale * 100)}; ascent-override: ${pct(LINE.ascent / f.scale)}; descent-override: ${pct(LINE.descent / f.scale)}; line-gap-override: ${pct(LINE.gap / f.scale)}; }`),
  END,
];
// After the last remaining Aptos face.
let at = -1;
kept.forEach((l, i) => { if (/^@font-face \{ font-family: 'odx-aptos(-display)?'/.test(l)) at = i; });
kept.splice(at + 1, 0, ...block);
fs.writeFileSync(CSS, kept.join('\n'));
console.log(FACES.map((f) => `${f.family} ${f.weight} ${f.style}: ${pct(f.scale * 100)}, ${fs.statSync(f.out).size} bytes`).join('\n'));
console.log(`removed ${removed} per-letter faces`);
