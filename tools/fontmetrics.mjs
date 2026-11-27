// Reads the vertical metrics of the fonts the editor sizes lines by (TrueType/OpenType files and collections) and writes
// them as one table of numbers, src/fonts/metrics.js, which src/fonts/wordmetrics.js turns into Word's line height and
// ascent. Only numbers leave the font files: the fonts themselves are never copied. Fonts are looked up by their
// family name as Word names them (the name table's family, nameID 1: "Calibri Light", "Arial Narrow"), regular face
// first. A font named in FAMILIES that no folder holds keeps what RECORDED says, or the run stops.
// usage: node tools/fontmetrics.mjs [out.js] [fontDir ...]
// (Run on Windows with Office's fonts: the default folders are the system's and the user's fonts. Aptos, a cloud font,
// may need its folder named: node tools/fontmetrics.mjs src/fonts/metrics.js C:/Windows/Fonts <user fonts> <aptos>.)
import { readFileSync, readdirSync, writeFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';

const out = process.argv[2] || 'src/fonts/metrics.js';
const dirs = process.argv.length > 3 ? process.argv.slice(3) : [
  'C:/Windows/Fonts',
  join(process.env.LOCALAPPDATA || '', 'Microsoft/Windows/Fonts'),
];

// Every font whose numbers the editor uses (props.js lineFactor, baseline.js, mixlines.js, picfix.js, pararules.js).
const FAMILIES = [
  'arial', 'arial black', 'arial narrow', 'times new roman', 'courier new', 'calibri', 'calibri light', 'cambria',
  'consolas', 'segoe ui', 'segoe ui light', 'segoe ui semibold', 'verdana', 'tahoma', 'georgia', 'trebuchet ms',
  'century schoolbook', 'century gothic', 'aptos', 'aptos display', 'aptos narrow', 'candara', 'corbel', 'constantia',
  'comic sans ms', 'impact', 'garamond', 'microsoft sans serif', 'ubuntu mono', 'ms gothic', 'ms pgothic', 'ms ui gothic',
];
// Fonts no folder here holds, as the editor had them: Word's line per em ("line") and the baseline under the line's
// top per em ("ascent"), with where each came from. (From each font's own tables, read where it was found.)
const RECORDED = {
  'open sans': { line: 1.3618 },
  'noto sans': { line: 1.362 },
  'noto serif': { line: 1.362 },
  ubuntu: { line: 1.149, ascent: 0.96 }, // (embedded in sample S125)
  'arial unicode ms': { line: 1.3398 },
  roboto: { line: 1.1719 },
  lato: { line: 1.2 },
};
// Metric-compatible stand-ins and other names for a font: the font whose numbers they take.
const ALIASES = {
  carlito: 'calibri', caladea: 'cambria', 'liberation sans': 'arial', helvetica: 'arial', 'liberation serif': 'times new roman',
  'liberation mono': 'courier new', century: 'century schoolbook',
  'dejavu sans': 'verdana', // (Bitstream Vera's metrics, which Word swaps Verdana for)
  'ms mincho': 'ms gothic', 'ms pmincho': 'ms gothic', // (the same 220 and 36 of 256 as MS Gothic)
};
// Word's own values where they aren't what its rule gives from the installed font (wordmetrics.js), measured. Each
// says from what.
const MEASURED = {
  // (Word's 11.473pt lines at 10pt and 13.769pt at 12pt; the rule gives 1.1475.)
  'arial narrow': { line: 1.1474 },
  // (Embedded in sample S125, its own hhea tables: the installed Ubuntu Mono is another version, 0.932.)
  'ubuntu mono': { ascent: 0.83, wordAscent: 0.83 },
  // (Trebuchet MS and MS Gothic: Word's prints of single-font documents give the rule's 1.1611 and 332/256 (1.2969),
  // 1.1610-1.1616 and 1.2967-1.2973em at 10 to 14pt, 2026-11-24, but the samples come out nearer Word's with these:
  // the rule's put 5 Trebuchet and 2 MS Gothic samples of the 151 14 lines worse. MS Gothic's is Word's checkbox
  // lines, 15.6pt over sample S072's 12pt Times and 13.0pt over the contingency plan's 10pt Calibri. Something else in
  // those lines isn't Word's yet; see TODO.md.)
  'trebuchet ms': { line: 1.1621 },
  'ms gothic': { line: 1.3 },
};
// East Asian fonts, which Word sizes by its own rule (wordmetrics.js eastMetrics).
const EAST = ['ms gothic', 'ms pgothic', 'ms ui gothic', 'ms mincho', 'ms pmincho'];

function tables(buf, off) {
  const n = buf.readUInt16BE(off + 4);
  const t = {};
  for (let i = 0; i < n; i++) {
    const r = off + 12 + i * 16;
    t[buf.toString('latin1', r, r + 4)] = buf.readUInt32BE(r + 8);
  }
  return t;
}

function names(buf, at) {
  const count = buf.readUInt16BE(at + 2);
  const strings = at + buf.readUInt16BE(at + 4);
  const got = {};
  for (let i = 0; i < count; i++) {
    const r = at + 6 + i * 12;
    const platform = buf.readUInt16BE(r), lang = buf.readUInt16BE(r + 4), id = buf.readUInt16BE(r + 6);
    const len = buf.readUInt16BE(r + 8), o = strings + buf.readUInt16BE(r + 10);
    if (![1, 2].includes(id)) continue;
    let s = null;
    if (platform === 3 && (lang === 0x409 || !got[id])) { s = ''; for (let k = 0; k < len; k += 2) s += String.fromCharCode(buf.readUInt16BE(o + k)); }
    else if (platform === 1 && !got[id]) s = buf.toString('latin1', o, o + len);
    if (s != null && (platform === 3 || !got[id])) got[id] = s;
  }
  return got;
}

/** [units per em, hhea ascender, descender (as a positive number), line gap, Windows ascent, Windows descent], and
 * where the font asks for its typo metrics (OS/2 USE_TYPO_METRICS) the typo ascender, descender and line gap after. */
function face(buf, off) {
  const t = tables(buf, off);
  if (t.head == null || t.hhea == null || t['OS/2'] == null || t.name == null) return null;
  const o = t['OS/2'];
  const nm = names(buf, t.name);
  if (!nm[1]) return null;
  const m = [buf.readUInt16BE(t.head + 18), buf.readInt16BE(t.hhea + 4), -buf.readInt16BE(t.hhea + 6), buf.readInt16BE(t.hhea + 8), buf.readUInt16BE(o + 74), buf.readUInt16BE(o + 76)];
  if (buf.readUInt16BE(o + 62) & 0x80) m.push(buf.readInt16BE(o + 68), -buf.readInt16BE(o + 70), buf.readInt16BE(o + 72));
  return { family: nm[1].toLowerCase(), regular: /^(regular|normal|book|roman)$/i.test(nm[2] || ''), m };
}

const fonts = {};
for (const dir of dirs) {
  if (!existsSync(dir)) continue;
  for (const f of readdirSync(dir)) {
    if (!/\.(ttf|otf|ttc)$/i.test(f)) continue;
    let buf;
    try { buf = readFileSync(join(dir, f)); } catch { continue; }
    const offs = buf.toString('latin1', 0, 4) === 'ttcf' ? Array.from({ length: buf.readUInt32BE(8) }, (_, i) => buf.readUInt32BE(12 + i * 4)) : [0];
    for (const off of offs) {
      let got;
      try { got = face(buf, off); } catch { continue; }
      if (!got || !FAMILIES.includes(got.family)) continue;
      // One entry per family: its regular face if there is one (Word's paragraph lines take the regular's).
      const cur = fonts[got.family];
      if (!cur || (got.regular && !cur.regular)) fonts[got.family] = got;
    }
  }
}
const missing = FAMILIES.filter((f) => !fonts[f]);
if (missing.length) { console.error(`not found: ${missing.join(', ')}`); process.exit(1); }

const obj = (o) => `{\n${Object.entries(o).map(([k, v]) => `  ${JSON.stringify(k)}: ${JSON.stringify(v)},`).join('\n')}\n}`;
const js = [
  '// Generated by tools/fontmetrics.mjs from the fonts\' own files (Windows, with Office\'s fonts). Don\'t edit: change the',
  '// tool and run it again. Only numbers: no font is copied. src/fonts/wordmetrics.js turns them into Word\'s line heights.',
  '',
  '/** Font (lowercase family) -> [units per em, hhea ascender, descender, line gap, Windows ascent, Windows descent],',
  ' * then the typo ascender, descender and line gap where the font asks for its typo metrics (USE_TYPO_METRICS). */',
  `export const FONT_UNITS = ${obj(Object.fromEntries(FAMILIES.slice().sort().map((f) => [f, fonts[f].m])))};`,
  '',
  '/** Fonts no file was read for: Word\'s line ("line") and the baseline under the line\'s top ("ascent"), per em. */',
  `export const RECORDED = ${obj(RECORDED)};`,
  '',
  '/** Other names for a font, which take its numbers (metric-compatible stand-ins). */',
  `export const ALIASES = ${obj(ALIASES)};`,
  '',
  '/** Word\'s own values, measured, where they aren\'t its rule\'s (the tool\'s MEASURED says from what). */',
  `export const MEASURED = ${obj(MEASURED)};`,
  '',
  '/** East Asian fonts, which Word sizes by its own rule. */',
  `export const EAST = ${JSON.stringify(EAST)};`,
  '',
].join('\n');
writeFileSync(out, js);
console.log(`${FAMILIES.length} fonts read, ${Object.keys(RECORDED).length} recorded -> ${out}`);
