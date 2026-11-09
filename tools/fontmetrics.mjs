// Reads the vertical metrics of the fonts installed on this machine (TrueType/OpenType files and collections) and writes
// them as a table of numbers for the layout engine (src/engine/metrics.json). Only numbers leave the font files: the
// fonts themselves are never copied. Word's single line for a font is its hhea ascender + descender + line gap, its
// baseline the ascender + line gap below the line's top (checked against Word's PDFs: Calibri 1.2207em, Times New Roman
// 1.1499em).
// usage: node tools/fontmetrics.mjs [out.json] [fontDir ...]
import { readFileSync, readdirSync, writeFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';

const out = process.argv[2] || 'src/engine/metrics.json';
const dirs = process.argv.length > 3 ? process.argv.slice(3) : [
  'C:/Windows/Fonts',
  join(process.env.LOCALAPPDATA || '', 'Microsoft/Windows/Fonts'),
];

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
    if (![1, 2, 4, 16, 17].includes(id)) continue;
    let s = null;
    if (platform === 3 && (lang === 0x409 || !got[id])) { s = ''; for (let k = 0; k < len; k += 2) s += String.fromCharCode(buf.readUInt16BE(o + k)); }
    else if (platform === 1 && !got[id]) s = buf.toString('latin1', o, o + len);
    if (s != null && (platform === 3 || !got[id])) got[id] = s;
  }
  return got;
}

function face(buf, off) {
  const t = tables(buf, off);
  if (t.head == null || t.hhea == null) return null;
  const upm = buf.readUInt16BE(t.head + 18);
  const asc = buf.readInt16BE(t.hhea + 4), desc = buf.readInt16BE(t.hhea + 6), gap = buf.readInt16BE(t.hhea + 8);
  let os2 = null;
  if (t['OS/2'] != null) {
    const o = t['OS/2'];
    os2 = { typoAsc: buf.readInt16BE(o + 68), typoDesc: buf.readInt16BE(o + 70), typoGap: buf.readInt16BE(o + 72), winAsc: buf.readUInt16BE(o + 74), winDesc: buf.readUInt16BE(o + 76), useTypo: !!(buf.readUInt16BE(o + 62) & 0x80) };
  }
  const nm = t.name != null ? names(buf, t.name) : {};
  const family = nm[16] || nm[1];
  const sub = (nm[17] || nm[2] || 'Regular').toLowerCase();
  if (!family) return null;
  const r = (v) => Math.round((v / upm) * 10000) / 10000;
  return {
    family: family.toLowerCase(), style: sub,
    m: { line: r(asc - desc + gap), ascent: r(asc + gap), hhea: [r(asc), r(-desc), r(gap)], win: os2 ? [r(os2.winAsc), r(os2.winDesc)] : null, typo: os2 ? [r(os2.typoAsc), r(-os2.typoDesc), r(os2.typoGap), os2.useTypo] : null },
  };
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
      if (!got) continue;
      // One entry per family: its regular face if there is one (line metrics rarely differ by weight).
      const cur = fonts[got.family];
      if (!cur || (/regular|normal|book|roman/.test(got.style) && !cur.regular)) fonts[got.family] = { ...got.m, regular: /regular|normal|book|roman/.test(got.style) };
    }
  }
}
for (const v of Object.values(fonts)) delete v.regular;
const sorted = Object.fromEntries(Object.entries(fonts).sort((a, b) => a[0].localeCompare(b[0])));
writeFileSync(out, JSON.stringify(sorted, null, 0).replace(/\},"/g, '},\n"') + '\n');
console.log(`${Object.keys(sorted).length} families -> ${out}`);
