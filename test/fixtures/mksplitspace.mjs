// Builds test/fixtures/split-row-space.docx: a one-cell table whose row runs across a page, its part on the next page
// starting with a paragraph spaced 195tw before (after an empty one, or straight after a line), as Word lays it out (its
// prints of the same tables, made here and printed by Word).
// With a folder (node test/fixtures/mksplitspace.mjs <dir>) it writes every variant tried in Word there instead.
import '../setup.js';
import fs from 'fs';
import path from 'path';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const out = process.argv[2] ? path.resolve(process.argv[2]) : null;
if (out) fs.mkdirSync(out, { recursive: true });

function write(name, body) {
  const files = unzipSync(blankDocxBytes());
  const main = 'word/document.xml';
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
  let st = strFromU8(files['word/styles.xml']);
  st = st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, '$1<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>$2');
  files['word/styles.xml'] = strToU8(st);
  fs.writeFileSync(out ? path.join(out, name) : new URL(`./${name}`, import.meta.url), zipSync(files));
}

const p = (t, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${t ? `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>` : ''}</w:p>`;
// `n` lines, then (empty paragraph or not) a paragraph 195tw before, then more lines, in one cell of one row.
const table = (n, empty, after = 0, before = 195) => {
  const A = after ? `<w:spacing w:after="${after}"/>` : '';
  const ps = Array.from({ length: n }, (_, k) => p(`Line ${k + 1}`, k === n - 1 ? A : '')).join('') + (empty ? p('', A) : '') + p('Spaced item', `<w:spacing w:before="${before}"/>`) + Array.from({ length: 5 }, (_, k) => p(`After ${k + 1}`)).join('');
  const b = ['top', 'left', 'bottom', 'right'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('');
  return `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:tblBorders>${b}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${ps}</w:tc></w:tr></w:tbl>${p('After table')}`;
};
const variants = {};
for (const n of [53, 54, 55, 56, 57]) { variants[`s${n}-empty`] = table(n, true); variants[`s${n}-direct`] = table(n, false); }
// The line left behind spaced after (80tw, 240tw), the carried one 40tw or 195tw before (sample 1283's table text: 40 before, 80 after).
for (const [af, bf] of [[80, 40], [80, 195], [240, 195], [0, 40]]) variants[`a${af}-b${bf}`] = table(55, false, af, bf);
if (out) {
  for (const [n, b] of Object.entries(variants)) write(`${n}.docx`, b);
  console.log(Object.keys(variants).length, 'documents in', out);
} else {
  write('split-row-space.docx', variants['s55-empty']);
}
