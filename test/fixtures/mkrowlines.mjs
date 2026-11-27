// Builds test/fixtures/row-dash-lines.docx: rows of a minimum height (w:trHeight, at least) under dotted and dashed
// lines, as Word lays them (its prints of the same tables, made here and printed by Word): a minimum leaves out the
// row's top line, dotted or dashed as solid, so each row is its minimum and that line.
// With a folder (node test/fixtures/mkrowlines.mjs <dir>) it writes every variant tried in Word there instead.
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
  st = st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, '$1<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr>$2');
  files['word/styles.xml'] = strToU8(st);
  fs.writeFileSync(out ? path.join(out, name) : new URL(`./${name}`, import.meta.url), zipSync(files));
}

const line = (side, val, sz) => (val ? `<w:${side} w:val="${val}" w:sz="${sz}" w:space="0" w:color="000000"/>` : '');
// A cell: its own top and bottom lines ([val, sz] or null), its text (Arial, half-points `hp`) on one line of exact height.
const cell = (w, top, bottom, text, { hp = 15, exact = 161, before = 10 } = {}) => {
  const b = (top || bottom) ? `<w:tcBorders>${top ? line('top', ...top) : ''}${bottom ? line('bottom', ...bottom) : ''}</w:tcBorders>` : '';
  const sp = exact ? `<w:spacing w:before="${before}" w:line="${exact}" w:lineRule="exact"/>` : '';
  const r = `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="${text ? hp : 8}"/>`;
  return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${b}</w:tcPr><w:p><w:pPr>${text ? sp : ''}<w:rPr>${r}</w:rPr></w:pPr>${text ? `<w:r><w:rPr>${r}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>` : ''}</w:p></w:tc>`;
};
const W = [2400, 1600, 1200, 1200];
const tbl = (rows, inside = ['single', 8]) => {
  const all = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => line(s, ...inside)).join('');
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${all}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${W.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`;
};
const p = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
// A header row (no lines of its own: the table's inside line under it), then n rows of `h` twips at least: row 1 with
// `first` lines [top, bottom] in its cells, the rest `rest`; the last column's cells with none of their own, or `lastCol`.
const table = (name, { h = 191, first, rest, n = 12, inside, opts, lastCol = [null, null] }) => {
  const head = `<w:tr><w:trPr><w:trHeight w:val="400"/></w:trPr>${W.map((w, i) => cell(w, null, null, `Head ${i + 1}`, { hp: 20, exact: 0 })).join('')}</w:tr>`;
  const rows = Array.from({ length: n }, (_, k) => {
    const [t, b] = k === 0 ? first : rest;
    return `<w:tr><w:trPr><w:trHeight w:val="${h}"/></w:trPr>${W.map((w, i) => (i === W.length - 1 ? cell(w, ...lastCol, '', opts) : cell(w, t, b, i === 0 ? `${name} row ${k + 1}` : i === 1 ? `${k + 1}-00-${k}` : '', opts))).join('')}</w:tr>`;
  }).join('');
  return p(`Before ${name}`) + tbl(head + rows, inside) + p(`After ${name}`);
};
const dash8 = ['dashSmallGap', 8], dot8 = ['dotted', 8], dot6 = ['dotted', 6], dash6 = ['dashSmallGap', 6], dash12 = ['dashed', 12];
const small = { hp: 10, exact: 113, before: 17 };
const variants = {
  // 1113 page 55: row 1 a dashed line under it only, the rest dashed over and under; the last column the table's lines
  'r01-dash8-first-bottom': table('A', { first: [null, dash8], rest: [dash8, dash8] }),
  'r02-dot8-first-bottom': table('B', { first: [null, dot8], rest: [dot8, dot8] }),
  // 1113 page 52: rows of 150tw, 5pt text, dotted 0.75pt lines over and under, the table's 0.75pt lines
  'r03-dot6-150': table('C', { h: 150, first: [null, dot6], rest: [dot6, dot6], inside: ['single', 6], opts: small }),
  'r04-dash6-150': table('D', { h: 150, first: [null, dash6], rest: [dash6, dash6], inside: ['single', 6], opts: small }),
  // every cell its own lines, the last column too
  'r05-dash8-allcells': table('E', { first: [null, dash8], rest: [dash8, dash8], lastCol: [dash8, dash8] }),
  'r06-dot8-allcells': table('F', { first: [null, dot8], rest: [dot8, dot8], lastCol: [dot8, dot8] }),
  // a wider dashed line, 1.5pt
  'r07-dash12-allcells': table('G', { first: [null, dash12], rest: [dash12, dash12], lastCol: [dash12, dash12] }),
  // row 1 a dashed top of its own too
  'r08-dash8-first-both': table('H', { first: [dash8, dash8], rest: [dash8, dash8], lastCol: [dash8, dash8] }),
  // bottoms only
  'r09-dash8-bottoms': table('I', { first: [null, dash8], rest: [null, dash8], lastCol: [null, dash8] }),
  // tops only
  'r10-dash8-tops': table('J', { first: [dash8, null], rest: [dash8, null], lastCol: [dash8, null] }),
};
if (out) {
  for (const [n, body] of Object.entries(variants)) write(`${n}.docx`, body);
  console.log(Object.keys(variants).length, 'documents in', out);
} else {
  write('row-dash-lines.docx', variants['r01-dash8-first-bottom'] + variants['r03-dot6-150']);
}
