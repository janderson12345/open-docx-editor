// Builds test/fixtures/turned-cells.docx: a header row of cells of text turned on its side (w:textDirection btLr), as
// Word sets them (its prints of the same tables, made here and printed by Word): where each turned line starts along
// the cell and where it sits across it, with space before, indents, centring and line spacing.
// With a folder (node test/fixtures/mkturned.mjs <dir>) it writes every variant tried in Word there instead.
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
  st = st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, '$1<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="15"/></w:rPr>$2');
  files['word/styles.xml'] = strToU8(st);
  fs.writeFileSync(out ? path.join(out, name) : new URL(`./${name}`, import.meta.url), zipSync(files));
}

const p = (t, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${t ? `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>` : ''}</w:p>`;
const B = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="8" w:space="0" w:color="000000"/>`).join('');
// Cells of 602tw across, as 1113's: each [paragraphs, turned?].
const table = (cells, h = 1532) => {
  const W = 602;
  const head = cells.map(([ps, turn]) => `<w:tc><w:tcPr><w:tcW w:w="${W}" w:type="dxa"/>${turn ? '<w:textDirection w:val="btLr"/>' : ''}</w:tcPr>${ps}</w:tc>`).join('');
  const row = cells.map((_, k) => `<w:tc><w:tcPr><w:tcW w:w="${W}" w:type="dxa"/></w:tcPr>${p(`R${k + 1}`)}</w:tc>`).join('');
  return p('Before') + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${B}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${cells.map(() => `<w:gridCol w:w="${W}"/>`).join('')}</w:tblGrid><w:tr><w:trPr><w:trHeight w:val="${h}"/></w:trPr>${head}</w:tr><w:tr>${row}</w:tr></w:tbl>` + p('After');
};
const cells = [
  [p('Plain one', ''), true],
  [p('Before 105', '<w:spacing w:before="105"/>'), true],
  [p('Before 105 centred', '<w:spacing w:before="105"/><w:ind w:left="25" w:right="38"/><w:jc w:val="center"/>') + p('(second para)', '<w:spacing w:before="31"/><w:ind w:left="24" w:right="38"/><w:jc w:val="center"/>'), true],
  [p('Category Name (A or B)', '<w:spacing w:before="100" w:line="283" w:lineRule="auto"/><w:ind w:left="474" w:right="129" w:hanging="260"/>'), true],
  [p('Left 335', '<w:spacing w:before="106"/><w:ind w:left="335"/>') + p('(enter A or B)', '<w:spacing w:before="31"/><w:ind w:left="297"/>'), true],
  [p('Two lines that wrap in the cell', '<w:spacing w:before="20"/>'), true],
  [p('Bottom aligned', '<w:spacing w:before="105"/><w:jc w:val="right"/>'), true],
  [p('Not turned'), false],
];
// The pieces of 1113's turned heading apart: a hanging indent, 1.18 lines, a left indent.
const parts = [
  [p('Hang 474 260', '<w:ind w:left="474" w:hanging="260"/>'), true],
  [p('Line 283 auto', '<w:spacing w:line="283" w:lineRule="auto"/>'), true],
  [p('Left 474 only', '<w:ind w:left="474"/>'), true],
  [p('First 214 only', '<w:ind w:firstLine="214"/>'), true],
  [p('Hang and line 283', '<w:spacing w:line="283" w:lineRule="auto"/><w:ind w:left="474" w:hanging="260"/>'), true],
  [p('Right 129 only', '<w:ind w:right="129"/><w:jc w:val="right"/>'), true],
  [p('Plain two'), true],
  [p('Flat'), false],
];
const variants = { 't01-header': table(cells), 't02-short': table(cells, 900), 't03-parts': table(parts) };
if (out) {
  for (const [n, b] of Object.entries(variants)) write(`${n}.docx`, b);
  console.log(Object.keys(variants).length, 'documents in', out);
} else {
  write('turned-cells.docx', variants['t01-header']);
}
