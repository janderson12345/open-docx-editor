// Builds test/fixtures/table-left-compat12.docx: tables indented from the margin, no cell margins at the sides, left
// lines of 0.5 to 6pt, in a document of Word 2007's compatibility (w:compatibilityMode 12), as Word places them (its prints
// of the same documents, made here and printed by Word): text where the indent and paragraph indent put it, the line
// outside, not half a line further in as in a Word 2013 document.
// With a folder (node test/fixtures/mktblcompat.mjs <dir>) it writes every variant tried in Word there instead.
import '../setup.js';
import fs from 'fs';
import path from 'path';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const out = process.argv[2] ? path.resolve(process.argv[2]) : null;
if (out) fs.mkdirSync(out, { recursive: true });

function write(name, body, mode) {
  const files = unzipSync(blankDocxBytes());
  const main = 'word/document.xml';
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
  let st = strFromU8(files['word/styles.xml']);
  st = st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, '$1<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr>$2');
  files['word/styles.xml'] = strToU8(st);
  // (A settings part of its own: the template has none, and a document without one opens in Word 2007's compatibility.)
  files['word/settings.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/></w:compat></w:settings>`);
  const ct = '[Content_Types].xml', rels = 'word/_rels/document.xml.rels';
  files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'));
  files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'));
  fs.writeFileSync(out ? path.join(out, name) : new URL(`./${name}`, import.meta.url), zipSync(files));
}

const p = (t, ind = 0) => `<w:p><w:pPr>${ind ? `<w:ind w:left="${ind}"/>` : ''}</w:pPr><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
// A table `ind` twips in, its lines `sz` eighths of a point, its cells' side margins `mar` twips; two columns.
const tbl = (label, ind, sz, mar) => {
  const b = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((k) => `<w:${k} w:val="single" w:sz="${sz}" w:space="0" w:color="000000"/>`).join('');
  const cell = (t) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t, 100)}</w:tc>`;
  const right = (t) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
  return p(`Before ${label}`) + `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblInd w:w="${ind}" w:type="dxa"/><w:tblBorders>${b}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="${mar}" w:type="dxa"/><w:right w:w="${mar}" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell(`${label} one`)}${right(`${label} two`)}</w:tr></w:tbl>`;
};
const set = tbl('A', 341, 18, 0) + tbl('B', 0, 18, 0) + tbl('C', 341, 4, 0) + tbl('D', 341, 48, 0) + tbl('E', 341, 18, 108) + p('After');
const variants = { 'c12-tables': [set, 12], 'c14-tables': [set, 14], 'c15-tables': [set, 15] };
if (out) {
  for (const [n, [b, m]] of Object.entries(variants)) write(`${n}.docx`, b, m);
  console.log(Object.keys(variants).length, 'documents in', out);
} else {
  write('table-left-compat12.docx', set, 12);
}
