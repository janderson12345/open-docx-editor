// Builds test/fixtures/label-exact.docx: numbered paragraphs on lines of exact height (247tw) and of single spacing, the
// number scaled 110% (w:w) as the text, in a table cell and in the body, as Word sets the number's baseline against the
// text's (its prints of the same paragraphs, made here and printed by Word).
// With a folder (node test/fixtures/mklabelexact.mjs <dir>) it writes every variant tried in Word there instead.
import '../setup.js';
import fs from 'fs';
import path from 'path';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const out = process.argv[2] ? path.resolve(process.argv[2]) : null;
if (out) fs.mkdirSync(out, { recursive: true });
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function write(name, body, lvlRpr) {
  const files = unzipSync(blankDocxBytes());
  const main = 'word/document.xml';
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
  let st = strFromU8(files['word/styles.xml']);
  st = st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, '$1<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr>$2');
  files['word/styles.xml'] = strToU8(st);
  files['word/numbering.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="486" w:hanging="366"/></w:pPr><w:rPr>${lvlRpr}</w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  fs.writeFileSync(out ? path.join(out, name) : new URL(`./${name}`, import.meta.url), zipSync(files));
}

const item = (t, line, w, pos = 0) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:tabs><w:tab w:val="left" w:pos="483"/></w:tabs>${line ? `<w:spacing w:line="${line}" w:lineRule="exact"/>` : ''}<w:ind w:left="480" w:hanging="359"/><w:rPr><w:color w:val="2F2838"/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val="2F2838"/>${w ? `<w:w w:val="${w}"/>` : ''}${pos ? `<w:position w:val="${pos}"/>` : ''}</w:rPr><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const items = (line, w, pos) => ['Item one text line', 'Second item text', 'Third item text'].map((t) => item(t, line, w, pos)).join('');
const cell = (inner) => `<w:tbl><w:tblPr><w:tblW w:w="3500" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3500"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3500" w:type="dxa"/></w:tcPr>${inner}</w:tc></w:tr></w:tbl>`;
const p = (t) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
const variants = {
  'l01-exact-w110': [p('Body') + items(247, 110) + p('Cell') + cell(items(247, 110)), '<w:spacing w:val="-1"/><w:w w:val="110"/>'],
  'l02-exact-plain': [p('Body') + items(247, 0) + p('Cell') + cell(items(247, 0)), ''],
  'l03-auto-w110': [p('Body') + items(0, 110) + p('Cell') + cell(items(0, 110)), '<w:spacing w:val="-1"/><w:w w:val="110"/>'],
  // The text raised half a point (w:position 1), the number not: sample 1113's numbered form items.
  'l05-exact-pos1': [p('Body') + items(247, 110, 1) + p('Cell') + cell(items(247, 110, 1)), '<w:spacing w:val="-1"/><w:w w:val="110"/>'],
  'l06-auto-pos1': [p('Body') + items(0, 110, 1) + p('Cell') + cell(items(0, 110, 1)), '<w:spacing w:val="-1"/><w:w w:val="110"/>'],
  'l07-exact-pos2': [p('Body') + items(247, 0, 2) + p('Cell') + cell(items(247, 0, 2)), ''],
  'l08-auto-posm1': [p('Body') + items(0, 0, -1) + p('Cell') + cell(items(0, 0, -1)), ''],
  'l04-exact-w110-plainlabel': [p('Body') + items(247, 110) + p('Cell') + cell(items(247, 110)), ''],
};
if (out) {
  for (const [n, [b, r]] of Object.entries(variants)) write(`${n}.docx`, b, r);
  console.log(Object.keys(variants).length, 'documents in', out);
} else {
  write('label-exact.docx', variants['l05-exact-pos1'][0] + p('End'), variants['l05-exact-pos1'][1]);
}
