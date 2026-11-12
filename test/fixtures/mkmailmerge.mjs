// Builds test/fixtures/mailmerge.docx: a letter with merge fields as Word writes them (complex fields whose results
// show «Name»), with switches (\* Upper, \b, \# number, \@ date), a simple field (w:fldSimple), MERGEREC, a bold field,
// and fields in a table cell.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const R = (t, rpr = '') => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
const fld = (instr, shown, rpr = '') => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:fldChar w:fldCharType="begin"/></w:r>`
  + `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:instrText xml:space="preserve"> ${esc(instr)} </w:instrText></w:r>`
  + `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:fldChar w:fldCharType="separate"/></w:r>${R(shown, rpr)}`
  + `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:fldChar w:fldCharType="end"/></w:r>`;
const P = (runs) => `<w:p>${runs}</w:p>`;
const body = [
  P(fld('MERGEFIELD FirstName', '«FirstName»') + R(' ') + fld('MERGEFIELD "Last Name" \\* Upper', '«Last Name»', '<w:b/>')),
  P(`<w:fldSimple w:instr=" MERGEFIELD City "><w:r><w:rPr><w:i/></w:rPr><w:t>«City»</w:t></w:r></w:fldSimple>` + R(', record ') + fld('MERGEREC', '«Rec»')),
  P(fld('MERGEFIELD Title \\b "Dear " \\f ","', '«Title»') + R(' end of greeting')),
  P(R('Balance due: ') + fld('MERGEFIELD Amount \\# "$#,##0.00"', '«Amount»') + R(' by ') + fld('MERGEFIELD Due \\@ "MMMM d, yyyy"', '«Due»') + R('.')),
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
    + P(R('Account: ') + fld('MERGEFIELD Account', '«Account»')) + '</w:tc></w:tr></w:tbl>',
  P(R('Yours sincerely.')),
].join('');
const files = unzipSync(blankDocxBytes());
files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
fs.writeFileSync(new URL('./mailmerge.docx', import.meta.url), zipSync(files));
