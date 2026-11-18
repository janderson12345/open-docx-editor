// Builds test/fixtures/w16du-tracked-table.docx: tracked changes with Word 2023's UTC time (w16du:dateUtc, declared on
// the document's root, as Word writes it) inside a table's and a cell's property changes (w:tblPrChange,
// w:tcPrChange), a paragraph's (w:pPrChange) and a run's (w:rPrChange). Those properties are kept as XML and read
// again on their own; a browser rejected them while w16du went undeclared there, and the document didn't open.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const W16DU = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const ch = (id) => `w:id="${id}" w:author="Reviewer" w:date="2026-11-18T09:00:00Z" w16du:dateUtc="2026-11-18T14:00:00Z"`;
const cell = (t, id) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:tcPrChange ${ch(id)}><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
const body = '<w:p><w:r><w:t>A table whose formatting was changed with tracking on:</w:t></w:r></w:p>'
  + `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders><w:tblPrChange ${ch(1)}><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr></w:tblPrChange></w:tblPr>`
  + `<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cell('One', 2)}${cell('Two', 3)}</w:tr></w:tbl>`
  + `<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange ${ch(4)}><w:pPr/></w:pPrChange></w:pPr><w:r><w:rPr><w:b/><w:rPrChange ${ch(5)}><w:rPr/></w:rPrChange></w:rPr><w:t>Centred and bold, both tracked.</w:t></w:r></w:p>`
  + `<w:p><w:ins ${ch(6)}><w:r><w:t>An inserted sentence.</w:t></w:r></w:ins></w:p>`;
const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
let xml = strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`);
xml = xml.replace(/<w:document\b/, `<w:document xmlns:w16du="${W16DU}"${/xmlns:mc=/.test(xml) ? '' : ` xmlns:mc="${MC}"`}`);
files[main] = strToU8(xml);
fs.writeFileSync(new URL('./w16du-tracked-table.docx', import.meta.url), zipSync(files));
