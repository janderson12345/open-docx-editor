// Builds test/fixtures/sect-after-table.docx: a table whose one row (at least 934.5px) ends 20px into the room of the
// empty next-page section-break paragraph after it (14pt Arial bold, 88tw before, two bookmarks), then "Next section".
// Letter, margins 720tw top, 1080tw bottom. Word keeps that paragraph on page 1 past the margin, so "Next section" is
// on page 2 (its prints of the same paragraph 3, 8, 14 and 20px over; with a word in it, it goes to page 2).
// usage: node test/fixtures/mksectbreak.mjs [out dir]  (default: here)
import '../setup.js';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const OUT = process.argv[2] ? path.resolve(process.argv[2]) : path.dirname(fileURLToPath(import.meta.url));
const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="580" w:bottom="1080" w:left="600" w:header="166" w:footer="898" w:gutter="0"/></w:sectPr>';
const HEAD = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:sz w:val="28"/>';
const rowPx = 984 - 48 - 21.5 + 20;
const tbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="${Math.round(rowPx * 15)}"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr><w:p>${run('Table text')}</w:p></w:tc></w:tr></w:tbl>`;
const brk = `<w:p><w:pPr><w:spacing w:before="88"/><w:rPr>${HEAD}</w:rPr>${sect}</w:pPr><w:bookmarkStart w:id="1" w:name="a"/><w:bookmarkStart w:id="2" w:name="b"/><w:bookmarkEnd w:id="1"/><w:bookmarkEnd w:id="2"/></w:p>`;
const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${tbl}${brk}<w:p>${run('Next section')}</w:p>${sect}</w:body>`));
const st = strFromU8(files['word/styles.xml']);
files['word/styles.xml'] = strToU8(st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, '$1<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/></w:rPr>$2'));
fs.writeFileSync(path.join(OUT, 'sect-after-table.docx'), zipSync(files));
