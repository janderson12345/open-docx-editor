// Builds test/fixtures/toc-wrap.docx: table-of-contents lines (hanging indent, a dotted right tab at the margin, the
// page number a PAGEREF field) whose titles grow a letter at a time, so some leave no room for the number after them.
// Word then keeps the title on its line and starts the next one with the tab, its dots running up to the number.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const fld = (n) => '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>'
  + `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${n}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const entry = (i) => '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs>'
  + '<w:ind w:left="720" w:hanging="720"/></w:pPr>'
  + `<w:r><w:t>${i + 1}.</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>ACCOUNTING MATTERS; ANNUAL BUDGET; INFORMATION ${'R'.repeat(i)}</w:t></w:r>`
  + `<w:r><w:tab/></w:r>${fld(17 + i)}</w:p>`;

const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
const body = '<w:p><w:r><w:t>Contents</w:t></w:r></w:p><w:bookmarkStart w:id="0" w:name="_Toc1"/><w:bookmarkEnd w:id="0"/>'
  + Array.from({ length: 30 }, (_, i) => entry(i)).join('');
files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
fs.writeFileSync(new URL('./toc-wrap.docx', import.meta.url), zipSync(files));
