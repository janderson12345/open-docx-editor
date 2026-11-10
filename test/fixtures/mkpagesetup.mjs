// Builds test/fixtures/page-setup.docx: two sections, each a Letter page with one-inch margins (the first ended by a
// paragraph's w:sectPr, a new page), for Page Setup applied to one section or the whole document.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>';
const P = (t, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}<w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
const body = P('Section one, first paragraph.') + P('Section one ends here.', sect) + P('Section two.');
files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
fs.writeFileSync(new URL('./page-setup.docx', import.meta.url), zipSync(files));
