// Builds test/fixtures/paged-tables.docx: three bordered tables, each after a page break (an empty paragraph, then one
// holding only a page break), as form-style documents put each section's box on a page of its own. Clicking in one
// box must leave every box on its page.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const P = (t, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${t}</w:p>`;
const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const border = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="12" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
const table = (title) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${border}<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>`
  + `<w:tr><w:trPr><w:trHeight w:val="302" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${P(run(title), '<w:spacing w:after="0"/><w:jc w:val="center"/>')}</w:tc></w:tr>`
  + `<w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${P(run('Details.'), '<w:spacing w:after="0"/>')}</w:tc></w:tr></w:tbl>`;
const pageBreak = P('') + P('<w:r><w:br w:type="page"/></w:r>', '<w:spacing w:after="0"/>');

const body = [table('A. SCOPE'), pageBreak, table('B. MILESTONES'), pageBreak, table('C. APPROACH'), P('')].join('');
const files = unzipSync(blankDocxBytes());
files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
fs.writeFileSync(new URL('./paged-tables.docx', import.meta.url), zipSync(files));
