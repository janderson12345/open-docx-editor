// Builds test/fixtures/merge-format.docx: a 3 x 3 table whose first column's cells are shaded, with a thick left border,
// centred bold text and a middle vertical alignment; for merging cells (their formatting kept, as in Word).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const cell = (t, first) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${first ? '<w:tcBorders><w:left w:val="single" w:sz="24" w:space="0" w:color="C00000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="FFE699"/><w:vAlign w:val="center"/>' : ''}</w:tcPr>`
  + `<w:p>${first ? '<w:pPr><w:jc w:val="center"/></w:pPr>' : ''}<w:r>${first ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t>${t}</w:t></w:r></w:p></w:tc>`;
const row = (r) => `<w:tr>${['A', 'B', 'C'].map((c, i) => cell(`${c}${r}`, i === 0)).join('')}</w:tr>`;
const tbl = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
  + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('')
  + '</w:tblBorders><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
  + [1, 2, 3].map(row).join('') + '</w:tbl>';
const files = unzipSync(blankDocxBytes());
files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body><w:p><w:r><w:t>Merging cells.</w:t></w:r></w:p>${tbl}<w:p/>$1`));
fs.writeFileSync(new URL('./merge-format.docx', import.meta.url), zipSync(files));
