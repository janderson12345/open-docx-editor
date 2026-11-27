// Builds test/fixtures/table-style-indent.docx: a table style's paragraph indent (0.25in hanging 0.25in), its own and
// from the table style it is based on, on cells' paragraphs in Normal, in a style based on Normal with no indent of
// its own, in a style setting its own, and in Normal with a direct left indent. Word's PDF of it: the hanging indent in
// Normal and in the style based on it, not in the one with its own, and with the direct 0.5in left indent; none had
// been applied (features.mjs).
import '../setup.js';
import fs from 'node:fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const TEXT = 'Alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november oscar papa quebec romeo.';
const t = (s) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
const cellP = (label, ppr = '') => `<w:p><w:pPr>${ppr}</w:pPr>${t(`${label}: ${TEXT}`)}</w:p>`;
const row = (p) => `<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${p}</w:tc><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p>${t('x')}</w:p></w:tc></w:tr>`;
const table = (style) => `<w:tbl><w:tblPr><w:tblStyle w:val="${style}"/><w:tblW w:w="4800" w:type="dxa"/><w:tblLook w:val="0000"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>`
  + row(cellP('Normal'))
  + row(cellP('Body', '<w:pStyle w:val="BodyTx"/>'))
  + row(cellP('Own', '<w:pStyle w:val="OwnInd"/>'))
  + row(cellP('Direct', '<w:ind w:left="720"/>'))
  + '</w:tbl>';
const bdr = '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>';
const styles = '<w:style w:type="table" w:styleId="TsBase"><w:name w:val="Ts Base"/><w:basedOn w:val="TableNormal"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr><w:tblPr>' + bdr + '</w:tblPr></w:style>'
  + '<w:style w:type="table" w:styleId="TsChild"><w:name w:val="Ts Child"/><w:basedOn w:val="TsBase"/><w:pPr><w:spacing w:before="0"/></w:pPr><w:tblPr>' + bdr + '</w:tblPr></w:style>'
  + '<w:style w:type="table" w:styleId="TsOwn"><w:name w:val="Ts Own"/><w:basedOn w:val="TableNormal"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr><w:tblPr>' + bdr + '</w:tblPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="BodyTx"><w:name w:val="Body Tx"/><w:basedOn w:val="Normal"/></w:style>'
  + '<w:style w:type="paragraph" w:styleId="OwnInd"><w:name w:val="Own Ind"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:style>';
const files = unzipSync(blankDocxBytes());
let st = strFromU8(files['word/styles.xml']);
if (!/w:styleId="TableNormal"/.test(st)) throw new Error('no TableNormal');
st = st.replace('</w:styles>', `${styles}</w:styles>`);
files['word/styles.xml'] = strToU8(st);
let doc = strFromU8(files['word/document.xml']);
const body = `<w:p>${t('Based on a style with the indent:')}</w:p>${table('TsChild')}<w:p>${t('Its own indent:')}</w:p>${table('TsOwn')}<w:p>${t('End.')}</w:p>`;
doc = doc.replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`);
files['word/document.xml'] = strToU8(doc);
fs.writeFileSync(new URL('./table-style-indent.docx', import.meta.url), zipSync(files));
