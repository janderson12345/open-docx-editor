// Builds test/fixtures/styleref-level.docx: STYLEREF with a heading level for its style (a caption's chapter number,
// STYLEREF 1 \s) under numbered Heading 1 paragraphs, the saved results wrong. Word's PDF of it (its fields updated as
// it prints): "Table 3-1." and "Plain: Place of Performance | heading 1 by name: Place of Performance | number by
// name: 3"; the level had been taken for a style named "1" (features.mjs).
import '../setup.js';
import fs from 'node:fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const fld = (instr, cached) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${cached}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const t = (s) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
const h1 = (s) => `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr>${t(s)}</w:p>`;
const p = (...runs) => `<w:p>${runs.join('')}</w:p>`;
const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="70"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="70"/></w:num></w:numbering>`;
const body = [
  h1('Background'), p(t('Some text.')),
  h1('Scope'), p(t('More text.')),
  h1('Place of Performance'),
  p(t('Table '), fld('STYLEREF 1 \\s', '9'), t('-'), fld('SEQ Table \\* ARABIC \\s 1', '1'), t('. Caption with the chapter number')),
  p(t('Plain: '), fld('STYLEREF 1', 'x'), t(' | heading 1 by name: '), fld('STYLEREF "Heading 1"', 'y'), t(' | number by name: '), fld('STYLEREF "Heading 1" \\n', 'z')),
].join('');
const files = unzipSync(blankDocxBytes());
files['word/numbering.xml'] = strToU8(numbering);
let doc = strFromU8(files['word/document.xml']);
doc = doc.replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`);
files['word/document.xml'] = strToU8(doc);
fs.writeFileSync(new URL('./styleref-level.docx', import.meta.url), zipSync(files));
