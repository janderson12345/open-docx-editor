// Builds test/fixtures/fields.docx: reference fields with stale saved results, as a document has them after it was
// edited and not updated. A front-matter section numbered in roman numerals holds REF (a bookmark's text and its
// paragraph's number), PAGEREF (its page, and "above"/"below" with \p); a second section numbered from 1 holds three
// numbered chapters over several pages; the header is a running head (STYLEREF "Heading 1") with the page number.
// Word prints it as saved (the header worked out per page, the body's fields as saved) and with its fields updated.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
const run = (t) => `<w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
// A complex field with its saved (stale) result.
const field = (instr, saved) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${esc(instr)} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run(saved)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const para = (inner, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${inner}</w:p>`;
const heading = (text, bm, id) => para(`${bm ? `<w:bookmarkStart w:id="${id}" w:name="${bm}"/>` : ''}${run(text)}${bm ? `<w:bookmarkEnd w:id="${id}"/>` : ''}`, '<w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>');
const LONG = 'The quick brown fox jumps over the lazy dog, and then it runs on to the end of the line so that the paragraph wraps onto a second line. ';
const filler = (n, label) => Array.from({ length: n }, (_, i) => para(run(`${label} paragraph ${i + 1}. ${LONG}`))).join('');
const HDR = '<w:headerReference w:type="default" r:id="rIdHdr1"/>';
const MARGINS = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';

const front = [
  para(run('Report on the test fields'), '<w:pStyle w:val="Title"/>'),
  para(run('The methods are in chapter ') + field('REF chap2 \\n \\h', '9') + run(', "') + field('REF chap2 \\h', 'Old title') + run('", on page ') + field('PAGEREF chap2 \\h', '99') + run('.')),
  para(run('The results are ') + field('PAGEREF chap3 \\p', 'here') + run(', on page ') + field('PAGEREF chap3 \\h', '77') + run(', and the introduction on page ') + field('PAGEREF chap1', '1') + run('.')),
  para(run('A reference to a bookmark that isn\'t there: ') + field('REF nowhere \\h', 'stale')),
  filler(4, 'Front matter'),
  // The front matter's section: page numbers in roman numerals.
  para('', `<w:sectPr>${HDR}${MARGINS}<w:pgNumType w:fmt="lowerRoman" w:start="1"/></w:sectPr>`),
].join('');
const chapters = [
  heading('Introduction', 'chap1', 1), filler(14, 'Introduction'),
  heading('Methods of measurement', 'chap2', 2), filler(16, 'Methods'),
  heading('Results', 'chap3', 3), filler(10, 'Results'),
].join('');
const body = front + chapters + `<w:sectPr>${HDR}${MARGINS}<w:pgNumType w:start="1"/></w:sectPr>`;
const header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}>${para(run('Report — ') + field('STYLEREF "Heading 1"', 'Stale heading') + run(' — page ') + field('PAGE', '1'), '<w:jc w:val="right"/>')}</w:hdr>`;

const files = unzipSync(blankDocxBytes());
files['word/document.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`);
files['word/header1.xml'] = strToU8(header);
files['word/_rels/document.xml.rels'] = strToU8(strFromU8(files['word/_rels/document.xml.rels']).replace('</Relationships>', '<Relationship Id="rIdHdr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'));
files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'));
fs.writeFileSync(new URL('./fields.docx', import.meta.url), zipSync(files));
