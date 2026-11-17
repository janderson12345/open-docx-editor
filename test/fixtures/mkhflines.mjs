// Builds test/fixtures/hf-lines.docx: a header line of 10pt Calibri with a 12pt run raised 3pt, and a footer line of
// 10pt Times New Roman, a 12pt tab, a 12pt run raised 3pt, another tab and 10pt text (as sample S078's footer): lines
// of mixed sizes in a header and a footer, which Word sizes as in the body (Word's PDFs: 17.111pt and 16.366pt).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const font = (f) => `<w:rFonts w:ascii="${f}" w:hAnsi="${f}"/>`;
const R = (f, sz, t, more = '') => `<w:r><w:rPr>${font(f)}${more}<w:sz w:val="${sz}"/></w:rPr>${t === '\t' ? '<w:tab/>' : `<w:t xml:space="preserve">${t}</w:t>`}</w:r>`;
const P = (...x) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>${x.join('')}</w:p>`;
const up = '<w:position w:val="6"/>';
const T = 'Times New Roman', C = 'Calibri';

const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
files[main] = strToU8(strFromU8(files[main])
  .replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${P(R(T, 20, 'Body text.'))}$1`)
  .replace(/<w:sectPr>/, '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>'));
files['word/header1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}>${P(R(C, 20, 'Header small '), R(C, 24, 'Raised', up), R(C, 20, ' small.'))}</w:hdr>`);
files['word/footer1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${W}>${P(R(T, 20, 'Footer small'), R(T, 24, '\t'), R(T, 24, 'Raised', up), R(T, 24, '\t'), R(T, 20, 'Page'))}</w:ftr>`);
const ct = (part, type) => `<Override PartName="/word/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`;
files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace('</Types>', ct('header1.xml', 'header') + ct('footer1.xml', 'footer') + '</Types>'));
const rel = (id, type, target) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
const rels = 'word/_rels/document.xml.rels';
files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', rel('rIdH', 'header', 'header1.xml') + rel('rIdF', 'footer', 'footer1.xml') + '</Relationships>'));
fs.writeFileSync(new URL('./hf-lines.docx', import.meta.url), zipSync(files));
