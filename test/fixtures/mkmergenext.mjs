// Builds test/fixtures/mergenext.docx and mergeskip.docx: NEXT, NEXTIF and SKIPIF fields as Word writes them (complex
// fields, the merge fields nested in their codes), with MERGEREC and MERGESEQ beside them and a header with a merge
// field. Word's own merge of them with the records in test/mailmerge.test.js gave the texts the test expects.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const R = (t) => `<w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
const I = (t) => `<w:r><w:instrText xml:space="preserve">${esc(t)}</w:instrText></w:r>`;
const C = (t) => `<w:r><w:fldChar w:fldCharType="${t}"/></w:r>`;
// A field: its code's parts (text, or a nested field's XML), then its shown result.
const F = (parts, shown) => C('begin') + parts.map((p) => (p.startsWith('<') ? p : I(p))).join('') + C('separate') + R(shown) + C('end');
const M = (name) => F([` MERGEFIELD ${name} `], `«${name}»`);
const P = (runs) => `<w:p>${runs}</w:p>`;
const who = (label) => P(R(`${label}: `) + F([' MERGEREC '], '«Rec»') + R(' ') + F([' MERGESEQ '], '«Seq»') + R(' ') + M('Name'));

function build(file, body, header) {
  const files = unzipSync(blankDocxBytes());
  files['word/header1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>${header}</w:p></w:hdr>`);
  const rels = 'word/_rels/document.xml.rels';
  files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdMmH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'));
  const ct = '[Content_Types].xml';
  files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'));
  files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml'])
    .replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body.join('')}$1`)
    .replace(/(<w:sectPr\b[^>]*>)/, '$1<w:headerReference w:type="default" r:id="rIdMmH"/>'));
  fs.writeFileSync(new URL(`./${file}`, import.meta.url), zipSync(files));
}

// Several records to a document: NEXTIF moving on when its comparison is true, NEXT always; an IF after them.
build('mergenext.docx', [
  who('a'),
  P(R('n1: ') + F([' NEXTIF ', M('Group'), ' = "x" '], '') + R('.')),
  who('b'),
  P(R('n2: ') + F([' NEXT '], '') + R('.')),
  who('c'),
  P(R('if: ') + F([' IF ', M('Amount'), ' > 10 "big" "small" '], 'small')),
], R('H: ') + M('Name'));

// Records left out: SKIPIF before and after a NEXTIF (does it leave out the records the document already had?).
build('mergeskip.docx', [
  who('a'),
  P(R('s1: ') + F([' SKIPIF ', M('Amount'), ' < 10 '], '') + R('.')),
  who('b'),
  P(R('n: ') + F([' NEXTIF ', M('Group'), ' = "x" '], '') + R('.')),
  who('c'),
  P(R('s2: ') + F([' SKIPIF ', M('Name'), ' = "E*" '], '') + R('.')),
  who('d'),
], R('H: ') + M('Name'));

// A SKIPIF true after a NEXTIF moved on: where does the next document start?
build('mergeskip2.docx', [
  who('a'),
  P(R('n: ') + F([' NEXTIF ', M('Group'), ' = "x" '], '') + R('.')),
  who('b'),
  P(R('s: ') + F([' SKIPIF ', M('Name'), ' = "B*" '], '') + R('.')),
], R('H: ') + M('Name'));
