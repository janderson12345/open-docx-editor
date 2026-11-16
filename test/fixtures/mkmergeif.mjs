// Builds test/fixtures/mergeif.docx: IF fields as Word writes them (complex fields, the merge fields nested in their
// codes and in their texts, an IF nested in another), each comparison Word's IF has, a header and a footer with merge
// fields. Word's own merge of it with the records in test/mailmerge.test.js gave the texts the test expects.
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
const line = (label, field) => P(R(`${label}: `) + field);
const body = [
  line('gt', F([' IF ', M('Amount'), ' > 100 "big" "small" '], 'small')),
  line('lt', F([' IF ', M('Amount'), ' < 100 "under" "not under" '], 'under')),
  line('ge', F([' IF ', M('Amount'), ' >= 100 yes no '], 'no')),
  line('le', F([' IF ', M('Amount'), ' <= 50 "le" "gt" '], 'le')),
  line('ne', F([' IF ', M('Amount'), ' <> "" "has amount" "no amount" '], 'no amount')),
  line('eqcase', F([' IF ', M('Name'), ' = "ada" "eq" "ne" '], 'ne')),
  line('star', F([' IF ', M('Name'), ' = "A*" "A-name" "other" '], 'other')),
  line('qmark', F([' IF ', M('Name'), ' = "?da" "q" "nq" '], 'nq')),
  line('starcase', F([' IF ', M('Name'), ' = "a*" "a-star" "no a-star" '], 'no a-star')),
  line('strlt', F([' IF ', M('Name'), ' < "B" "before B" "from B" '], 'before B')),
  line('strltcase', F([' IF ', M('Name'), ' < "b" "before b" "from b" '], 'before b')),
  line('quoted', F([' IF "', M('City'), '" = "New York" "NY" "not NY" '], 'not NY')),
  line('bare', F([' IF ', M('City'), ' = "New York" "NY" "not NY" '], 'not NY')),
  line('numstr', F([' IF ', M('Amount'), ' = "100.0" "hundred" "not hundred" '], 'not hundred')),
  line('text', F([' IF ', M('Amount'), ' > 100 "Dear ', M('Name'), ', thanks" "Hello ', M('Name'), '" '], 'Hello')),
  line('nested', F([' IF ', M('Amount'), ' > 100 "big" "', F([' IF ', M('Amount'), ' > 40 "mid" "tiny" '], 'tiny'), '" '], 'tiny')),
  line('nofalse', F([' IF ', M('Amount'), ' > 100 "only big" '], '')),
  line('upper', F([' IF ', M('Amount'), ' > 100 "big" "small" \\* Upper '], 'SMALL')),
  line('rec', F([' IF ', F([' MERGEREC '], '«Rec»'), ' = 2 "second" "not second" '], 'not second')),
];
const files = unzipSync(blankDocxBytes());
const story = (tag, style, text) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${text}</w:p></w:${tag}>`;
files['word/header1.xml'] = strToU8(story('hdr', 'Header', R('Statement for ') + M('Name')));
files['word/footer1.xml'] = strToU8(story('ftr', 'Footer', R('City: ') + M('City') + R(' / ') + F([' IF ', M('Amount'), ' > 100 "big" "small" '], 'small')));
const rels = 'word/_rels/document.xml.rels';
files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdMmH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
  + '<Relationship Id="rIdMmF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>'));
const ct = '[Content_Types].xml';
files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
  + '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'));
files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml'])
  .replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body.join('')}$1`)
  .replace(/(<w:sectPr\b[^>]*>)/, '$1<w:headerReference w:type="default" r:id="rIdMmH"/><w:footerReference w:type="default" r:id="rIdMmF"/>'));
fs.writeFileSync(new URL('./mergeif.docx', import.meta.url), zipSync(files));
