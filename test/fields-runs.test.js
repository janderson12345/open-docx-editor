import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { DocxSession } from '../src/ooxml/docx.js';
import { blankDocxBytes } from '../src/index.js';
import { parseField } from '../src/editor/refs.js';

const R = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
function docWith(body) {
  const files = unzipSync(blankDocxBytes());
  files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
  return zipSync(files);
}
const docXml = (bytes) => strFromU8(unzipSync(bytes)['word/document.xml']);

// Some generators write a whole field in one run (begin, code, separator, result and end in the same w:r), and Word
// nests fields (an IF around a PAGE): text and codes must read, and save, as they were.
test('a field written in one run, and nested fields, read and save as they were', () => {
  const oneRun = '<w:p>' + R('Page ') + '<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> PAGE </w:instrText><w:fldChar w:fldCharType="separate"/><w:t>7</w:t><w:fldChar w:fldCharType="end"/></w:r>' + R(' of the report.') + '</w:p>';
  const nested = '<w:p>' + R('Status: ') + '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:instrText xml:space="preserve"> = 1 "first" "later" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>first</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>' + R(' page.') + '</w:p>';
  const s = new DocxSession(docWith(oneRun + nested));
  const paras = [];
  s.doc.forEach((p) => paras.push(p.textContent));
  assert.equal(paras[0], 'Page 7 of the report.');
  assert.equal(paras[1], 'Status: first page.');
  const out = s.save();
  const xml = docXml(out);
  assert.match(xml, /<w:instrText[^>]*> PAGE <\/w:instrText>[\s\S]*?<w:t[^>]*>7<\/w:t>/);
  assert.match(xml, /> IF <\/w:instrText>[\s\S]*> PAGE <[\s\S]*> = 1 "first" "later" <[\s\S]*>first</);
  // Read back: the same document, and the same codes in the same order.
  const again = new DocxSession(out);
  assert.deepEqual(again.doc.toJSON(), s.doc.toJSON());
  const codes = [...docXml(out).matchAll(/<w:instrText[^>]*>([^<]*)</g)].map((m) => m[1].trim());
  assert.deepEqual(codes, ['PAGE', 'IF', 'PAGE', '= 1 "first" "later"']);
  assert.equal(parseField(' PAGE ').name, 'PAGE');
});
