// Builds test/fixtures/strict.docx: rich.docx in the Strict form (ISO/IEC 29500 Strict namespaces, its
// relationship types, w:conformance="strict"), with a table half the page wide written as Strict does ("50%").
// Word opens it as a Strict document and draws it as it does the original.
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

const S = 'http://purl.oclc.org/ooxml/';
const T = 'http://schemas.openxmlformats.org/';
const MAP = [
  ['officeDocument/2006/relationships/extended-properties', 'officeDocument/relationships/extendedProperties'],
  ['officeDocument/2006/relationships/custom-properties', 'officeDocument/relationships/customProperties'],
  ['officeDocument/2006/relationships', 'officeDocument/relationships'],
  ['officeDocument/2006/extended-properties', 'officeDocument/extendedProperties'],
  ['officeDocument/2006/docPropsVTypes', 'officeDocument/docPropsVTypes'],
  ['officeDocument/2006/math', 'officeDocument/math'],
  ['wordprocessingml/2006/main', 'wordprocessingml/main'],
  ['drawingml/2006/wordprocessingDrawing', 'drawingml/wordprocessingDrawing'],
  ['drawingml/2006/picture', 'drawingml/picture'],
  ['drawingml/2006/main', 'drawingml/main'],
];
const files = unzipSync(fs.readFileSync(new URL('./rich.docx', import.meta.url)));
for (const [path, bytes] of Object.entries(files)) {
  if (!/\.(xml|rels)$/.test(path) || path === '[Content_Types].xml') continue;
  let t = strFromU8(bytes);
  for (const [tr, s] of MAP) t = t.split(T + tr).join(S + s);
  if (path === 'word/document.xml') {
    t = t.replace(/<w:document\b/, '<w:document w:conformance="strict"');
    const cell = (s) => `<w:tc><w:tcPr><w:tcW w:w="50%" w:type="pct"/></w:tcPr><w:p><w:r><w:t>${s}</w:t></w:r></w:p></w:tc>`;
    t = t.replace('<w:body>', `<w:body><w:tbl><w:tblPr><w:tblW w:w="50%" w:type="pct"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2340"/><w:gridCol w:w="2340"/></w:tblGrid><w:tr>${cell('Half the page wide')}${cell('in Strict percentages')}</w:tr></w:tbl>`);
  }
  files[path] = strToU8(t);
}
fs.writeFileSync(new URL('./strict.docx', import.meta.url), zipSync(files));
