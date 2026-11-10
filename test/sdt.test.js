import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { DocxSession } from '../src/ooxml/docx.js';
import { blankDocxBytes } from '../src/ooxml/template.js';

const SDT = '<w:sdtPr><w:tag w:val="t1"/></w:sdtPr>';
const cell = (t) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
const BODY = `<w:sdt>${SDT}<w:sdtContent><w:p><w:r><w:t>In control</w:t></w:r></w:p>`
  + `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell('A')}</w:tr>`
  + `<w:sdt><w:sdtPr><w:tag w:val="rows"/></w:sdtPr><w:sdtContent><w:tr>${cell('B')}</w:tr><w:tr>${cell('C')}</w:tr></w:sdtContent></w:sdt></w:tbl>`
  + '</w:sdtContent></w:sdt><w:p><w:r><w:t>After</w:t></w:r></w:p>';

test('content controls around blocks and table rows are shown and saved', () => {
  const files = unzipSync(blankDocxBytes());
  const main = new DocxSession(blankDocxBytes()).pkg.mainPath;
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${BODY}$1`));
  const s = new DocxSession(zipSync(files));
  assert.equal(s.doc.textContent, 'In controlABCAfter');
  const out = strFromU8(unzipSync(s.save(s.doc))[s.pkg.mainPath]);
  assert.match(out, /<w:sdt><w:sdtPr><w:tag w:val="t1"\/><\/w:sdtPr><w:sdtContent><w:p>.*In control.*<w:tbl>.*<w:sdt><w:sdtPr><w:tag w:val="rows"\/><\/w:sdtPr><w:sdtContent><w:tr>.*>B<.*<w:tr>.*>C<.*<\/w:tr><\/w:sdtContent><\/w:sdt><\/w:tbl><\/w:sdtContent><\/w:sdt><w:p>.*After/);
});

test('content controls around table cells keep their cells and are saved', () => {
  const files = unzipSync(blankDocxBytes());
  const main = new DocxSession(blankDocxBytes()).pkg.mainPath;
  const tbl = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
    + `<w:tr>${cell('Label')}<w:sdt><w:sdtPr><w:tag w:val="cell"/></w:sdtPr><w:sdtContent>${cell('Value')}</w:sdtContent></w:sdt></w:tr></w:tbl>`;
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${tbl}<w:p/>$1`));
  const s = new DocxSession(zipSync(files));
  assert.equal(s.doc.textContent, 'LabelValue');
  const out = strFromU8(unzipSync(s.save(s.doc))[s.pkg.mainPath]);
  assert.match(out, /<w:tr><w:tc>.*>Label<.*<\/w:tc><w:sdt><w:sdtPr><w:tag w:val="cell"\/><\/w:sdtPr><w:sdtContent><w:tc>.*>Value<.*<\/w:tc><\/w:sdtContent><\/w:sdt><\/w:tr>/);
});

test('a repeating section in table rows (controls inside a control) keeps its rows and both levels of control', () => {
  const files = unzipSync(blankDocxBytes());
  const main = new DocxSession(blankDocxBytes()).pkg.mainPath;
  const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
  const item = (id, t) => `<w:sdt><w:sdtPr><w:id w:val="${id}"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent><w:tr>${cell(t)}${cell(t + ' fee')}</w:tr></w:sdtContent></w:sdt>`;
  const nestedCell = '<w:sdt><w:sdtPr><w:id w:val="8"/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent>' + cell('Inner cell') + '</w:sdtContent></w:sdt></w:sdtContent></w:sdt>';
  const tbl = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell('Service')}${nestedCell}</w:tr>`
    + `<w:sdt><w:sdtPr><w:id w:val="4"/><w15:repeatingSection/></w:sdtPr><w:sdtContent>${item(5, 'Hosting')}${item(6, 'Support')}</w:sdtContent></w:sdt></w:tbl><w:p/>`;
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:document\b/, `<w:document xmlns:w15="${W15}"`).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${tbl}$1`));
  const s = new DocxSession(zipSync(files));
  assert.match(s.doc.textContent, /Inner cell.*Hosting.*Hosting fee.*Support/);
  const out = strFromU8(unzipSync(s.save())[main]);
  assert.match(out, /<w15:repeatingSection\/><\/w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:id w:val="5"\/><w15:repeatingSectionItem\/><\/w:sdtPr><w:sdtContent><w:tr>[\s\S]*?Hosting fee[\s\S]*?<\/w:tr><\/w:sdtContent><\/w:sdt><w:sdt><w:sdtPr><w:id w:val="6"\/>[\s\S]*?Support fee[\s\S]*?<\/w:tr><\/w:sdtContent><\/w:sdt><\/w:sdtContent><\/w:sdt><\/w:tbl>/);
  assert.match(out, /<w:id w:val="8"\/><\/w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:id w:val="9"\/><\/w:sdtPr><w:sdtContent><w:tc>[\s\S]*?Inner cell[\s\S]*?<\/w:tc><\/w:sdtContent><\/w:sdt><\/w:sdtContent><\/w:sdt><\/w:tr>/);
});
