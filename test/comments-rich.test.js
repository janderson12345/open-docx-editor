import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { DocxSession } from '../src/ooxml/docx.js';
import { blankDocxBytes } from '../src/index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
const R = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
// Comments as Word can write them: one with a table, one with a picture, one with a bookmark pair, one plain.
function withComments() {
  const files = unzipSync(blankDocxBytes());
  const anchor = (id) => `<w:commentRangeStart w:id="${id}"/>${R(`Text ${id}.`)}<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;
  files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body><w:p>${[0, 1, 2, 3].map(anchor).join('')}</w:p>$1`));
  const c = (id, inner) => `<w:comment w:id="${id}" w:author="A" w:date="2026-11-11T00:00:00Z" w:initials="A">${inner}</w:comment>`;
  const table = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p>${R('Cell one')}</w:p></w:tc><w:tc><w:p>${R('Cell two')}</w:p></w:tc></w:tr></w:tbl>`;
  const picture = `<w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="P"/></wp:inline></w:drawing></w:r>`;
  files['word/comments.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W}>${c(0, `<w:p>${R('Before the table')}</w:p>${table}<w:p/>`)}${c(1, `<w:p>${R('A picture: ')}${picture}</w:p>`)}${c(2, `<w:p><w:bookmarkStart w:id="9" w:name="b"/>${R('Marked words')}<w:bookmarkEnd w:id="9"/></w:p>`)}${c(3, `<w:p>${R('Plain')}</w:p>`)}</w:comments>`);
  files['word/_rels/document.xml.rels'] = strToU8(strFromU8(files['word/_rels/document.xml.rels']).replace('</Relationships>', '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'));
  files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'));
  return zipSync(files);
}

test('comments with a table, a picture or a bookmark: read as text, and the rich ones kept as they are', () => {
  const s = new DocxSession(withComments());
  const byId = (id) => s.comments.find((c) => c.id === String(id));
  assert.equal(byId(0).text, 'Before the table\nCell one\tCell two\n');
  assert.ok(byId(0).rich && byId(1).rich && !byId(2).rich && !byId(3).rich);
  assert.equal(byId(1).text, 'A picture: ');
  assert.equal(byId(2).text, 'Marked words');
  // Another comment edited: the table and the picture are still there when saved.
  Object.assign(byId(3), { text: 'Plain, edited', xml: null });
  s.commentsDirty = true;
  const xml = strFromU8(unzipSync(s.save())['word/comments.xml']);
  assert.match(xml, /<w:tbl>[\s\S]*Cell two/);
  assert.match(xml, /<w:drawing>/);
  assert.match(xml, /Plain, edited/);
});
