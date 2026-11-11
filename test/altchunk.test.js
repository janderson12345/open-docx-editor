import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { EditorState } from 'prosemirror-state';
import { DocxSession } from '../src/ooxml/docx.js';
import { chunkToWml, chunkKind } from '../src/ooxml/altchunk.js';
import { parseHtml } from '../src/ooxml/xml.js';

const fixture = () => fs.readFileSync(new URL('./fixtures/altchunk.docx', import.meta.url));
const env = { parseHtml, list: (k) => (k === 'bullet' ? '7' : '8'), link: () => 'rIdL' };
const docXml = (bytes) => strFromU8(unzipSync(new Uint8Array(bytes))['word/document.xml']);
const texts = (node) => { const out = []; node.descendants((n) => { if (n.isTextblock) out.push(n.textContent); return true; }); return out; };

test('imported content: the kinds read, and the ones left as they are', () => {
  assert.equal(chunkKind('application/xhtml+xml'), 'html');
  assert.equal(chunkKind('text/html'), 'html');
  assert.equal(chunkKind('text/plain'), 'text');
  assert.equal(chunkKind('application/rtf'), null);
  assert.equal(chunkKind('application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'), null);
});

test('imported content: HTML as the paragraphs Word makes of it', () => {
  // Only white space: one empty paragraph.
  assert.equal(chunkToWml('<html>\n <head/>\n <body>\n   \n </body>\n</html>', 'html', env), '<w:p/>');
  // Text straight in the body: the document's own paragraph formatting, the white space around it gone.
  assert.equal(chunkToWml('<html><head/><body>\n   Basis of Estimate...\n  </body></html>', 'html', env), '<w:p><w:r><w:t xml:space="preserve">Basis of Estimate...</w:t></w:r></w:p>');
  // A <p>: Times New Roman 12pt, automatic spacing; bold and italic runs; a line break.
  const p = chunkToWml('<p>One <b>two</b> <i>three</i><br/>four</p>', 'html', env);
  assert.match(p, /^<w:p><w:pPr><w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1" w:line="240" w:lineRule="auto"\/>/);
  assert.match(p, /<w:rFonts w:ascii="Times New Roman"[^>]*\/><w:b\/><w:bCs\/><w:sz w:val="24"\/>/);
  assert.match(p, /<w:i\/><w:iCs\/>/);
  assert.match(p, /<w:br\/>/);
  assert.equal((p.match(/<w:p>/g) || []).length, 1);
  // Headings, lists (a nested one a level down), a link, inline styles, a table.
  const rich = chunkToWml('<h2>Title</h2><ul><li>a</li><li>b<ul><li>c</li></ul></li></ul><ol><li>x</li></ol><p><a href="https://example.com/">site</a> <span style="color:#C00000;font-size:16pt;font-family:\'Arial\'">red</span></p><table border="1"><tr><td>1</td><th>2</th></tr></table>', 'html', env);
  assert.match(rich, /<w:pStyle w:val="Heading2"\/>/);
  assert.match(rich, /<w:numPr><w:ilvl w:val="0"\/><w:numId w:val="7"\/><\/w:numPr><w:spacing w:before="100" w:beforeAutospacing="1" w:after="0" w:line="240" w:lineRule="auto"\/>[\s\S]*>a</);
  assert.match(rich, /<w:numPr><w:ilvl w:val="1"\/><w:numId w:val="7"\/><\/w:numPr><w:spacing w:before="0" w:after="100" w:afterAutospacing="1" w:line="240" w:lineRule="auto"\/>[\s\S]*>c</);
  assert.match(rich, /<w:numId w:val="8"\/>/);
  assert.match(rich, /<w:hyperlink r:id="rIdL" w:history="1">/);
  assert.match(rich, /<w:rFonts w:ascii="Arial"[^>]*\/><w:color w:val="C00000"\/><w:sz w:val="32"\/>/);
  assert.match(rich, /<w:tbl>.*<w:tblBorders>.*<w:tc>.*>1<.*<w:jc w:val="center"\/>.*<w:b\/>.*>2</);
  // Plain text: a paragraph a line.
  assert.equal((chunkToWml('one\r\ntwo\n', 'text', env).match(/<w:p>/g) || []).length, 2);
});

test('imported content: shown as paragraphs, saved back as it was until edited', () => {
  const s = new DocxSession(fixture());
  const all = texts(s.doc);
  assert.ok(!JSON.stringify(s.doc.toJSON()).includes('opaque_block'), 'no chunk left as a placeholder');
  for (const t of ['Basis of Estimate: hours from the last three task orders.', 'First line of the field, bold and italic.', 'Heading one', 'First bullet', 'Nested bullet', 'Cell four', 'Centred paragraph.', 'From an .htm part & an entity: café.', 'Plain text part, first line.', 'Second line.']) {
    assert.ok(all.some((x) => x.includes(t)), `shown: ${t}`);
  }
  // The rich field's line break and the lists' numbering.
  let bullets = 0;
  s.doc.descendants((n) => { if (n.type.name === 'paragraph' && n.attrs.numId != null && n.attrs.chunk) bullets++; });
  assert.equal(bullets, 5);
  // Saved with no edit: every w:altChunk written back as it was, and its part kept.
  const saved = s.save();
  const xml = docXml(saved);
  assert.equal((xml.match(/<w:altChunk\b/g) || []).length, 6);
  assert.match(xml, /<w:altChunk r:id="chunkPlain"><w:altChunkPr><w:matchSrc w:val="false"\/><\/w:altChunkPr><\/w:altChunk>/);
  assert.ok(!xml.includes('Basis of Estimate'));
  assert.ok(unzipSync(new Uint8Array(saved))['word/afchunk2.dat']);
  // Opened again, it reads the same.
  assert.deepEqual(texts(new DocxSession(saved).doc), all);
});

test('imported content: an edited chunk is saved as the paragraphs it now is', () => {
  const s = new DocxSession(fixture());
  let pos = null;
  s.doc.descendants((n, p) => { if (pos == null && n.isText && n.text.startsWith('Basis of Estimate')) pos = p; });
  const doc = EditorState.create({ doc: s.doc }).tr.insertText('Revised. ', pos).doc;
  const xml = docXml(s.save(doc));
  assert.ok(!/<w:altChunk r:id="chunkPlain"/.test(xml), 'the edited one is no longer a chunk');
  assert.match(xml, /<w:p><w:r><w:t xml:space="preserve">Revised\. Basis of Estimate: hours from the last three task orders\.<\/w:t><\/w:r><\/w:p>/);
  assert.equal((xml.match(/<w:altChunk\b/g) || []).length, 5, 'the others still are');
  // Unedited, its lists aren't written; edited, they are, with the numbering they're shown with.
  const num = (bytes) => strFromU8(unzipSync(new Uint8Array(bytes))['word/numbering.xml']);
  const before = strFromU8(unzipSync(new Uint8Array(fixture()))['word/numbering.xml']);
  assert.equal(num(s.save(doc)), before);
  let at = null;
  s.doc.descendants((n, p) => { if (at == null && n.isText && n.text === 'First bullet') at = p; });
  const listed = EditorState.create({ doc: s.doc }).tr.insertText('Very ', at).doc;
  const out = s.save(listed);
  const ids = [...docXml(out).matchAll(/<w:numId w:val="([^"]+)"/g)].map((m) => m[1]);
  assert.ok(ids.length >= 5);
  for (const id of new Set(ids)) assert.match(num(out), new RegExp(`<w:num w:numId="${id}">`), `list ${id} defined`);
});
