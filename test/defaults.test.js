import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { zipSync, strToU8 } from 'fflate';
import { DocxSession } from '../src/ooxml/docx.js';

// A .docx with no styles part (some generators write none): Word sets its text in its own defaults, Aptos 12pt with
// 8pt after and 1.15 lines (Word's PDF of such a file: 33px from one paragraph to the next).
test('a document with no styles part takes Word\'s own defaults', () => {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const bytes = zipSync({
    '[Content_Types].xml': strToU8('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    '_rels/.rels': strToU8('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/document.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`),
  });
  const { defaults } = new DocxSession(bytes).styles;
  assert.deepEqual(defaults, { rPr: { font: 'Aptos', size: 24 }, pPr: { after: 160, line: 276, lineRule: 'auto' } });
});

test('built-in styles by a localised document\'s own ids (Spanish Word\'s "Ttulo1" is Heading 1)', async () => {
  const fs = await import('node:fs');
  const s = new DocxSession(new Uint8Array(fs.readFileSync(new URL('./fixtures/localised-styles.docx', import.meta.url))));
  assert.equal(s.builtInId('Heading1'), 'Ttulo1');
  assert.equal(s.builtInId('TOC1'), 'TDC1');
  assert.equal(s.builtInId('TOCHeading'), 'TtuloTDC');
  assert.equal(s.builtInId('Title'), 'Title');
  // Not added again under the English ids.
  assert.equal(s.addBuiltInStyles(['TOCHeading', 'TOC1', 'TOC2']), false);
});
