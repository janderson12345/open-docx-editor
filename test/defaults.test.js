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

// "Automatically update document styles" (w:linkStyles) with no template named: Word takes Normal from its own Normal
// template on opening (its prints of generated documents: a Calibri 11pt Normal set at 12pt, 8pt after, 1.16 lines;
// a style of the document's own, and its defaults, unchanged; nothing changed when the template named isn't there).
test('w:linkStyles with no template named takes Normal from Word\'s own Normal template; a named template changes nothing', () => {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const make = (settings) => zipSync({
    '[Content_Types].xml': strToU8('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'),
    '_rels/.rels': strToU8('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/_rels/document.xml.rels': strToU8('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'),
    'word/document.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`),
    'word/styles.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/></w:rPr></w:style><w:style w:type="paragraph" w:customStyle="1" w:styleId="Own"><w:name w:val="Own"/><w:rPr><w:i/></w:rPr></w:style></w:styles>`),
    'word/settings.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><w:settings ${W} ${R}>${settings}</w:settings>`),
  });
  const linked = new DocxSession(make('<w:linkStyles/>')).styles;
  assert.deepEqual(linked.styles.get('Normal').rPr, { themeFont: 'minorHAnsi', themeFontCs: 'minorBidi', size: 24, sizeCs: 24, kern: 2 });
  assert.deepEqual(linked.styles.get('Normal').pPr, { after: 160, line: 278, lineRule: 'auto' });
  assert.deepEqual(linked.styles.get('Own').rPr, { italic: true });
  assert.equal(linked.defaults.rPr.font, 'Times New Roman');
  for (const s of ['', '<w:linkStyles w:val="0"/>', '<w:attachedTemplate r:id="rId1"/><w:linkStyles/>']) {
    assert.equal(new DocxSession(make(s)).styles.styles.get('Normal').rPr.size, 22, s || 'no linkStyles');
  }
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
