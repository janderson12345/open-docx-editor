import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { DocxSession } from '../src/ooxml/docx.js';

const fixture = (f) => fs.readFileSync(new URL(`./fixtures/${f}`, import.meta.url));
/** The fixture with its settings part's text changed. */
const withSettings = (f, edit) => {
  const files = unzipSync(fixture(f));
  files['word/settings.xml'] = strToU8(edit(strFromU8(files['word/settings.xml'])));
  return zipSync(files);
};

test('hyphenation zone: followed in Word 2010\'s layout and older (0.25in when not given), none from Word 2013\'s on', () => {
  assert.equal(new DocxSession(fixture('hyphenation-zone.docx')).hyphenationZone, 1440);
  assert.equal(new DocxSession(withSettings('hyphenation-zone.docx', (x) => x.replace(/<w:hyphenationZone[^>]*\/>/, ''))).hyphenationZone, 360);
  assert.equal(new DocxSession(withSettings('hyphenation-zone.docx', (x) => x.replace('w:val="14"', 'w:val="15"'))).hyphenationZone, 0);
  assert.equal(new DocxSession(fixture('hyphenation-free.docx')).hyphenationZone, 0);
});

test('consecutive-hyphen limit: as given, none (0) when not', () => {
  assert.equal(new DocxSession(fixture('hyphenation-limit.docx')).consecutiveHyphenLimit, 1);
  assert.equal(new DocxSession(fixture('hyphenation-free.docx')).consecutiveHyphenLimit, 0);
});

test('the zone and the limit are kept on saving', () => {
  for (const [f, re] of [['hyphenation-zone.docx', /<w:hyphenationZone w:val="1440"\/>/], ['hyphenation-limit.docx', /<w:consecutiveHyphenLimit w:val="1"\/>/]]) {
    const s = new DocxSession(fixture(f));
    assert.match(strFromU8(unzipSync(s.save(s.doc))['word/settings.xml']), re);
  }
});
