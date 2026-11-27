import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import { fontAlts } from '../src/ooxml/fonts.js';

// A package with only a font table, as fontAlts reads it.
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const sig = (usb0) => `<w:sig w:usb0="${usb0}" w:usb1="500079DB" w:usb2="00000010" w:usb3="00000000" w:csb0="00000001" w:csb1="00000000"/>`;
const pkgOf = (fonts) => {
  const doc = new DOMParser().parseFromString(`<w:fonts ${W}>${fonts}</w:fonts>`, 'application/xml');
  return { rels: [{ type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable', target: 'word/fontTable.xml' }], xml: () => doc };
};
const entry = (name, inner) => `<w:font w:name="${name}"><w:panose1 w:val="02000503000000020004"/><w:charset w:val="00"/><w:family w:val="auto"/><w:pitch w:val="variable"/>${inner}</w:font>`;

// Word's PDFs of generated documents, one missing font each (only its font table entry changing between them).
test('a missing font whose entry claims Armenian or Georgian letters is set in Sylfaen, as Word prints it', () => {
  const alts = fontAlts(pkgOf([
    entry('Helvetica Neue', sig('E50002FF')), // (the Mac's Helvetica Neue: Sylfaen in Word's PDF, not Arial)
    entry('Font Georgian', sig('04000003')), // (Georgian, bit 26: Sylfaen)
    entry('Font Armenian', sig('00000403')), // (Armenian, bit 10: Sylfaen)
    entry('Font Latin', sig('00000003')), // (Latin alone: not Sylfaen)
    entry('Font Unsigned', ''), // (no w:sig: not Sylfaen)
    entry('Font Named', `<w:altName w:val="Arial"/>${sig('E50002FF')}`), // (the file's own stand-in wins: Arial)
    entry('Sylfaen', sig('04000687')), // (Sylfaen's own entry: no stand-in for itself)
  ].join('')));
  assert.equal(alts.get('Sylfaen'), undefined);
  assert.equal(alts.get('Helvetica Neue'), 'Sylfaen');
  assert.equal(alts.get('Font Georgian'), 'Sylfaen');
  assert.equal(alts.get('Font Armenian'), 'Sylfaen');
  assert.equal(alts.get('Font Latin'), undefined);
  assert.equal(alts.get('Font Unsigned'), undefined);
  assert.equal(alts.get('Font Named'), 'Arial');
});

test('an altName that is no name (a converter\'s U+FDFD run) is passed over, as Word passes it', () => {
  const alts = fontAlts(pkgOf([
    entry('Palatino', `<w:altName w:val="${'\ufdfd'.repeat(8)}iqua"/>${sig('A00002FF')}`),
    entry('Font Kept', `<w:altName w:val="Book Antiqua"/>${sig('A00002FF')}`),
  ].join('')));
  assert.equal(alts.get('Palatino'), undefined);
  assert.equal(alts.get('Font Kept'), 'Book Antiqua');
});

test('a missing font Word knows nothing else of falls back by its entry: roman to Cambria, others to Calibri', async () => {
  const { fontFallbacks } = await import('../src/ooxml/fonts.js');
  const { fontStack, setFontFallbacks, lineFactor } = await import('../src/ooxml/props.js');
  const fam = (name, family, charset = '00') => `<w:font w:name="${name}"><w:charset w:val="${charset}"/><w:family w:val="${family}"/><w:pitch w:val="variable"/>${sig('00000003')}</w:font>`;
  const fb = fontFallbacks(pkgOf([fam('Qzv Swiss', 'swiss'), fam('Qzv Roman', 'roman'), fam('Qzv Script', 'script'), fam('Qzv Mac', 'auto', '4D'),
    fam('Qzv Hebrew', 'script', 'B1'), fam('Qzv Hebrew Roman', 'roman', 'B1'), fam('Qzv Jis', 'auto', '80')].join('')));
  assert.deepEqual([...fb], [['Qzv Swiss', 'Calibri'], ['Qzv Roman', 'Cambria'], ['Qzv Script', 'Calibri'], ['Qzv Mac', 'Calibri'],
    ['Qzv Hebrew', 'Arial'], ['Qzv Hebrew Roman', 'Times New Roman'], ['Qzv Jis', null]]);
  setFontFallbacks(fb);
  assert.match(fontStack('Qzv Swiss'), /^"Qzv Swiss", "Calibri"/);
  assert.equal(lineFactor('Qzv Swiss'), lineFactor('Calibri'));
  assert.match(fontStack('Qzv Roman'), /^"Qzv Roman", "Cambria"/);
  assert.match(fontStack('Qzv Jis'), /^"Qzv Jis", sans-serif$/);
  assert.match(fontStack('Qzv Unlisted'), /^"Qzv Unlisted", "Cambria"/); // (no entry)
  // (A name Word knows beats the entry's family: Lucida Grande with a swiss entry is still Segoe UI.)
  setFontFallbacks(new Map([['Lucida Grande', 'Calibri']]));
  assert.match(fontStack('Lucida Grande'), /^"Lucida Grande", "Segoe UI"/);
});
