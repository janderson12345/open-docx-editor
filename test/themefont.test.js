import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import { readStyles, resolveStyle, stylesToCss } from '../src/ooxml/styles.js';

// Theme fonts against named fonts, as Word prints them (new test documents: per slot, the nearest level that names
// the font wins, by theme or by name; on one level the theme's wins; docDefaults last).
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const rf = (a) => (a ? `<w:rPr><w:rFonts ${a}/></w:rPr>` : '');
const style = (id, based, a, type = 'paragraph') => `<w:style w:type="${type}" w:styleId="${id}"><w:name w:val="${id}"/>${based ? `<w:basedOn w:val="${based}"/>` : ''}${rf(a)}</w:style>`;
const MAJ = 'w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"';
const theme = { major: 'Cambria', minor: 'Candara' };

const fontOf = (dd, styles, id) => {
  const xml = `<w:styles ${W}><w:docDefaults><w:rPrDefault>${rf(dd)}</w:rPrDefault></w:docDefaults>${styles.join('')}</w:styles>`;
  const info = readStyles({ partByType: () => 'styles.xml', xml: () => new DOMParser().parseFromString(xml, 'text/xml') });
  const r = resolveStyle(info.styles, id).rPr;
  const css = stylesToCss('.d', info, theme);
  const rule = new RegExp(`p\\[data-style="${id}"\\][^{]*\\{([^}]*)\\}`).exec(css)?.[1] || '';
  return { font: r.font ?? null, themeFont: r.themeFont ?? null, css: /font-family:"?([^",;]+)/.exec(rule)?.[1] || null };
};

test('a style\'s theme font replaces the named font of the style it is based on', () => {
  const st = [style('Normal', null, 'w:ascii="Times New Roman" w:hAnsi="Times New Roman"'), style('Heading1', 'Normal', MAJ)];
  const f = fontOf('w:asciiTheme="minorHAnsi"', st, 'Heading1');
  assert.deepEqual([f.font, f.themeFont, f.css], [null, 'majorHAnsi', 'Cambria']);
});

test('a style\'s named font replaces the theme font of the style it is based on', () => {
  const st = [style('Base', null, MAJ), style('S', 'Base', 'w:ascii="Verdana" w:hAnsi="Verdana"')];
  const f = fontOf('', st, 'S');
  assert.deepEqual([f.font, f.themeFont, f.css], ['Verdana', null, 'Verdana']);
});

test('on one level the theme font wins over a named one beside it, in styles and in the defaults', () => {
  assert.equal(fontOf('', [style('S', null, `w:ascii="Arial" w:hAnsi="Arial" ${MAJ}`)], 'S').css, 'Cambria');
  const info = readStyles({ partByType: () => 'styles.xml', xml: () => new DOMParser().parseFromString(`<w:styles ${W}><w:docDefaults><w:rPrDefault>${rf('w:ascii="Tahoma" w:asciiTheme="minorHAnsi"')}</w:rPrDefault></w:docDefaults></w:styles>`, 'text/xml') });
  assert.deepEqual([info.defaults.rPr.font, info.defaults.rPr.themeFont], [null, 'minorHAnsi']);
});

test('a font set only for the other slots (hAnsi, eastAsia, cs) leaves the inherited one', () => {
  const st = [style('Base', null, 'w:ascii="Arial" w:hAnsi="Arial"'), style('S', 'Base', 'w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"'), style('T', 'Base', 'w:hAnsiTheme="majorHAnsi"')];
  assert.equal(fontOf('', st, 'S').font, 'Arial');
  assert.equal(fontOf('', st, 'T').font, 'Arial');
});

test('a character style\'s theme font over its named base', () => {
  const st = [style('CBase', null, 'w:ascii="Verdana"', 'character'), style('C2', 'CBase', MAJ, 'character')];
  const f = fontOf('', st, 'C2');
  assert.deepEqual([f.font, f.themeFont], [null, 'majorHAnsi']);
});
