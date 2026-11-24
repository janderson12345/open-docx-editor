import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { legacyInfo } from '../src/ooxml/ffdata.js';
import { bevelFx, textEffectsCss } from '../src/ooxml/texteffects.js';
import { readBorders, borderCss, cleanBorders, pPrToCss } from '../src/ooxml/props.js';
import { parseXml, esc } from '../src/ooxml/xml.js';
import { DocxSession } from '../src/ooxml/docx.js';
import { boundPr, boundDisplay, showsPlc, dropPlc, storeValue, formatDate } from '../src/ooxml/databind.js';
import { listFields } from '../src/editor/fields.js';
import { tablePlacementCss } from '../src/ooxml/styles.js';
import { bindingDocx, bound } from './fixtures/mkbinding.mjs';

// Crafted values from a file, bounded (the security review of the leftovers after 0.8.0).

test('a form field read in one pass however many unclosed tags a comment holds', () => {
  const junk = '<w:maxLength '.repeat(40000);
  const xml = `<w:ffData><w:name w:val="F"/><!-- ${junk} --><w:textInput><w:maxLength w:val="5"/></w:textInput></w:ffData>`;
  const t0 = Date.now();
  const info = legacyInfo(xml);
  assert.ok(Date.now() - t0 < 1000);
  assert.equal(info.kind, 'text');
  assert.equal(legacyInfo('<w:ffData><w:textInput><w:maxLength w:val="7"/></w:textInput></w:ffData>').maxLength, 7);
  assert.equal(legacyInfo('<w:ffData><w:textInput/></w:ffData>').maxLength, null);
});

// (The security review of 0.9.0: a cell's left line found in one pass, CodeQL js/polynomial-redos.)
test("a table's first cell's left line read in one pass however many unclosed tags its borders hold", () => {
  const tblPr = '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/></w:tblPr>';
  const t0 = Date.now();
  assert.equal(tablePlacementCss({}, tblPr, 15, parseXml, null, 0, '<w:left '.repeat(40000)), '');
  assert.ok(Date.now() - t0 < 1000);
  assert.equal(tablePlacementCss({}, tblPr, 15, parseXml, null, 0, '<w:tcBorders><w:left w:val="single" w:sz="24"/></w:tcBorders>'), 'margin-left:2px');
});

test("a bevel's size kept to 64px however large the file says", () => {
  const big = '<w14:props3d w14:prstMaterial="matte"><w14:bevelT w14:w="999999999999" w14:h="999999999999" w14:prst="angle"/></w14:props3d>';
  const out = bevelFx(big);
  assert.ok(out);
  for (const m of out.matchAll(/(?:radius|stdDeviation|surfaceScale)="([\d.\s]+)"/g)) {
    for (const n of m[1].trim().split(/\s+/)) assert.ok(Number(n) <= 64, `${m[0]} past 64`);
  }
});

test("a border's size and space bounded as Word bounds them", () => {
  const doc = parseXml('<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:top w:val="triple" w:sz="100000000" w:space="99999" w:color="000000"/><w:bottom w:val="single" w:sz="12" w:space="1"/></w:pBdr>');
  const b = readBorders(doc.documentElement);
  assert.equal(b.top.sz, 255);
  assert.equal(b.top.space, 31);
  assert.equal(b.bottom.sz, 12);
  assert.ok(parseFloat(borderCss(b.top)) <= 5 * 255 / 8);
});

test("a co-editor's borders of any shape drawn safely: numbers as numbers, colours as hex", () => {
  const css = pPrToCss({ borders: { top: { val: 'double', sz: 2, space: '1', color: 'red;background:url(x)' }, left: { val: 'single', sz: '8', space: 'x' }, constructor: { val: 'single' } } });
  assert.doesNotMatch(css, /url\(x\)|red;/);
  assert.doesNotMatch(css, /NaN|undefined/);
  assert.deepEqual(cleanBorders({ top: { val: 'single', sz: '1e9', space: -4, color: 'FF0000' } }).top, { val: 'single', sz: 255, color: 'FF0000', space: 0 });
});

// The security review of 0.8.3: content controls bound to custom XML.

test('thousands of controls bound past a part of many siblings open in bounded time', () => {
  const root = `<r>${'<a/>'.repeat(200000)}<z>end</z></r>`;
  const body = Array.from({ length: 300 }, (_, i) => `<w:p>${bound({ xpath: i % 2 ? '/r/a[999999]' : '/r/z', store: '{00000000-0000-0000-0000-000000000000}' })}</w:p>`).join('');
  const t0 = Date.now();
  const s = new DocxSession(bindingDocx({ body, parts: [{ xml: root }, { xml: root }, { xml: root }] }));
  assert.ok(Date.now() - t0 < 8000, `${Date.now() - t0}ms`);
  assert.equal(listFields(s.doc).filter((f) => f.text === 'end').length, 150);
});

test("a co-editor's control properties of unclosed tags read in one pass", () => {
  const pr = '<w:sdtPr><w:dataBinding w:xpath="/r"/>' + '<w:showingPlcHdr'.repeat(40000) + '<w:storeMappedDataAs w:val="'.repeat(10000) + '<w:text w:multiLine="'.repeat(10000);
  const t0 = Date.now();
  showsPlc(pr);
  dropPlc(pr);
  boundPr(pr, { text: 'x', checked: true, date: new Date(2020, 0, 1) });
  boundDisplay(pr, 'text', 'v');
  boundDisplay(pr, 'checkbox', 'true');
  storeValue(pr, 'date', { text: '1/1/2020' });
  formatDate(new Date(2020, 0, 1), `<w:dateFormat w:val="${"'".repeat(100000)}"/>`);
  assert.ok(Date.now() - t0 < 1000, `${Date.now() - t0}ms`);
  // Still what the patterns did.
  assert.equal(showsPlc('<w:sdtPr><w:showingPlcHdr/></w:sdtPr>'), true);
  assert.equal(showsPlc('<w:sdtPr><w:showingPlcHdr w:val="0"/></w:sdtPr>'), false);
  assert.equal(dropPlc('<w:sdtPr><w:alias w:val="a"/><w:showingPlcHdr/><w:tag w:val="t"/></w:sdtPr>'), '<w:sdtPr><w:alias w:val="a"/><w:tag w:val="t"/></w:sdtPr>');
  assert.match(boundPr('<w:sdtPr><w:date w:fullDate="x"><w:dateFormat w:val="M/d/yyyy"/></w:date></w:sdtPr>', { date: new Date(2020, 0, 2) }), /<w:date w:fullDate="2020-01-02T00:00:00Z"><w:dateFormat/);
});

test("characters XML can't hold are left out of what is saved", () => {
  assert.equal(esc('a\u0001b\uFFFEc\uD800d\te\n'), 'abcd\te\n');
  assert.equal(esc('😀'), '😀');
});

test('a bound control nested 100,000 deep opens', () => {
  const deep = '<w:r><w:t>x</w:t></w:r>';
  const content = '<w:customXml w:element="e">'.repeat(100000) + deep + '</w:customXml>'.repeat(100000);
  let s;
  try { s = new DocxSession(bindingDocx({ body: `<w:p>${bound({ xpath: '/r/a', plc: false, content })}</w:p>`, parts: [{ xml: '<r><a>val</a></r>' }] })); } catch (e) {
    // (A parser may refuse such nesting itself; what must not happen is our own walk overflowing.)
    assert.doesNotMatch(String(e), /call stack/i);
    return;
  }
  assert.ok(s.doc);
});

test("a reflection's blur bounded", () => {
  const x = '<w14:reflection w14:blurRad="999999999999" w14:stA="55000" w14:stPos="0" w14:endA="300" w14:endPos="45500" w14:dist="0" w14:dir="5400000" w14:fadeDir="5400000" w14:sx="100000" w14:sy="-100000" w14:kx="0" w14:ky="0" w14:algn="bl"/>';
  const css = textEffectsCss([['reflection', x]]).join(';');
  const m = /--odx-rf-blur:([\d.]+)pt/.exec(css);
  assert.ok(m, css);
  assert.ok(Number(m[1]) <= 20);
});
