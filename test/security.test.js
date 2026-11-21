import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readPPr, readRPr, pPrToCss, rPrToCss, readBorders, borderCss, shdFill, fontStack, highlightColor } from '../src/ooxml/props.js';
import { fromStrict } from '../src/ooxml/strict.js';
import { parseXml, kid } from '../src/ooxml/xml.js';
import { strToU8, strFromU8 } from 'fflate';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const el = (xml) => parseXml(`<w:x ${W}>${xml}</w:x>`).documentElement;
// What a crafted file tries to slip into CSS: closing the rule or declaration, or fetching from elsewhere.
const ESCAPES = /[{}]|url\(|;\s*background|expression|<\/?style/i;
const evil = 'left}*{background:url(https://example.invalid/x)}x{';

test('a document\'s values can\'t write CSS of their own', () => {
  const p = readPPr(kid(el(`<w:pPr><w:jc w:val="${evil}"/><w:shd w:fill="${evil}"/></w:pPr>`), 'pPr'));
  assert.doesNotMatch(pPrToCss(p), ESCAPES);
  const r = readRPr(kid(el(`<w:rPr><w:color w:val="000;background:url(https://example.invalid/c)"/><w:highlight w:val="${evil}"/><w:rFonts w:ascii="Arial&quot;;background:url(https://example.invalid/f);x:&quot;"/></w:rPr>`), 'rPr'));
  // (The font name stays inside its quoted string, which it can't close: what's in it is a name, not CSS.)
  const css = rPrToCss(r, {});
  assert.match(css, /font-family:"[^"]*", sans-serif;/);
  assert.doesNotMatch(css.replace(/"[^"]*"/g, '""'), ESCAPES);
  const b = readBorders(el(`<w:top w:val="single" w:sz="4" w:color="${evil}"/>`));
  assert.doesNotMatch(borderCss(b.top), ESCAPES);
  assert.equal(shdFill(kid(el(`<w:shd w:fill="${evil}"/>`), 'shd')), null);
  assert.doesNotMatch(fontStack('A"}body{background:url(x)}').replace(/"[^"]*"/g, '""'), ESCAPES);
  assert.equal(highlightColor(evil), 'transparent');
  // Ordinary values still work.
  assert.equal(pPrToCss(readPPr(kid(el('<w:pPr><w:jc w:val="center"/></w:pPr>'), 'pPr'))), '--odx-trail:1;text-align:center');
  assert.match(rPrToCss(readRPr(kid(el('<w:rPr><w:color w:val="C00000"/><w:highlight w:val="yellow"/></w:rPr>'), 'rPr')), {}), /color:#C00000;background-color:#ffff00/);
});

test('a crafted Strict part full of unclosed tags is read in moments', () => {
  const S = 'http://purl.oclc.org/ooxml/wordprocessingml/main';
  const t0 = performance.now();
  fromStrict({ 'word/document.xml': strToU8(`<x xmlns:w="${S}">` + '<w:a '.repeat(200000) + '</x>') });
  assert.ok(performance.now() - t0 < 1000, 'under a second');
  assert.match(strFromU8(fromStrict({ 'word/document.xml': strToU8(`<w:document xmlns:w="${S}"><w:tblW w:w="50%" w:type="pct"/></w:document>`) })['word/document.xml']), /w:w="2500"/);
});

test('Strict charts\' percentages are plain numbers', () => {
  const C = 'http://purl.oclc.org/ooxml/drawingml/chart';
  const out = strFromU8(fromStrict({ 'word/charts/chart1.xml': strToU8(`<c:chartSpace xmlns:c="${C}"><c:gapWidth val="150%"/><c:overlap val="-20%"/></c:chartSpace>`) })['word/charts/chart1.xml']);
  assert.match(out, /<c:gapWidth val="150"\/><c:overlap val="-20"\/>/);
});

// Tags whose two attributes were read by one regex (<x\b[^>]*a="..."[^>]*b="...") took time growing with the square of
// a crafted tag's length: the first attribute repeated inside another one's single-quoted value (valid XML) gave the
// second [^>]* a place to start at each. Read tag by tag now (xml.js startTag), each in moments.
test('tags read two attributes at a time in moments, however long a crafted one is', async () => {
  const { startTag, tagAttr, setTagAttr } = await import('../src/ooxml/xml.js');
  const { w14Color } = await import('../src/ooxml/texteffects.js');
  const { shapeInfo } = await import('../src/editor/imgresize.js');
  const { docProps } = await import('../src/editor/pagefields.js');
  const { DocxSession } = await import('../src/ooxml/docx.js');
  const many = (s) => s.repeat(40000);
  const quick = (what, fn) => { const t = performance.now(); fn(); const ms = performance.now() - t; assert.ok(ms < 1000, `${what}: ${Math.round(ms)}ms`); };
  // The reader itself: a '>' in a quoted value isn't the tag's end, and the attribute is the one by that name.
  const t = startTag(`<a x='>"' cx="5" w:cx="6"/>`, 'a');
  assert.equal(t.tag, `<a x='>"' cx="5" w:cx="6"/>`);
  assert.equal(tagAttr(t.tag, 'cx'), '5');
  assert.equal(tagAttr(t.tag, 'w:cx'), '6');
  assert.equal(setTagAttr(t.tag, 'cx', '7'), `<a x='>"' cx="7" w:cx="6"/>`);
  assert.equal(startTag('<ab/><a/>', 'a').at, 5); // (a longer name isn't it)
  quick('w14Color', () => assert.equal(w14Color(`<w14:srgbClr x='${many('w14:val="FF0000" ')}' w14:val="00FF00"/>`), '#00FF00'));
  quick('shapeInfo', () => shapeInfo(`<wps:wsp><wps:spPr><a:xfrm/><a:ext q='${many('cx="1" ')}' cx="9" cy="9"/><a:prstGeom prst="rect"><a:avLst><a:gd z='${many('name="adj" ')}' name="adj" fmla="val 1"/></a:avLst></a:prstGeom></wps:spPr></wps:wsp>`));
  quick('docProps', () => docProps({ text: (p) => (p === 'docProps/custom.xml' ? `<Properties><property q='${many('name="x" ')}' name="Client"><vt:lpwstr>Ann</vt:lpwstr></property></Properties>` : '') }));
  const session = DocxSession.blank();
  const crafted = session.addShape('rect', 96, 72).xml.replace('<wp:extent ', `<wp:extent q='${many('cx="1" ')}' `);
  let resized;
  quick('resizeShape', () => { resized = session.resizeShape(crafted, 48, 36); });
  assert.match(resized.xml, /<wp:extent q='[^']*' cx="457200" cy="342900"\/>/); // (its own cx and cy, not those in q)
  assert.match(resized.xml, /<a:ext cx="457200" cy="342900"\/>/);
});

test('a long run of spaces that does not end its paragraph is marked in time that grows with its length', async () => {
  const { schema } = await import('../src/editor/schema.js');
  const { spanDecos } = await import('../src/editor/textspans.js');
  const para = (n) => schema.node('doc', null, [schema.node('paragraph', null, [schema.text('a' + ' '.repeat(n) + 'b '.repeat(3))])]);
  const time = (n) => { const d = para(n); let best = Infinity; for (let i = 0; i < 3; i++) { const t0 = performance.now(); spanDecos(d); best = Math.min(best, performance.now() - t0); } return best; };
  time(1000);
  const big = time(200000);
  // (The pattern " +$" took seconds here: 200,000 spaces before a letter, each a place it tried again from.)
  assert.ok(big < 250, `${big.toFixed(1)}ms for 200,000 spaces`);
  // The closing spaces are still marked, from the first of them.
  const decos = spanDecos(para(5));
  const trail = decos.find((d) => d.type.attrs.class === 'odx-trail');
  assert.ok(trail && trail.to - trail.from === 1, 'the one closing space is marked');
});
