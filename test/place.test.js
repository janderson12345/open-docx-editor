import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { placeVar, padOf, cssDecls, cssText } from '../src/ooxml/place.js';
import { padPlace, turned, fromParagraphTop } from '../src/ooxml/reader.js';
import { parseXml, unescapeXml } from '../src/ooxml/xml.js';

// A drawing's place is CSS with custom properties the layout reads back (src/ooxml/place.js).
test('placeVar and padOf read a place\'s custom properties, by name, wherever they are in it', () => {
  const place = 'float:left;margin:0px 12px 0px 9px;--odx-lrel:page;--odx-ly:-12.5px;--odx-lyalign:none;--odx-pad:3.25px';
  assert.equal(placeVar(place, 'odx-lrel'), 'page');
  assert.equal(parseFloat(placeVar(place, 'odx-ly')), -12.5);
  assert.equal(placeVar(place, 'odx-lyalign'), 'none');
  assert.equal(placeVar(place, 'odx-l'), null); // (a name, not a prefix)
  assert.equal(padOf(place), 3.25);
  // (A style attribute as the browser may give it back, with spaces.)
  assert.equal(padOf('width: 10px; --odx-pad: 2px'), 2);
  assert.equal(padOf(''), 0);
  assert.equal(padOf(null), 0);
  assert.equal(placeVar('a:1;--odx-grown:1', 'odx-grown'), '1');
});

test('cssDecls and cssText give a place back as it was', () => {
  for (const place of [
    'float:left;margin:0px 12px 0px calc(9px - var(--odx-page-left, 96px));--odx-lrel:page',
    'display:block;margin:calc(1px - var(--odx-sb, 0pt)) auto 0px',
    'position:absolute;left:calc(var(--odx-page-left, 96px) + 10px);margin-top:4px;z-index:1;--odx-arel:page;--odx-ay:0px;--odx-ayalign:top',
  ]) assert.equal(cssText(cssDecls(place)), place);
});

// padPlace: a shape drawn with room round its outline (pad px each side, its picture that much larger) keeps its place,
// its margins that much in, whatever order its properties are in.
test('padPlace takes the pad off every margin anchorCss writes, and keeps every other property', () => {
  assert.equal(padPlace('float:left;margin:0px 12px 0px 9px', 3), 'float:left;margin:-3px 9px -3px 6px');
  assert.equal(padPlace('float:right;margin:2px 0px 4px calc(9px - var(--odx-page-left, 96px));--odx-lrel:page', 2),
    'float:right;margin:0px -2px 2px calc(calc(9px - var(--odx-page-left, 96px)) - 2px);--odx-lrel:page');
  // A header or footer shape's --odx-hfwrap after its margin stays as it is (it had been read into the left margin).
  assert.equal(padPlace('float:left;margin:0px 12px 0px 9px;--odx-hfwrap:bothSides 0px 12px 0px 12px', 3),
    'float:left;margin:-3px 9px -3px 6px;--odx-hfwrap:bothSides 0px 12px 0px 12px');
  assert.equal(padPlace('display:block;margin-top:5px;margin-left:10px;margin-bottom:8px;--odx-band:1', 2),
    'display:block;margin-top:3px;margin-left:8px;margin-bottom:6px;--odx-band:1');
  // Centred and aligned right, top and bottom: the sides that are auto stay auto, a 0 one goes past the margin.
  assert.equal(padPlace('display:block;margin:0px auto 6px', 2), 'display:block;margin:-2px auto 4px');
  assert.equal(padPlace('display:block;margin:0px 0 6px auto', 2), 'display:block;margin:-2px -2px 4px auto');
  assert.equal(padPlace('position:absolute;left:0px;margin-top:4px;z-index:1', 2), 'position:absolute;left:0px;margin-top:2px;margin-left:-2px;z-index:1');
  assert.equal(padPlace('float:left;margin:0px 12px 0px 9px', 0), 'float:left;margin:0px 12px 0px 9px');
  assert.equal(padPlace('', 3), '');
});

test("fromParagraphTop: a place's top moved up by its paragraph's space before (none above a page's top), whatever its wrapping", () => {
  const up = (v) => `calc(${v} - min(var(--odx-sb, 0px), var(--odx-sbtop, 100000px)))`;
  const down = (v) => `calc(${v} + min(var(--odx-sb, 0px), var(--odx-sbtop, 100000px)))`;
  assert.equal(fromParagraphTop('float:left;margin:-3px 9px -3px calc(9px - var(--odx-page-left, 96px));--odx-pad:3px'),
    `float:left;margin:${up('-3px')} 9px -3px calc(9px - var(--odx-page-left, 96px));--odx-pad:3px`);
  assert.equal(fromParagraphTop('position:absolute;left:0px;margin-top:4px;z-index:1'), `position:absolute;left:0px;margin-top:${up('4px')};z-index:1`);
  // (Turned, its top already a calc(): kept whole.)
  assert.equal(fromParagraphTop('float:right;margin:calc(5px - 2px) 0px 0px 6px'), `float:right;margin:${up('calc(5px - 2px)')} 0px 0px 6px`);
  // Top and bottom: that space again under it, above the text; not for a column-wide one wrapped round (a band).
  assert.equal(fromParagraphTop('display:block;margin:0px auto 6px'), `display:block;margin:${up('0px')} auto ${down('6px')}`);
  assert.equal(fromParagraphTop('display:block;margin-top:5px;margin-left:10px;margin-bottom:8px'), `display:block;margin-top:${up('5px')};margin-left:10px;margin-bottom:${down('8px')}`);
  assert.equal(fromParagraphTop('display:block;margin-top:5px;margin-left:10px;margin-bottom:8px;--odx-band:1'), `display:block;margin-top:${up('5px')};margin-left:10px;margin-bottom:8px;--odx-band:1`);
  assert.equal(fromParagraphTop(''), '');
  assert.equal(fromParagraphTop('z-index:1'), 'z-index:1');
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
const drawing = (rot) => parseXml(`<w:drawing ${W}><wp:anchor><wp:extent cx="914400" cy="457200"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:xfrm rot="${rot * 60000}"><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`).documentElement;

test('turned: a floating shape\'s margins widened for its turned outline, its other properties kept', () => {
  // 96 x 48 at 90 degrees: 24px more each side down, 24px less across.
  const out = turned('float:left;margin:0px 12px 0px 9px;--odx-hfwrap:bothSides 0px 12px 0px 12px', drawing(90), 96, 48);
  assert.equal(out, 'float:left;margin:-24px 12px 24px 9px;--odx-hfwrap:bothSides 0px 12px 0px 12px;transform:translate(0px, 24px) rotate(90deg)');
  const block = turned('display:block;margin-top:5px;margin-left:10px;margin-bottom:8px', drawing(90), 96, 48);
  assert.equal(block, 'display:block;margin-top:calc(5px - 24px);margin-left:10px;margin-bottom:32px;transform:translate(0px, 24px) rotate(90deg)');
  // Unturned: as it was.
  assert.equal(turned('float:left;margin:0px 12px 0px 9px', drawing(0), 96, 48), 'float:left;margin:0px 12px 0px 9px');
});

test('unescapeXml decodes the named and numeric entities, &amp; last', () => {
  assert.equal(unescapeXml('&quot;A&amp;B&quot; &lt;x&gt; &apos;y&apos;'), '"A&B" <x> \'y\'');
  assert.equal(unescapeXml('&#65;&#x42;&#x1F600;'), 'AB\u{1F600}');
  assert.equal(unescapeXml('&amp;lt;'), '&lt;'); // (not "<": decoded once, as the parser would)
  assert.equal(unescapeXml('&#1114112;'), '\uFFFD'); // (past the last code point)
  assert.equal(unescapeXml(null), null);
});
