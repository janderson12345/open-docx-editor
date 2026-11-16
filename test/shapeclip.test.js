import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { DocxSession } from '../src/ooxml/docx.js';
import { shapesSvg } from '../src/ooxml/shapes.js';
import { parseXml } from '../src/ooxml/xml.js';
import { placeVar, padOf } from '../src/ooxml/place.js';

// shapeclip-word.docx (made in Word, its print checked): two shapes in line with the text with 6pt and 8pt outlines; a
// floating rectangle with a 10pt outline, text on both sides; straight lines with arrowheads in front of the text,
// behind it and wrapped round; a text box with an 8pt outline.
const s = new DocxSession(fs.readFileSync(new URL('./fixtures/shapeclip-word.docx', import.meta.url)));
const shapes = [];
s.doc.descendants((n) => { if (n.type.name === 'image' && /wps:wsp/.test(n.attrs.xml)) shapes.push(n.attrs); });
const named = (name) => shapes.find((a) => a.xml.includes(`name="${name}"`));
const px = (v) => Number(v) / 9525;
const offsetV = (a) => px(/<wp:positionV relativeFrom="paragraph"><wp:posOffset>(-?\d+)/.exec(a.xml)[1]);

test('every shape is drawn: straight lines as tall or as wide as nothing too', () => {
  assert.equal(shapes.length, 8, shapes.map((a) => /name="([^"]+)"/.exec(a.xml)?.[1]).join(', '));
  for (const name of ['Straight Connector 4', 'Straight Connector 5', 'Straight Connector 7']) {
    const a = named(name);
    assert.ok(a.src && (a.width === 0 || a.height === 0), name);
    assert.ok(padOf(a.place) > 0, `${name}: room drawn round it for its arrowheads (${a.place})`);
  }
  // (One across the page, unpadded: still a picture, its line along its top.)
  const d = parseXml('<w:drawing xmlns:w="w" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphic><a:graphicData><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="0"/></a:xfrm><a:prstGeom prst="line"/><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr></wps:wsp></a:graphicData></a:graphic></w:drawing>').documentElement;
  assert.match(shapesSvg(d, {}, 100, 0, 2) || '', /<path d="M0 0L100 0"/);
  assert.equal(shapesSvg(d, {}, 0, 0, 2), null);
});

test('a shape in line with the text: room drawn round its outline, its line\'s room as before, cut to the 1/64px the browser lays out in', () => {
  for (const name of ['Rectangle 1', 'Oval 2']) {
    const a = named(name);
    assert.match(a.xml, /<wp:inline/);
    const pad = padOf(a.place);
    assert.ok(pad >= 5, `${name}: ${a.place}`);
    for (const k of ['odx-fx', 'odx-fxt', 'odx-fxl', 'odx-fxr']) {
      const v = parseFloat(placeVar(a.place, k));
      assert.ok(Number.isInteger(v * 64) && v < 0, `${name} --${k}: ${v}`);
    }
  }
});

// (Floating in front of the text: its picture's top margin is its offset, to the px, less the room drawn round it.)
const top = (a) => parseFloat(/margin-top:(-?[\d.]+)px/.exec(a.place)[1]) + padOf(a.place);
const roomAbove = (a) => px(/<wp:effectExtent l="\d+" t="(\d+)"/.exec(a.xml)[1]);

test('a floating shape sits at its offset, its effect extent not added: a line with arrowheads, a text box', () => {
  for (const name of ['Straight Connector 4', 'Text Box 8']) {
    const a = named(name);
    assert.ok(roomAbove(a) > 0, name);
    assert.ok(Math.abs(top(a) - offsetV(a)) <= 0.5, `${name}: ${a.place}`);
  }
});

test('a text box\'s text is inset from the inside of its outline: half the outline in, past the insets', () => {
  const svg = decodeURIComponent(named('Text Box 8').src.replace(/^data:image\/svg\+xml[^,]*,/, ''));
  // (The default insets, 0.1in and 0.05in, and half its 8pt outline: 5.33px.)
  const [x, y] = /<text x="([\d.]+)" y="([\d.]+)"/.exec(svg).slice(1).map(Number);
  assert.ok(Math.abs(x - (9.6 + 5.33)) < 0.1, `x ${x}`);
  assert.ok(y > 4.8 + 5.33 && y < 4.8 + 5.33 + 16, `y ${y}`);
});

test('in a Word 2010 document (compatibility mode 14), one centred across the page sits its effect extent lower', () => {
  // (The fixture in mode 14, its line and text box centred: Word's print has them that much lower, sample S078's box
  // too. Mode 14 with an offset across, or centred in mode 15, as above: at the offset.)
  const files = unzipSync(new Uint8Array(fs.readFileSync(new URL('./fixtures/shapeclip-word.docx', import.meta.url))));
  const centre = (xml) => xml.replace(/<wp:positionH relativeFrom="column"><wp:posOffset>-?\d+<\/wp:posOffset><\/wp:positionH>((?:(?!<\/wp:anchor>)[\s\S])*?name="(?:Straight Connector 4|Text Box 8)")/g, '<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>$1');
  const read = (mode, centred) => {
    const f = { ...files };
    f['word/settings.xml'] = strToU8(strFromU8(files['word/settings.xml']).replace(/(w:name="compatibilityMode" w:uri="[^"]*" w:val=")\d+/, `$1${mode}`));
    if (centred) f['word/document.xml'] = strToU8(centre(strFromU8(files['word/document.xml'])));
    const out = [];
    new DocxSession(zipSync(f)).doc.descendants((n) => { if (n.type.name === 'image' && /name="(Straight Connector 4|Text Box 8)"/.test(n.attrs.xml)) out.push(n.attrs); });
    return out;
  };
  const centred14 = read(14, true);
  assert.equal(centred14.length, 2);
  for (const a of centred14) {
    assert.match(a.xml, /<wp:align>center<\/wp:align>/);
    assert.ok(Math.abs(top(a) - offsetV(a) - roomAbove(a)) <= 0.5, a.place);
  }
  for (const a of [...read(14, false), ...read(15, true)]) assert.ok(Math.abs(top(a) - offsetV(a)) <= 0.5, a.place);
});

test('a floating shape wrapped round keeps the text clear of its effect extent past its wrap distance', () => {
  const rect = named('Rectangle 3');
  const [t, r, b, l, fxT] = placeVar(rect.place, 'odx-both').split(' ').map(parseFloat);
  const fx = /<wp:effectExtent l="(\d+)" t="(\d+)" r="(\d+)" b="(\d+)"/.exec(rect.xml).slice(1).map(px);
  assert.deepEqual([t, r, b, l], [fx[1], 12 + fx[2], fx[3], 12 + fx[0]].map((v) => +v.toFixed(2)), rect.place);
  // (And the extent's share of the top apart: a paragraph above, its space after, doesn't reach into it: wrapboth.js.)
  assert.equal(fxT, +fx[1].toFixed(2), rect.place);
});

test('an outline\'s corners: round on a shape with no style (a text box), mitred on one taking its theme line style', () => {
  const svgOf = (a) => decodeURIComponent(a.src.replace(/^data:image\/svg\+xml[^,]*,/, ''));
  assert.match(svgOf(named('Text Box 8')), /<path d="[^"]*Z?"[^>]*stroke-linejoin="round"/);
  assert.match(svgOf(named('Rectangle 3')), /stroke-miterlimit="8"/);
  assert.doesNotMatch(svgOf(named('Rectangle 3')), /stroke-linejoin="round"/);
});
