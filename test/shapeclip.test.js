import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
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

test('a floating shape without text sits at its offset (its effect extent not added); a text box with it', () => {
  // (Floating in front of the text: its picture's top margin is its offset, to the px, less the room drawn round it.)
  const line = named('Straight Connector 4');
  assert.ok(Math.abs(parseFloat(/margin-top:(-?[\d.]+)px/.exec(line.place)[1]) + padOf(line.place) - offsetV(line)) <= 0.5, line.place);
  const box = named('Text Box 8');
  const fxT = px(/<wp:effectExtent l="\d+" t="(\d+)"/.exec(box.xml)[1]);
  assert.ok(Math.abs(parseFloat(/margin-top:(-?[\d.]+)px/.exec(box.place)[1]) + padOf(box.place) - offsetV(box) - fxT) <= 0.5, box.place);
});

test('a floating shape wrapped round keeps the text clear of its effect extent past its wrap distance', () => {
  const rect = named('Rectangle 3');
  const [t, r, b, l] = placeVar(rect.place, 'odx-both').split(' ').map(parseFloat);
  const fx = /<wp:effectExtent l="(\d+)" t="(\d+)" r="(\d+)" b="(\d+)"/.exec(rect.xml).slice(1).map(px);
  assert.deepEqual([t, r, b, l], [fx[1], 12 + fx[2], fx[3], 12 + fx[0]].map((v) => +v.toFixed(2)), rect.place);
});
