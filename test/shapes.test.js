import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { DocxSession } from '../src/ooxml/docx.js';

const bytes = fs.readFileSync(new URL('./fixtures/textbox.docx', import.meta.url));

/** The fixture's floating text box made as wide as the column (624px): text can't go beside it, so it's a band. */
function wideBox() {
  const files = unzipSync(bytes);
  const xml = strFromU8(files['word/document.xml']).replace(/(<wp:anchor\b[\s\S]*?<wp:extent cx=")\d+(" cy="\d+")/, `$1${624 * 9525}$2`).replace(/(<wp:positionH relativeFrom="column"><wp:posOffset>)\d+/, '$10');
  files['word/document.xml'] = strToU8(xml);
  const s = new DocxSession(zipSync(files));
  let node = null;
  s.doc.descendants((n) => { if (!node && n.type.name === 'image' && /<wp:anchor/.test(n.attrs.xml)) node = n; });
  return { s, node };
}

test('a box as wide as the column stays a band (text above and below it) when it is resized or moved', () => {
  const { s, node } = wideBox();
  assert.match(node.attrs.place, /--odx-band:1/, 'read as a band');
  const bigger = s.resizeShape(node.attrs.xml, 624, node.attrs.height + 20);
  assert.match(bigger.place, /--odx-band:1/, 'after a resize');
  const moved = s.moveShape(node.attrs.xml, 0, 16);
  assert.match(moved.place, /--odx-band:1/, 'after a move');
});

test('a resize changes the frame, the shape and the VML copy alike, and nothing else of the shape', () => {
  const { s, node } = wideBox();
  const r = s.resizeShape(node.attrs.xml, 300, 120);
  assert.equal(r.width, 300);
  assert.equal(r.height, 120);
  assert.match(r.xml, new RegExp(`<wp:extent cx="${300 * 9525}" cy="${120 * 9525}"`));
  assert.match(r.xml, new RegExp(`<a:ext cx="${300 * 9525}" cy="${120 * 9525}"`));
  assert.match(r.xml, /width:225pt;height:90pt/);
  const strip = (x) => x.replace(/cx="\d+"|cy="\d+"|width:[\d.]+pt;height:[\d.]+pt/g, '');
  assert.equal(strip(r.xml), strip(node.attrs.xml));
});

test('a move shifts the offsets and the VML margins; a box placed by alignment isn\'t moved', () => {
  const { s, node } = wideBox();
  const m = s.moveShape(node.attrs.xml, 10, -4);
  const off = (x, a) => Number(new RegExp(`<wp:position${a}\\b[^>]*>\\s*<wp:posOffset>(-?\\d+)`).exec(x)[1]);
  assert.equal(off(m.xml, 'H') - off(node.attrs.xml, 'H'), 10 * 9525);
  assert.equal(off(m.xml, 'V') - off(node.attrs.xml, 'V'), -4 * 9525);
  assert.match(m.xml, /margin-left:295.5pt;margin-top:4.5pt/);
  const aligned = node.attrs.xml.replace(/<wp:posOffset>-?\d+<\/wp:posOffset>(<\/wp:positionH>)/, '<wp:align>center</wp:align>$1');
  assert.equal(s.moveShape(aligned, 10, 0), null);
});
