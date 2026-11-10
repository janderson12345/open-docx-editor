import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { wrapOf, setWrap, WRAPS } from '../src/ooxml/wrap.js';
import { parseXml } from '../src/ooxml/xml.js';

const graphic = '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="x"><pic/></a:graphicData></a:graphic>';
const inline = `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="7" name="Picture 7"/><wp:cNvGraphicFramePr/>${graphic}</wp:inline></w:drawing></w:r>`;
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
/** The anchor's children's names, in order: Word reads them only in the schema's order. */
const kids = (xml) => [...parseXml(`<w:p ${W}>${xml}</w:p>`).getElementsByTagName('*')].find((e) => e.localName === 'anchor' || e.localName === 'inline').childNodes;
const names = (xml) => [...kids(xml)].map((n) => n.localName);

test('Wrap Text: each choice written as Word writes it, and read back', () => {
  assert.equal(wrapOf(inline), 'inline');
  for (const [kind] of WRAPS) {
    const xml = setWrap(inline, kind);
    assert.equal(wrapOf(xml), kind, kind);
    assert.match(xml, /<wp:docPr id="7" name="Picture 7"\/>/);
    assert.ok(xml.includes(graphic), `${kind}: the picture kept`);
    if (kind === 'inline') { assert.equal(xml, inline); continue; }
    const order = names(xml);
    assert.deepEqual(order.slice(0, 5), ['simplePos', 'positionH', 'positionV', 'extent', 'effectExtent'], kind);
    assert.ok(order[5].startsWith('wrap'), `${kind}: its wrap element after the extents`);
    assert.deepEqual(order.slice(6), ['docPr', 'cNvGraphicFramePr', 'graphic'], kind);
  }
  assert.match(setWrap(inline, 'behind'), /behindDoc="1"/);
  assert.match(setWrap(inline, 'front'), /behindDoc="0"/);
});

test('Wrap Text: a floating picture changes its wrapping in place, or goes back in line', () => {
  const square = setWrap(inline, 'square').replace('<wp:posOffset>0</wp:posOffset></wp:positionH>', '<wp:posOffset>1270000</wp:posOffset></wp:positionH>');
  const tight = setWrap(square, 'tight');
  assert.equal(wrapOf(tight), 'tight');
  assert.match(tight, /<wp:posOffset>1270000<\/wp:posOffset><\/wp:positionH>/, 'its place kept');
  assert.equal((tight.match(/<wp:wrap(?!Polygon)/g) || []).length, 1, 'one wrap element');
  const behind = setWrap(tight, 'behind');
  assert.equal(wrapOf(behind), 'behind');
  assert.equal(wrapOf(setWrap(behind, 'front')), 'front');
  const back = setWrap(behind, 'inline');
  assert.equal(wrapOf(back), 'inline');
  assert.deepEqual(names(back), ['extent', 'effectExtent', 'docPr', 'cNvGraphicFramePr', 'graphic']);
});
