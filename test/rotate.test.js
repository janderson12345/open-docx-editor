import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxSession } from '../src/ooxml/docx.js';

// Turned shapes (a:xfrm rot), as Word places them (its PDFs of shapes it turned itself): the picture drawn unturned and
// turned by the page; in line with the text, drawn its left and top room (wp:effectExtent) along from its box.
const session = DocxSession.blank();
const shape = (deg) => session.addShape('rect', 96, 72).xml.replace('<wps:spPr><a:xfrm>', `<wps:spPr><a:xfrm rot="${deg * 60000}">`);
const node = (xml) => session._shapeNode(xml);

test('a turned shape in line with the text: turned by the page, as far along as Word draws it', () => {
  const n = node(shape(30).replace(/<wp:effectExtent [^>]*\/>/, '<wp:effectExtent l="133350" t="209550" r="95250" b="190500"/>'));
  assert.equal(n.width, 96);
  assert.equal(n.height, 72);
  assert.match(n.place, /transform:translate\(14px, 22px\) rotate\(30deg\)/);
  // (Unturned: no transform.)
  assert.doesNotMatch(node(session.addShape('rect', 96, 72).xml).place || '', /transform/);
});

test('rotateShape: its rot, and the room round its turned outline (none once turned back)', () => {
  const r = session.rotateShape(session.addShape('rect', 96, 72).xml, 90);
  assert.match(r.xml, /<a:xfrm rot="5400000">/);
  // 96 x 72 at 90 degrees: 24px taller, 24px narrower: 12px (and half its 1pt outline) above and below, none beside.
  assert.match(r.xml, /<wp:effectExtent l="0" t="123825" r="0" b="123825"\/>/);
  const back = session.rotateShape(r.xml, 0);
  assert.doesNotMatch(back.xml, /<a:xfrm\b[^>]*\brot="/);
  assert.match(back.xml, /<wp:effectExtent l="0" t="0" r="0" b="0"\/>/);
  // Resized while turned: the room follows.
  const longer = session.resizeShape(r.xml, 146, 72);
  assert.match(longer.xml, /<wp:effectExtent l="0" t="361950" r="0" b="361950"\/>/);
});

test('a turned floating shape wrapped round: the text keeps clear of its turned outline, the shape where its box is', () => {
  const xml = shape(30)
    .replace(/<wp:inline distT="0" distB="0" distL="0" distR="0">/, '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>952500</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>95250</wp:posOffset></wp:positionV>')
    .replace(/(<wp:effectExtent [^>]*\/>)/, '<wp:effectExtent l="0" t="190500" r="0" b="190500"/><wp:wrapSquare wrapText="bothSides"/>')
    .replace('</wp:inline>', '</wp:anchor>');
  const n = node(xml);
  // 96 x 72 at 30 degrees: 119.1 x 110.4, so 11.57px wider and 19.18px taller each side. Its top room isn't added to
  // its offset (10px), as for an unturned shape: Word draws it about its box's centre.
  const m = /float:left;margin:(-?[\d.]+)px (-?[\d.]+)px (-?[\d.]+)px (-?[\d.]+)px;transform:translate\(([\d.]+)px, ([\d.]+)px\) rotate\(30deg\)/.exec(n.place);
  assert.ok(m, n.place);
  const [t, r, b, l, tx, ty] = m.slice(1).map(Number);
  assert.ok(Math.abs(t - (10 - 19.18)) < 0.02 && Math.abs(l - (100 - 11.57)) < 0.02, n.place);
  assert.ok(Math.abs(tx - 11.57) < 0.02 && Math.abs(ty - 19.18) < 0.02, n.place);
  assert.ok(Math.abs(r - (12 + 11.57)) < 0.02 && Math.abs(b - 19.18) < 0.02, n.place);
});
