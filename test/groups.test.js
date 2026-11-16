import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { shapesSvg, shapesPad, gradientStops, turnedInside } from '../src/ooxml/shapes.js';
import { parseXml } from '../src/ooxml/xml.js';

const NSS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"';
const E = 9525; // EMU per px
const drawing = (inner) => parseXml(`<w:drawing ${NSS}><a:graphic><a:graphicData>${inner}</a:graphicData></a:graphic></w:drawing>`).documentElement;
const xfrm = (x, y, w, h, attrs = '') => `<a:xfrm${attrs}><a:off x="${x * E}" y="${y * E}"/><a:ext cx="${w * E}" cy="${h * E}"/></a:xfrm>`;
const sp = (x, y, w, h, attrs = '', more = '') => `<wps:wsp><wps:spPr>${xfrm(x, y, w, h, attrs)}<a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="156082"/></a:solidFill></wps:spPr>${more}</wps:wsp>`;
/** A group w x h px whose children were drawn in cw x ch (stretched to w x h). */
const group = (w, h, cw, ch, kids) => `<wpg:wgp><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w * E}" cy="${h * E}"/><a:chOff x="0" y="0"/><a:chExt cx="${cw * E}" cy="${ch * E}"/></a:xfrm></wpg:grpSpPr>${kids}</wpg:wgp>`;
const ctx = { image: () => null, colors: {} };

test('a group stretched taller: its turned shapes are stretched first, then turned, not skewed', () => {
  // groupturn-word: a group 1.8 times as tall as it was drawn. A rectangle at 30 degrees keeps square corners, 80 x 90.
  const svg = shapesSvg(drawing(group(200, 180, 200, 100, sp(10, 20, 80, 50, ' rot="1800000"'))), ctx, 200, 180);
  assert.doesNotMatch(svg, /matrix\(1 0 0 1\.8/);
  assert.match(svg, /translate\(10\.00 36\.00\) rotate\(30 40 45\)/);
  // One turned nearer upright (60 degrees) is stretched along its own width: 144 x 50, about the same centre (50, 81).
  const up = shapesSvg(drawing(group(200, 180, 200, 100, sp(10, 20, 80, 50, ' rot="3600000"'))), ctx, 200, 180);
  assert.match(up, /translate\(-22\.00 56\.00\) rotate\(60 72 25\)/);
});

test('a turned or flipped group inside a group turns and flips its shapes with it, about its own centre', () => {
  const inner = `<wpg:grpSp><wpg:grpSpPr>${xfrm(20, 10, 100, 40, ' rot="1500000"').replace('</a:xfrm>', '<a:chOff x="0" y="0"/><a:chExt cx="952500" cy="381000"/></a:xfrm>')}</wpg:grpSpPr>${sp(0, 0, 40, 40)}${sp(60, 0, 40, 40)}</wpg:grpSp>`;
  const svg = shapesSvg(drawing(group(200, 100, 200, 100, inner)), ctx, 200, 100);
  assert.match(svg, /<g transform="rotate\(25 70 30\)"><g transform="translate\(20\.00 10\.00\)/);
  const flipped = shapesSvg(drawing(group(200, 100, 200, 100, inner.replace(' rot="1500000"', ' flipH="1"'))), ctx, 200, 100);
  assert.match(flipped, /<g transform="matrix\(-1 0 0 1 140 0\)">/);
});

test('a flipped shape\'s text isn\'t mirrored: kept as it was flipped left to right, upside down flipped top to bottom', () => {
  const txt = '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/>';
  const v = shapesSvg(drawing(group(200, 100, 200, 100, sp(0, 0, 80, 50, ' rot="900000" flipV="1"', txt))), ctx, 200, 100);
  // The shape mirrored; its text in a group of its own, turned 15 + 180 degrees.
  assert.match(v, /rotate\(15 40 25\) matrix\(1 0 0 -1 0 50\)/);
  assert.match(v, /<g transform="translate\(0\.00 0\.00\) rotate\(195 40 25\)">[^]*Text/);
  const h = shapesSvg(drawing(group(200, 100, 200, 100, sp(0, 0, 80, 50, ' flipH="1"', txt))), ctx, 200, 100);
  assert.match(h, /<g transform="translate\(0\.00 0\.00\)">[^]*Text/);
});

test('room round a group whose turned shapes reach past its box (Word draws them there)', () => {
  // A 100 x 60 rectangle at 45 degrees in the group's corner: its corners reach about 27px past the box.
  const d = drawing(group(200, 100, 200, 100, sp(0, 0, 100, 60, ' rot="2700000"')));
  assert.ok(shapesPad(d) >= 27, String(shapesPad(d)));
  assert.equal(shapesPad(drawing(group(200, 100, 200, 100, sp(10, 10, 100, 60)))), 0);
});

/** A rectangle 100 x 50 px with gradient fill `fill` (the a:gradFill's inside). */
const gradient = (fill, attrs = '') => shapesSvg(drawing(`<wps:wsp><wps:spPr>${xfrm(0, 0, 100, 50, attrs)}<a:prstGeom prst="rect"/><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="C02020"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst>${fill}</a:gradFill></wps:spPr></wps:wsp>`), ctx, 100, 50);

test('gradient fills: linear along their angle, scaled to the box or not, and path gradients from their focus', () => {
  // Top to bottom (90 degrees): from the top edge's middle to the bottom's.
  assert.match(gradient('<a:lin ang="5400000" scaled="0"/>'), /<linearGradient id="odxg1" gradientUnits="userSpaceOnUse" x1="50" y1="0" x2="50" y2="50">/);
  // 45 degrees scaled: worked out in a unit square, stretched to the box (gradients-word's wide rectangle).
  assert.match(gradient('<a:lin ang="2700000" scaled="1"/>'), /x1="0" y1="0" x2="1" y2="1" gradientTransform="matrix\(100 0 0 50 0 0\)"/);
  // From the centre (a:path rect, focus 50% in from each side): four sides clipped to the shape, from the centre out.
  const rect = gradient('<a:path path="rect"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>');
  assert.match(rect, /<clipPath id="odxg1c">/);
  assert.equal((rect.match(/<linearGradient /g) || []).length, 4);
  assert.match(rect, /id="odxg1s0" gradientUnits="userSpaceOnUse" x1="0" y1="25" x2="0" y2="0"/);
  // From the corner: the focus on the top left corner, the tile a box past it, so its last colour is at the far corner.
  const corner = gradient('<a:path path="rect"><a:fillToRect r="100000" b="100000"/></a:path><a:tileRect l="-100000" t="-100000"/>');
  assert.match(corner, /id="odxg1s1" gradientUnits="userSpaceOnUse" x1="0" y1="0" x2="100" y2="0"/);
  // A circle's: rings out to the corners.
  assert.match(gradient('<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>'), /<radialGradient id="odxg1" gradientUnits="userSpaceOnUse" cx="0\.5" cy="0\.5" r="0\.7071"/);
  // Not turned with the shape (rotWithShape="0"): a shape turned 30 degrees keeps its top-to-bottom gradient level.
  const fixed = shapesSvg(drawing(`<wpg:wgp><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${200 * E}" cy="${100 * E}"/><a:chOff x="0" y="0"/><a:chExt cx="${200 * E}" cy="${100 * E}"/></a:xfrm></wpg:grpSpPr><wps:wsp><wps:spPr>${xfrm(0, 0, 100, 50, ' rot="1800000"')}<a:prstGeom prst="rect"/><a:gradFill rotWithShape="0"><a:gsLst><a:gs pos="0"><a:srgbClr val="C02020"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></wps:spPr></wps:wsp></wpg:wgp>`), ctx, 200, 100);
  const [, x1, y1, x2, y2] = fixed.match(/x1="([-\d.]+)" y1="([-\d.]+)" x2="([-\d.]+)" y2="([-\d.]+)"/).map(Number);
  assert.equal(Math.round((Math.atan2(y2 - y1, x2 - x1) * 180) / Math.PI), 60);
});

test('gradient colours blended as Word blends them: in linear light, easing in and out', () => {
  const s = gradientStops([{ pos: 0, c: { hex: 'C02020', alpha: 1 } }, { pos: 1, c: { hex: 'FFFFFF', alpha: 1 } }]);
  assert.equal(s.length, 9);
  const green = (pos) => parseInt(s.find((x) => Math.abs(x.pos - pos) < 1e-9).c.hex.slice(2, 4), 16);
  // Word's PDF (gradients-word): 184 halfway; a straight blend would give 145.
  assert.ok(Math.abs(green(0.5) - 184) <= 6, String(green(0.5)));
  assert.ok(green(0.125) < 70 && green(0.875) > 240);
});

test('SmartArt with a gradient style draws its gradients (smartart.docx: the organization chart\'s boxes)', async () => {
  const fs = await import('node:fs');
  const { DocxSession } = await import('../src/ooxml/docx.js');
  const s = new DocxSession(new Uint8Array(fs.readFileSync(new URL('./fixtures/smartart.docx', import.meta.url))));
  const svgs = [];
  s.doc.descendants((n) => { if (n.type.name === 'image' && /^data:image\/svg/.test(n.attrs.src || '')) svgs.push(decodeURIComponent(n.attrs.src.split(',')[1])); });
  assert.equal((svgs[2].match(/<linearGradient /g) || []).length, 5);
  assert.match(svgs[2], /fill="url\(#odxg\d+\)"/);
  assert.doesNotMatch(svgs[0], /Gradient/);
});

test('a group with shapes turned inside it is known, so its room for them doesn\'t move it down (reader.js)', () => {
  assert.equal(turnedInside(drawing(group(200, 100, 200, 100, sp(0, 0, 80, 50, ' rot="1800000"')))), true);
  assert.equal(turnedInside(drawing(group(200, 100, 200, 100, sp(0, 0, 80, 50, ' flipH="1"')))), false);
  assert.equal(turnedInside(drawing(sp(0, 0, 80, 50, ' rot="1800000"'))), false);
});

test('a shape\'s text placed top to bottom as its bodyPr anchor says: centred (Word\'s own shapes), at the bottom, at the top', () => {
  const svg = (anchor) => shapesSvg(drawing(sp(0, 0, 200, 110, '', `<wps:txbx><w:txbxContent><w:p><w:r><w:t>One line</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr${anchor ? ` anchor="${anchor}"` : ''}/>`)), ctx, 200, 110);
  const dy = (s) => Number((s.match(/<g transform="translate\(0 ([\d.]+)\)"><text/) || [])[1] || 0);
  // (anchor.docx, made by Word: one line in a 200 x 110 shape, within 0.12pt of Word's centred, 0.5pt of its bottom.)
  assert.ok(dy(svg('ctr')) > 30, String(dy(svg('ctr'))));
  assert.ok(Math.abs(dy(svg('b')) - 2 * dy(svg('ctr'))) < 0.02);
  assert.equal(dy(svg('t')), 0);
  assert.equal(dy(svg('')), 0);
});
