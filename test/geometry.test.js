import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { guides, geometryPaths, readGeometry, adjustOf, adjustHandles, dragAdjust } from '../src/ooxml/geometry.js';
import { PRESETS, SHAPE_GALLERY } from '../src/ooxml/presets.js';
import { shapesSvg } from '../src/ooxml/shapes.js';
import { parseXml } from '../src/ooxml/xml.js';

test('guide formulas, as DrawingML defines them', () => {
  const g = guides({ av: [['adj', 'val 25000']], gd: [
    ['a', '*/ ss adj 100000'], ['b2', '+- r 0 a'], ['c', '+/ a b2 2'], ['p', 'pin 0 adj 10000'], ['q', '?: a 1 2'],
    ['s', 'sin 100 5400000'], ['k', 'cos 100 10800000'], ['m', 'mod 3 4 0'], ['t', 'at2 1 1'], ['mx', 'max w h'],
  ] }, 200, 100, {});
  assert.equal(g.a, 25);
  assert.equal(g.b2, 175);
  assert.equal(g.c, 100);
  assert.equal(g.p, 10000);
  assert.equal(g.q, 1);
  assert.equal(Math.round(g.s), 100);
  assert.equal(Math.round(g.k), -100);
  assert.equal(g.m, 5);
  assert.equal(Math.round(g.t), 2700000);
  assert.equal(g.mx, 200);
  // A shape's own adjust value replaces the default.
  assert.equal(guides({ av: [['adj', 'val 25000']], gd: [['a', '*/ ss adj 100000']] }, 200, 100, { adj: 'val 50000' }).a, 50);
});

test('paths: lines, arcs (a quarter of an ellipse from its left), a path\'s own coordinates', () => {
  const [rect] = geometryPaths(PRESETS.rect, 100, 50);
  assert.equal(rect.d, 'M0 0L100 0L100 50L0 50Z');
  const [ell] = geometryPaths(PRESETS.ellipse, 100, 50);
  assert.match(ell.d, /^M0 25A50 25 0 0 1 50 0A50 25 0 0 1 100 25A50 25 0 0 1 50 50A50 25 0 0 1 0 25Z$/);
  const [dec] = geometryPaths(PRESETS.flowChartDecision, 100, 50);
  assert.equal(dec.d, 'M0 25L50 0L100 25L50 50Z');
  // Units: EMU definitions drawn at a scale (px).
  const [r2] = geometryPaths(PRESETS.rect, 952500, 476250, {}, 1 / 9525);
  assert.equal(r2.d, 'M0 0L100 0L100 50L0 50Z');
  // A can: its body, a lighter top and the outline as separate paths.
  assert.deepEqual(geometryPaths(PRESETS.can, 100, 100).map((p) => [p.fill, p.stroke]), [['norm', false], ['lighten', false], ['none', true]]);
});

test('custom geometry (a:custGeom) with its own path size', () => {
  const el = parseXml('<a:custGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:avLst/><a:gdLst/><a:pathLst><a:path w="200" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="200" y="100"/></a:lnTo><a:arcTo wR="50" hR="50" stAng="0" swAng="5400000"/><a:close/></a:path></a:pathLst></a:custGeom>').documentElement;
  const [p] = geometryPaths(readGeometry(el), 400, 200, adjustOf(el));
  assert.match(p.d, /^M0 0L400 200A100 100 0 0 1 300 300Z$/);
});

test('every gallery shape is defined or left out; theme colours as Word mixes them', () => {
  for (const [prst] of SHAPE_GALLERY) assert.ok(!PRESETS[prst] || PRESETS[prst].path.length, prst);
  const svg = (clr) => shapesSvg(parseXml(`<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><a:graphic><a:graphicData><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill>${clr}</a:solidFill></wps:spPr></wps:wsp></a:graphicData></a:graphic></w:drawing>`).documentElement, { image: () => null, colors: { accent1: '4472C4' } }, 100, 50);
  // Word's own values: accent 1 darker 25% (lumMod 75%), lighter 40% (lumMod 60% lumOff 40%), a 15% shade in linear light.
  assert.match(svg('<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>'), /fill="#2F5597"/);
  assert.match(svg('<a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>'), /fill="#8FAB?D[BC]"/);
  assert.match(svg('<a:srgbClr val="156082"><a:shade val="15000"/></a:srgbClr>'), /fill="#042433"/);
});

test('SmartArt: drawn from the drawing Word saves with it (shapes, text, colourful variations)', async () => {
  const fs = await import('node:fs');
  const { DocxSession } = await import('../src/ooxml/docx.js');
  const s = new DocxSession(new Uint8Array(fs.readFileSync(new URL('./fixtures/smartart.docx', import.meta.url))));
  const svgs = [];
  s.doc.descendants((n) => { if (n.type.name === 'image' && /^data:image\/svg/.test(n.attrs.src || '')) svgs.push(decodeURIComponent(n.attrs.src.split(',')[1])); });
  assert.equal(svgs.length, 8);
  // Basic Process: three rounded boxes with their text, two arrows between.
  assert.match(svgs[0], />Plan<\/tspan>[\s\S]*>Build<\/tspan>[\s\S]*>Test<\/tspan>/);
  // Basic Block List in a colourful variation: five different fills (hueOff steps).
  const fills = new Set([...svgs[1].matchAll(/<path d="[^"]*" fill="#([0-9A-F]{6})"/g)].map((m) => m[1]));
  assert.ok(fills.size >= 5, [...fills].join(','));
});

test('adjust handles: where they are, and the values that put one where it is dragged', () => {
  const W = 914400, H = 685800; // 96 x 72 px
  // A rounded rectangle's corner handle sits on its top edge, where the rounding ends.
  const [rr] = adjustHandles(PRESETS.roundRect, W, H, {});
  assert.equal(Math.round(rr.x), Math.round(H * 0.16667));
  assert.equal(rr.y, 0);
  // Dragged right, the corner value follows it (within a unit); held to its limit (half the shorter side).
  const v = dragAdjust(PRESETS.roundRect, W, H, {}, 0, rr.x + 95250, 0);
  assert.ok(Math.abs(adjustHandles(PRESETS.roundRect, W, H, v)[0].x - (rr.x + 95250)) < 20, JSON.stringify(v));
  assert.deepEqual(dragAdjust(PRESETS.roundRect, W, H, {}, 0, W * 5, 0), { adj: 'val 50000' });
  // A callout's pointer has no limits: it goes where it's dragged, both ways.
  const [tip] = adjustHandles(PRESETS.wedgeRectCallout, W, H, {});
  const t = dragAdjust(PRESETS.wedgeRectCallout, W, H, {}, 0, tip.x + 400000, tip.y + 300000);
  const now = adjustHandles(PRESETS.wedgeRectCallout, W, H, t)[0];
  assert.ok(Math.hypot(now.x - tip.x - 400000, now.y - tip.y - 300000) < 30, JSON.stringify(t));
  // A polar handle (a block arc's angle) lands on its circle, at the angle of the point dragged to.
  const arc = dragAdjust(PRESETS.blockArc, W, W, {}, 0, 0, W / 2);
  assert.ok(Math.abs(Number(arc.adj1.split(' ')[1]) - 10800000) < 60000, JSON.stringify(arc));
  // Every preset's handles are readable (the ECMA file's ahLst), each at a number.
  let n = 0;
  for (const [name] of SHAPE_GALLERY) for (const p of adjustHandles(PRESETS[name], W, H, {})) { n++; assert.ok(Number.isFinite(p.x) && Number.isFinite(p.y), name); }
  assert.ok(n > 60, `${n} handles in the gallery's shapes`);
});
