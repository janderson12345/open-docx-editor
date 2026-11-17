import { test } from 'node:test';
import assert from 'node:assert/strict';
import { borderCss, pPrToCss, compoundSide, junctionRects, cellLines } from '../src/ooxml/props.js';

const width = (val, sz) => parseFloat(borderCss({ val, sz, color: '1F497D' }));

test('a border of two or three lines is as wide as Word lays it out: its lines and gaps', () => {
  // (Word's PDFs of an empty header paragraph bordered below, hdrgap-word: the body that much lower, in its 0.12pt steps.)
  const word = {
    thickThinSmallGap: [[6, 2.16], [12, 3.0], [18, 3.72], [24, 4.44], [36, 6.0], [48, 7.44]],
    thinThickSmallGap: [[18, 3.72]],
    thickThinMediumGap: [[12, 3.0], [18, 4.44], [36, 9.0]],
    thickThinLargeGap: [[6, 3.0], [18, 4.44], [36, 6.72]],
    thinThickLargeGap: [[18, 4.44]],
    thinThickThinSmallGap: [[6, 3.72], [18, 5.16], [36, 7.44]],
    thinThickThinMediumGap: [[18, 6.72]],
    thinThickThinLargeGap: [[6, 4.44], [18, 7.44], [36, 12.0]],
    triple: [[6, 3.72], [18, 11.16]],
    double: [[6, 2.16]],
  };
  for (const [val, sizes] of Object.entries(word)) {
    for (const [sz, pt] of sizes) assert.ok(Math.abs(width(val, sz) - pt) <= 0.13, `${val} ${sz}: ${width(val, sz)}pt, Word ${pt}`);
    assert.match(borderCss({ val, sz: 18 }), / double /);
  }
  assert.equal(width('single', 18), 2.25);
  // (Not a name from the object's prototype: an ordinary line.)
  assert.equal(borderCss({ val: 'constructor', sz: 8 }), '1pt solid #000');
});

test('a paragraph\'s compound border thinner than CSS double can draw: its lines as Word prints them, at Word\'s width', () => {
  // (thinbdr-word: boxes of thin triple, medium-gap and double borders in Word's PDF; the lines below each box where Word
  // puts them, within its 0.12pt steps.)
  const box = (val, sz) => Object.fromEntries(['top', 'bottom', 'left', 'right'].map((s) => [s, { val, sz, color: '1F497D', space: s === 'top' || s === 'bottom' ? 1 : 4 }]));
  const css = pPrToCss({ borders: box('triple', 2) });
  const svg = decodeURIComponent(/border-image:url\('data:image\/svg\+xml,([^']*)'\) 10 stretch/.exec(css)[1]);
  // Triple 0.25pt: three 0.25pt lines 0.25pt apart, 1.25pt (1.67px) in all, drawn 2px wide: 1px of space less.
  assert.match(css, /border-top:1\.5pt solid #1F497D;padding-top:1px;/);
  assert.equal(+/padding-top:([\d.]+)px/.exec(css)[1], +(4 / 3 + 5 / 3 - 2).toFixed(3));
  assert.equal(+/padding-bottom:([\d.]+)px/.exec(css)[1], +(4 / 3 + 5 / 3 - 2).toFixed(3));
  // Its side lines outside the indent: their width and space taken off it (--odx-bdl).
  assert.match(css, /border-left:1\.5pt solid #1F497D;padding-left:4pt;--odx-bdl:5\.5pt/);
  // The top band's three lines, a fifth of its width each, and each corner's three nested Ls.
  assert.equal((svg.match(/<rect x="10" y="(0|4|8)" width="10" height="2"/g) || []).length, 3);
  assert.match(svg, /<rect x="4" y="4" width="6" height="2" fill="#1F497D"\/>/); // (top-left, the middle line's arm)
  assert.match(svg, /<rect x="20" y="4" width="6" height="2" fill="#1F497D"\/>/); // (top-right)
  // Medium gap 0.25pt: its size line first in the page's order on every side, half as much gap and thin line after.
  const med = decodeURIComponent(/svg\+xml,([^']*)'/.exec(pPrToCss({ borders: box('thinThickMediumGap', 2) }))[1]);
  assert.match(med, /<rect x="10" y="0" width="10" height="5"[^>]*\/><rect x="10" y="7\.5" width="10" height="2\.5"/);
  assert.match(med, /<rect x="10" y="20" width="10" height="5"[^>]*\/><rect x="10" y="27\.5" width="10" height="2\.5"/);
  assert.match(med, /<rect x="0" y="10" width="5" height="10"[^>]*\/><rect x="7\.5" y="10" width="2\.5" height="10"/);
  // One wide enough for CSS double drawn the same way (a triple's three lines, not CSS's two), and a paragraph's own
  // borders aren't drawn by a style's picture.
  assert.equal((decodeURIComponent(/svg\+xml,([^']*)'/.exec(pPrToCss({ borders: box('triple', 24) }))[1]).match(/<rect x="10" y="(0|4|8)" width="10" height="2"/g) || []).length, 3);
  assert.match(pPrToCss({ borders: { top: { val: 'single', sz: 4 } } }), /border-image:none.*border-top:0\.5pt solid/);
  // A dotted side the picture can't draw: as before, at least 2.25pt.
  assert.doesNotMatch(pPrToCss({ borders: { ...box('triple', 2), left: { val: 'dotted', sz: 4 } } }), /svg/);
  // Only a colour goes into the picture.
  assert.match(decodeURIComponent(/svg\+xml,([^']*)'/.exec(pPrToCss({ borders: { bottom: { val: 'double', sz: 2, color: '"/><x' } } }))[1]), /fill="#000"/);
});

test("compound borders of 3pt to 6pt: their lines as Word prints them, a large gap's lines 1.5pt and 0.75pt", () => {
  // (bdr3-word: boxes of each kind at 3, 4.5 and 6pt in Word's PDF, in its 0.12pt steps.)
  const lines = (val, sz) => { const s = compoundSide({ val, sz }); return s.lines.map(([a, l]) => [+(a * s.px * 0.75).toFixed(2), +(l * s.px * 0.75).toFixed(2)]); };
  assert.deepEqual(lines('thickThinSmallGap', 24), [[0, 0.75], [1.5, 3]]);
  assert.deepEqual(lines('thinThickLargeGap', 24), [[0, 1.5], [4.5, 0.75]]);
  assert.deepEqual(lines('thickThinLargeGap', 36), [[0, 0.75], [5.25, 1.5]]);
  assert.deepEqual(lines('thinThickThinLargeGap', 48), [[0, 0.75], [6.75, 1.5], [14.25, 0.75]]);
  assert.deepEqual(lines('triple', 24), [[0, 3], [6, 3], [12, 3]]);
  assert.equal(compoundSide({ val: 'single', sz: 4 }), null);
  assert.equal(compoundSide({ val: 'toString', sz: 4 }), null);
});

test("where table borders of three lines meet: Word's cross, T and corner", () => {
  // (bdr3-word's tables, triple 0.25pt: a cross is its middle lines crossing and a dot where each quadrant's nearest
  // lines turn; a T on the top edge keeps its outer line; a corner nests its lines.)
  const t = compoundSide({ val: 'triple', sz: 2 }).lines;
  const draw = (arms) => junctionRects(t, t, arms).map(([x, y, w, h, k]) => `${k}${[x, y, w, h].map((v) => Math.round(v * 5)).join(',')}`).sort().join(' ');
  assert.equal(draw({ l: 1, r: 1, u: 1, d: 1 }), 'h0,0,1,1 h0,2,5,1 h0,4,1,1 h4,0,1,1 h4,4,1,1 v0,0,1,1 v0,4,1,1 v2,0,1,5 v4,0,1,1 v4,4,1,1');
  assert.equal(draw({ l: 1, r: 1, u: 0, d: 1 }), 'h0,0,5,1 h0,2,5,1 h0,4,1,1 h4,4,1,1 v0,4,1,1 v2,2,1,3 v4,4,1,1');
  assert.equal(draw({ l: 0, r: 1, u: 0, d: 1 }), 'h0,0,5,1 h2,2,3,1 h4,4,1,1 v0,0,1,5 v2,2,1,3 v4,4,1,1');
});

test("a table cell's compound lines: pictures on whole pixels, the same rects for the PDF, only numbers and colours", () => {
  const d = compoundSide({ val: 'double', sz: 24, color: '1F497D' }); // 9pt: 12px
  const thin = compoundSide({ val: 'triple', sz: 2, color: '"/><x' }); // 1.25pt: 1.67px
  const out = cellLines({ top: thin, left: d }, {});
  // A corner square, the top's strip after it, the left's strip under the square's whole pixels.
  assert.match(out.css, /background-position:left 0px top 0px,left 12px top 0px,left 0px top 2px;/);
  assert.match(out.css, /background-size:12px 2px,calc\(100% - 12px\) 2px,12px calc\(100% - 2px\);/);
  assert.doesNotMatch(out.css, /%3Cx/);
  // Every rect: eight numbers and a colour.
  for (const r of out.lines.split(',')) assert.match(r, /^(-?[\d.]+ ){8}#[0-9A-Fa-f]{6}$/);
  // The top's first line right across the cell from the square: x 12px, y 0, width the cell less 12px.
  assert.match(out.lines, /(^|,)0 12 0 0 1 -12 0 0\.333 #000000(,|$)/);
  assert.equal(cellLines({}, {}), null);
});
