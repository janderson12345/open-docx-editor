import { test } from 'node:test';
import assert from 'node:assert/strict';
import { borderCss } from '../src/ooxml/props.js';

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
