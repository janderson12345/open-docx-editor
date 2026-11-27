import { test } from 'node:test';
import assert from 'node:assert/strict';
import { fontMetrics, eastMetrics } from '../src/fonts/wordmetrics.js';
import { lineFactor } from '../src/ooxml/props.js';
import { ASCENT, wordAscent } from '../src/editor/baseline.js';

test('Word\'s line and ascent come from the fonts\' own numbers, as measured in Word\'s prints', () => {
  // Calibri: hhea and Windows totals equal (1.2207); hhea's ascender + line gap 0.9707, Word's mixed-line ascent its
  // Windows ascent 0.9521 and its Windows descent 0.2686 (larger than hhea's 0.25).
  assert.deepEqual(fontMetrics('calibri'), { line: 1.2207, ascent: 0.9707, wordAscent: 0.9521, winDescent: 0.2686 });
  // Arial: hhea's total is the larger, the difference above the Windows ascent.
  assert.equal(fontMetrics('arial').line, 1.1499);
  assert.equal(fontMetrics('arial').wordAscent, 0.938);
  // Corbel: the Windows total is the larger (its hhea total is 1.2075).
  assert.equal(fontMetrics('corbel').line, 1.2207);
  // Aptos asks for its typo metrics (its Windows total is 1.2847).
  assert.equal(fontMetrics('aptos').line, 1.2207);
  // Stand-ins take their original's numbers; unknown names (and prototype names) have none.
  assert.deepEqual(fontMetrics('carlito'), fontMetrics('calibri'));
  assert.equal(fontMetrics('book antiqua x'), null);
  assert.equal(fontMetrics('constructor'), null);
  // MS Gothic: 220 and 36 of 256, each 1.3 times, cut to a whole unit.
  assert.deepEqual(eastMetrics('ms gothic'), [258 / 256, 332 / 256]);
  assert.equal(eastMetrics('arial'), null);
  // The tables that use them.
  assert.equal(lineFactor('Calibri'), 1.2207);
  assert.equal(ASCENT.calibri, 0.9707);
  assert.equal(wordAscent('calibri'), 0.9521);
  assert.equal(wordAscent('corbel'), undefined);
});
