import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { formatNumber } from '../src/ooxml/styles.js';

// Word's own labels for items 1-24 and 98-101 of each format, read from Word's PDF of test/fixtures/lists-more.docx
// (mklists2.mjs).
const word = JSON.parse(fs.readFileSync(new URL('./fixtures/lists-more.word-labels.json', import.meta.url), 'utf8'));

test('numbering formats beyond decimal, letters and Roman give Word\'s labels', () => {
  let n = 0;
  for (const [fmt, items] of Object.entries(word)) {
    for (const [i, label] of Object.entries(items)) { assert.equal(`${formatNumber(Number(i), fmt)}.`, label, `${fmt} ${i}`); n++; }
  }
  assert.ok(n > 300);
});
