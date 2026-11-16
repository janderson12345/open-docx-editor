// Word's squeezing of justified lines (justify.js squeezeFits): a whole line's rule, and a side of a line split beside a
// shape (wrapboth.js), measured from Word's print of wrap-word3's documents (wrap-squeeze.docx is one).
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { squeezeFits } from '../src/editor/justify.js';

// (16px Calibri: a space 3.6px; "a" 7.7px.)
const sp = 3.6, em = 16;

test('a split side squeezes by its spaces\' limits alone: "a" sticking out by 44% of its width goes on', () => {
  // sq-none's Text 9: four spaces, "a" over the side's end by 3.41px (23.6% of the spaces).
  assert.equal(squeezeFits(3.41, 4, 4 * sp, 4 * em, 7.66, true), true);
  // A whole line keeps its rule: a word that far out stays off (MAX_OUT).
  assert.equal(squeezeFits(3.41, 4, 4 * sp, 4 * em, 7.66), false);
});

test('a split side squeezes its spaces by a quarter at most', () => {
  // Word's most on a side, 24.9% of the spaces; its least left off, 25.4%.
  assert.equal(squeezeFits(0.249 * 6 * sp, 6, 6 * sp, 6 * em, 40, true), true);
  assert.equal(squeezeFits(0.254 * 4 * sp, 4, 4 * sp, 4 * em, 30, true), false);
  assert.equal(squeezeFits(1, 0, 0, 0, 30, true), false);
});
