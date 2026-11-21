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
});

test('a whole line takes the word on when narrowing its spaces costs each no more than half of what leaving it out would widen each by', () => {
  // (why-tests gen2.py justfit, justsp, justnar: Word's prints, 3,949 words that stuck out of a line. The two nearest
  // the rule in plain text: 11.77px over 23 spaces of 14.67px Calibri, a 31.02px word, taken on at 2.005 times; 10.45px
  // over 14 spaces of 16px Times New Roman, a 25.77px word, left at 1.993 times.)
  assert.equal(squeezeFits(11.768, 23, 76.294, 23 * 14.667, 31.017, false, 3.317, 22), true);
  assert.equal(squeezeFits(10.445, 14, 56, 14 * 16, 25.773, false, 4, 13), false);
  // The quarter's limit still holds, whatever the widening saved.
  assert.equal(squeezeFits(0.26 * 5 * sp, 5, 5 * sp, 5 * em, 60, false, sp, 4), false);
  // A line of one word before it can't be widened: the squeeze, when its one space allows it.
  assert.equal(squeezeFits(0.8, 1, sp, em, 30, false, sp, 0), true);
  // After a hyphen no space is spent on the word: the line without it has as many spaces to widen.
  assert.equal(squeezeFits(6, 10, 10 * sp, 10 * em, 20, false, 0, 10), true);
  assert.equal(squeezeFits(8, 10, 10 * sp, 10 * em, 20, false, 0, 10), false);
});

test('a split side squeezes its spaces by a quarter at most', () => {
  // Word's most on a side, 24.9% of the spaces; its least left off, 25.4%.
  assert.equal(squeezeFits(0.249 * 6 * sp, 6, 6 * sp, 6 * em, 40, true), true);
  assert.equal(squeezeFits(0.254 * 4 * sp, 4, 4 * sp, 4 * em, 30, true), false);
  assert.equal(squeezeFits(1, 0, 0, 0, 30, true), false);
});
