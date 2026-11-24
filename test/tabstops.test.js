// A paragraph's tab stops as Word takes them (props.js mergeTabs): in rising order, a stop written after one further
// along dropped (why-tests gen2.py tabdup, Word's prints of 34 lines: stops written 900 then 880 twips take a tab to 900,
// written 880 then 900 to 880).
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mergeTabs } from '../src/ooxml/props.js';

const left = (...pos) => pos.map((p) => ({ val: 'left', pos: p, leader: null }));
const at = (tabs) => tabs.map((t) => t.pos);

test('a tab stop written after one further along is not a stop', () => {
  assert.deepEqual(at(mergeTabs(left(880, 900))), [880, 900]);
  assert.deepEqual(at(mergeTabs(left(900, 880))), [900]);
  assert.deepEqual(at(mergeTabs(left(900, 880, 2319))), [900, 2319]);
  assert.deepEqual(at(mergeTabs(left(880, 879))), [880]);
  assert.deepEqual(at(mergeTabs(left(900, 800, 850, 1200))), [900, 1200]);
});

test('the order is each list\'s own: a paragraph\'s stops still add to its style\'s, and clear them', () => {
  // A style's stop at 1440, the paragraph's at 720: both stand (two lists, each in order).
  assert.deepEqual(at(mergeTabs(left(1440), left(720))), [720, 1440]);
  // A clear takes the style's stop away wherever it is written, and is no stop itself.
  assert.deepEqual(at(mergeTabs(left(720, 1440), [{ val: 'left', pos: 2000 }, { val: 'clear', pos: 720 }])), [1440, 2000]);
});
