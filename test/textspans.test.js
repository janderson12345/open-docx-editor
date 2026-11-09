import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { EditorState } from 'prosemirror-state';
import { schema } from '../src/editor/schema.js';
import { textSpansPlugin } from '../src/editor/textspans.js';

const { paragraph, hard_break: br } = schema.nodes;
const para = (...kids) => paragraph.create(null, kids.map((k) => (typeof k === 'string' ? schema.text(k) : k)));
const make = (...paras) => EditorState.create({ doc: schema.nodes.doc.create(null, paras), plugins: [textSpansPlugin] });
const decos = (state) => textSpansPlugin.getState(state).find();
const spans = (state) => decos(state).filter((d) => d.spec.key !== 'odx-strut').map((d) => [d.from, d.to]);
const struts = (state) => decos(state).filter((d) => d.spec.key === 'odx-strut').map((d) => d.from);

test('every text node gets a span of its own', () => {
  const s = make(para('ab', br.create(), 'cd'), para('e'));
  // paragraph 1: "ab" at 1-3, the break at 3, "cd" at 4-6; paragraph 2 starts at 7: "e" at 8-9
  assert.deepEqual(spans(s), [[1, 3], [4, 6], [8, 9]]);
  assert.deepEqual(struts(s), []);
});

test('a line holding only a break gets a strut: a break right after a break', () => {
  const s = make(para('a', br.create(), br.create(), 'b'));
  assert.deepEqual(struts(s), [3]); // before the second break, the line it ends
});

test('spans follow edits: typing into a paragraph and marking part of it', () => {
  let s = make(para('hello'));
  s = s.apply(s.tr.insertText('XY', 3));
  assert.deepEqual(spans(s), [[1, 8]]);
  s = s.apply(s.tr.addMark(2, 4, schema.marks.bold.create()));
  // (one text node per mark run now: the decoration is redrawn over each)
  assert.deepEqual(spans(s), [[1, 2], [2, 4], [4, 8]]);
  s = s.apply(s.tr.split(5));
  assert.deepEqual(spans(s), [[1, 2], [2, 4], [4, 5], [7, 10]]);
});
