import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { pastedHeading } from '../src/editor/schema.js';

// A pasted <p> as pastedHeading reads it: its class and style attributes, and the CSS the browser keeps of the style.
const p = (attrs = {}, style = {}) => ({ getAttribute: (k) => attrs[k] ?? null, style });

test('pasted <p> headings: Word\'s clipboard HTML, heading classes, role=heading, bold at a heading\'s size', () => {
  assert.equal(pastedHeading(p({ class: 'MsoHeading1' })), 'Heading1');
  assert.equal(pastedHeading(p({ class: 'MsoNormal', style: 'mso-style-name:"Heading 2";margin:0' })), 'Heading2');
  assert.equal(pastedHeading(p({ style: "mso-style-name: ' heading3 ' ;" })), 'Heading3');
  assert.equal(pastedHeading(p({ style: 'mso-style-name:"Heading 1 Char"' })), null);
  assert.equal(pastedHeading(p({ style: 'mso-outline-level:4' })), 'Heading4');
  assert.equal(pastedHeading(p({ class: 'x heading-2' })), 'Heading2');
  assert.equal(pastedHeading(p({ role: 'heading', 'aria-level': '5' })), 'Heading5');
  assert.equal(pastedHeading(p({}, { fontSize: '2em', fontWeight: 'bold' })), 'Heading1');
  assert.equal(pastedHeading(p({}, { fontWeight: 'bold' })), null);
  assert.equal(pastedHeading(p({ class: 'MsoNormal' })), null);
});

test('pasted <p> headings: a style with a long run of spaces is read in linear time', () => {
  const sp = ' '.repeat(40000);
  const t0 = performance.now();
  for (const style of [`mso-style-name:${sp}x`, `mso-style-name:"heading 1${sp}x`, `mso-style-name:heading 1${sp}"${sp}x`]) assert.equal(pastedHeading(p({ style })), null);
  assert.ok(performance.now() - t0 < 200, `${(performance.now() - t0).toFixed(0)}ms`); // (was 5.6s: 1.4s, 2.4s, 1.8s)
});
