import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { searchPattern, replacementFor } from '../src/editor/find.js';

// A replacement as String.prototype.replace makes it, to compare with.
const native = (text, re, rep) => text.replace(new RegExp(re.source, re.flags.replace('g', '')), rep);
const ours = (text, query, rep, opts) => {
  const re = searchPattern(query, opts);
  const m = re.exec(text);
  return text.slice(0, m.index) + replacementFor({ groups: opts.regex ? [...m] : null, named: m.groups || null }, rep) + text.slice(m.index + m[0].length);
};

test('a plain search finds the text as typed, special characters and all', () => {
  assert.ok(searchPattern('a.b (c)').test('x a.b (c) y'));
  assert.ok(!searchPattern('a.b').test('axb'));
  assert.ok(searchPattern('Word', { wholeWord: true }).test('in Word.'));
  assert.ok(!searchPattern('Word', { wholeWord: true }).test('Wordy'));
  assert.ok(!searchPattern('word', { matchCase: true }).test('Word'));
});

test('a pattern\'s replacement fills in groups as String.replace does', () => {
  const text = 'Invoice 2024-03 paid';
  const cases = [
    ['(\\d{4})-(\\d{2})', '$2/$1'],
    ['(\\d{4})-(\\d{2})', '[$&]'],
    ['(\\d{4})-(\\d{2})', '$$1'],
    ['(?<y>\\d{4})-(?<m>\\d{2})', '$<m>/$<y>'],
    ['(\\d{4})', '$10'],
    ['(\\d{4})-(\\d{2})', '$9'],
    ['(\\d{4})-(\\d{2})', 'x$'],
  ];
  for (const [q, rep] of cases) assert.equal(ours(text, q, rep, { regex: true }), native(text, searchPattern(q, { regex: true }), rep), `${q} -> ${rep}`);
  // Without patterns a replacement is taken as it is.
  assert.equal(ours('cost: 5', 'cost', '$1', {}), '$1: 5');
});

test('a broken pattern is a SyntaxError (the find bar says so)', () => {
  assert.throws(() => searchPattern('(unclosed', { regex: true }), SyntaxError);
});

test('a phrase is found across comment ranges, comment references and bookmarks', async () => {
  const { findAll } = await import('../src/editor/find.js');
  const { schema } = await import('../src/editor/schema.js');
  const mark = (kind) => schema.nodes.opaque_inline.create({ kind, xml: `<w:${kind} w:id="1"/>` });
  const p = schema.nodes.paragraph.create(null, [mark('commentRangeStart'), schema.text('An Income Strategy'), mark('commentRangeEnd'), mark('commentReference'), schema.text(' '), mark('bookmarkStart'), schema.text('Approach'), mark('bookmarkEnd')]);
  const doc = schema.nodes.doc.create(null, [p]);
  const m = findAll(doc, 'Strategy Approach');
  assert.equal(m.length, 1);
  assert.equal(doc.textBetween(m[0].from, m[0].to, '', ''), 'Strategy Approach');
});
