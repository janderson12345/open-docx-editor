import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { correction } from '../src/editor/autocorrect.js';

// Types `text` one character at a time through the corrections, as the editor does.
const type = (text, start = '') => {
  let s = start;
  for (const ch of text) {
    s += ch;
    const fix = correction(s);
    if (fix) s = s.slice(0, fix.from) + fix.text + s.slice(fix.to);
  }
  return s;
};

test('quotes turn curly, opening or closing by what comes before', () => {
  assert.equal(type('He said "it\'s fine" and (\'yes\')'), 'He said “it’s fine” and (‘yes’)');
  assert.equal(type('"Quoted"'), '“Quoted”');
});

test('symbols, ellipsis, arrows, dashes and fractions', () => {
  assert.equal(type('Acme(c) and (TM) and (r) '), 'Acme© and ™ and ® ');
  assert.equal(type('Wait... go --> there <-- back ==> on'), 'Wait… go → there ← back ⇒ on');
  assert.equal(type('well--maybe, a - b, c -- d'), 'well—maybe, a – b, c – d');
  assert.equal(type('add 1/2 cup and 3/4 '), 'Add ½ cup and ¾ ');
  assert.equal(type('a-b stays, 2-3 too'), 'A-b stays, 2-3 too');
});

test('a sentence\'s first word and a lone i are capitalised; abbreviations and initials are not sentence ends', () => {
  assert.equal(type('hello there. this is it! and so? yes '), 'Hello there. This is it! And so? Yes ');
  assert.equal(type('see e.g. the list, etc. and more '), 'See e.g. the list, etc. and more ');
  assert.equal(type('J. smith said i would '), 'J. smith said I would ');
  assert.equal(type('iPhone stays '), 'iPhone stays ');
});

test('common misspellings, keeping a capital', () => {
  assert.equal(type('Teh cat adn teh dog dont recieve '), 'The cat and the dog don’t receive ');
  assert.equal(type('the cathedral stays '), 'The cathedral stays ');
});

test('words the built-in object has are left alone; "i" only on its own; a fixed word starting a sentence gets its capital', () => {
  assert.equal(type('the constructor toString valueOf hasOwnProperty Constructor '), 'The constructor toString valueOf hasOwnProperty Constructor ');
  assert.equal(type('use i.e. this, and i) that; i think '), 'Use i.e. this, and i) that; I think ');
  assert.equal(type('teh end. dont stop '), 'The end. Don’t stop ');
  assert.equal(type('left <== right '), 'Left ⇐ right ');
});
