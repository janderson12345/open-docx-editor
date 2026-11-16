import { test } from 'node:test';
import assert from 'node:assert/strict';
import { decode, covers, withFonts, shownSrc } from '../src/editor/svgfonts.js';

// Shapes' pictures get the stand-in faces their text needs embedded (src/editor/svgfonts.js). The parts that need no
// page: reading a picture's data: URL, and whether a face's unicode-range takes the picture's text.
test('decode reads a data: URL, base64 or percent-encoded', () => {
  const svg = '<svg xmlns="http://www.w3.org/2000/svg"><text font-family="Aptos">Hé</text></svg>';
  assert.equal(decode(`data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`), svg);
  assert.equal(decode(`data:image/svg+xml;base64,${Buffer.from(svg, 'utf8').toString('base64')}`), svg);
  assert.equal(decode('not a data url'), null);
  assert.equal(decode('data:image/svg+xml,%E0%A4%A'), null); // (a broken escape: none)
});

test('covers: a unicode-range takes the text when any of its characters is in it', () => {
  assert.equal(covers(null, 'anything'), true);
  assert.equal(covers('U+0000-00FF', 'Hello'), true);
  assert.equal(covers('U+0370-03FF', 'Hello'), false);
  assert.equal(covers('U+0370-03FF, U+1F00-1FFF', 'Hello αβ'), true);
  assert.equal(covers('U+4??', 'Ж'), true); // (U+0400-04FF)
  assert.equal(covers('U+0000-00FF', '   '), false); // (spaces and controls don't count)
});

test('without a page, pictures are shown as drawn', async () => {
  const src = 'data:image/svg+xml,%3Csvg%3E%3C/svg%3E';
  assert.equal(shownSrc(src), src);
  assert.equal(await withFonts('data:image/svg+xml,%3Csvg%3E%3C/svg%3E'), null); // (no text: nothing to embed)
});
