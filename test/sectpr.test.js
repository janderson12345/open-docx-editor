import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as sp from '../src/ooxml/sectpr.js';

const LETTER = sp.DEFAULT_SECTPR;

test('margins are set on w:pgMar, the other attributes kept', () => {
  const x = sp.setMargins(LETTER, 720, 720, 720, 720);
  assert.match(x, /<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"\/>/);
});

test('landscape swaps the sides and marks the orientation; portrait swaps back and drops it', () => {
  const land = sp.setOrientation(LETTER, true);
  assert.deepEqual(sp.pageSize(land), { w: 15840, h: 12240 });
  assert.match(land, /w:orient="landscape"/);
  const port = sp.setOrientation(land, false);
  assert.deepEqual(sp.pageSize(port), { w: 12240, h: 15840 });
  assert.doesNotMatch(port, /w:orient/);
  assert.equal(sp.setOrientation(LETTER, false), LETTER, 'already portrait: unchanged');
});

test('a paper size keeps a landscape page landscape', () => {
  assert.deepEqual(sp.pageSize(sp.setPageSize(LETTER, 11906, 16838)), { w: 11906, h: 16838 });
  assert.deepEqual(sp.pageSize(sp.setPageSize(sp.setOrientation(LETTER, true), 11906, 16838)), { w: 16838, h: 11906 });
});

test('columns: added in their place in w:sectPr, replaced when there, one column drops w:num', () => {
  const titled = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:titlePg/></w:sectPr>';
  const two = sp.setColumns(titled, 2);
  assert.equal(two, '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:cols w:space="720" w:num="2"/><w:titlePg/></w:sectPr>');
  assert.equal(sp.columnCount(two), 2);
  const three = sp.setColumns(two, 3);
  assert.equal(sp.columnCount(three), 3);
  assert.equal(three.match(/<w:cols/g).length, 1);
  const one = sp.setColumns('<w:sectPr><w:cols w:num="2"><w:col w:w="4000"/><w:col w:w="4000"/></w:cols></w:sectPr>', 1);
  assert.equal(one, '<w:sectPr><w:cols w:space="720"/></w:sectPr>');
  assert.equal(sp.columnCount(one), 1);
  assert.equal(sp.setColumns(LETTER, 2), LETTER.replace('</w:sectPr>', '<w:cols w:space="720" w:num="2"/></w:sectPr>'));
});

test('a missing page size reads as Letter', () => {
  assert.deepEqual(sp.pageSize('<w:sectPr/>'), { w: 12240, h: 15840 });
  assert.equal(sp.columnCount(null), 1);
});
