import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readPPr, readRPr, pPrToCss, rPrToCss, readBorders, borderCss, shdFill, fontStack, highlightColor } from '../src/ooxml/props.js';
import { fromStrict } from '../src/ooxml/strict.js';
import { parseXml, kid } from '../src/ooxml/xml.js';
import { strToU8, strFromU8 } from 'fflate';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const el = (xml) => parseXml(`<w:x ${W}>${xml}</w:x>`).documentElement;
// What a crafted file tries to slip into CSS: closing the rule or declaration, or fetching from elsewhere.
const ESCAPES = /[{}]|url\(|;\s*background|expression|<\/?style/i;
const evil = 'left}*{background:url(https://example.invalid/x)}x{';

test('a document\'s values can\'t write CSS of their own', () => {
  const p = readPPr(kid(el(`<w:pPr><w:jc w:val="${evil}"/><w:shd w:fill="${evil}"/></w:pPr>`), 'pPr'));
  assert.doesNotMatch(pPrToCss(p), ESCAPES);
  const r = readRPr(kid(el(`<w:rPr><w:color w:val="000;background:url(https://example.invalid/c)"/><w:highlight w:val="${evil}"/><w:rFonts w:ascii="Arial&quot;;background:url(https://example.invalid/f);x:&quot;"/></w:rPr>`), 'rPr'));
  // (The font name stays inside its quoted string, which it can't close: what's in it is a name, not CSS.)
  const css = rPrToCss(r, {});
  assert.match(css, /font-family:"[^"]*", sans-serif;/);
  assert.doesNotMatch(css.replace(/"[^"]*"/g, '""'), ESCAPES);
  const b = readBorders(el(`<w:top w:val="single" w:sz="4" w:color="${evil}"/>`));
  assert.doesNotMatch(borderCss(b.top), ESCAPES);
  assert.equal(shdFill(kid(el(`<w:shd w:fill="${evil}"/>`), 'shd')), null);
  assert.doesNotMatch(fontStack('A"}body{background:url(x)}').replace(/"[^"]*"/g, '""'), ESCAPES);
  assert.equal(highlightColor(evil), 'transparent');
  // Ordinary values still work.
  assert.equal(pPrToCss(readPPr(kid(el('<w:pPr><w:jc w:val="center"/></w:pPr>'), 'pPr'))), 'text-align:center');
  assert.match(rPrToCss(readRPr(kid(el('<w:rPr><w:color w:val="C00000"/><w:highlight w:val="yellow"/></w:rPr>'), 'rPr')), {}), /color:#C00000;background-color:#ffff00/);
});

test('a crafted Strict part full of unclosed tags is read in moments', () => {
  const S = 'http://purl.oclc.org/ooxml/wordprocessingml/main';
  const t0 = performance.now();
  fromStrict({ 'word/document.xml': strToU8(`<x xmlns:w="${S}">` + '<w:a '.repeat(200000) + '</x>') });
  assert.ok(performance.now() - t0 < 1000, 'under a second');
  assert.match(strFromU8(fromStrict({ 'word/document.xml': strToU8(`<w:document xmlns:w="${S}"><w:tblW w:w="50%" w:type="pct"/></w:document>`) })['word/document.xml']), /w:w="2500"/);
});

test('Strict charts\' percentages are plain numbers', () => {
  const C = 'http://purl.oclc.org/ooxml/drawingml/chart';
  const out = strFromU8(fromStrict({ 'word/charts/chart1.xml': strToU8(`<c:chartSpace xmlns:c="${C}"><c:gapWidth val="150%"/><c:overlap val="-20%"/></c:chartSpace>`) })['word/charts/chart1.xml']);
  assert.match(out, /<c:gapWidth val="150"\/><c:overlap val="-20"\/>/);
});
