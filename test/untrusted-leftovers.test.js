import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { legacyInfo } from '../src/ooxml/ffdata.js';
import { bevelFx, textEffectsCss } from '../src/ooxml/texteffects.js';

// Crafted values from a file, bounded (the security review of the leftovers after 0.8.0).

test('a form field read in one pass however many unclosed tags a comment holds', () => {
  const junk = '<w:maxLength '.repeat(40000);
  const xml = `<w:ffData><w:name w:val="F"/><!-- ${junk} --><w:textInput><w:maxLength w:val="5"/></w:textInput></w:ffData>`;
  const t0 = Date.now();
  const info = legacyInfo(xml);
  assert.ok(Date.now() - t0 < 1000);
  assert.equal(info.kind, 'text');
  assert.equal(legacyInfo('<w:ffData><w:textInput><w:maxLength w:val="7"/></w:textInput></w:ffData>').maxLength, 7);
  assert.equal(legacyInfo('<w:ffData><w:textInput/></w:ffData>').maxLength, null);
});

test("a bevel's size kept to 64px however large the file says", () => {
  const big = '<w14:props3d w14:prstMaterial="matte"><w14:bevelT w14:w="999999999999" w14:h="999999999999" w14:prst="angle"/></w14:props3d>';
  const out = bevelFx(big);
  assert.ok(out);
  for (const m of out.matchAll(/(?:radius|stdDeviation|surfaceScale)="([\d.\s]+)"/g)) {
    for (const n of m[1].trim().split(/\s+/)) assert.ok(Number(n) <= 64, `${m[0]} past 64`);
  }
});

test("a reflection's blur bounded", () => {
  const x = '<w14:reflection w14:blurRad="999999999999" w14:stA="55000" w14:stPos="0" w14:endA="300" w14:endPos="45500" w14:dist="0" w14:dir="5400000" w14:fadeDir="5400000" w14:sx="100000" w14:sy="-100000" w14:kx="0" w14:ky="0" w14:algn="bl"/>';
  const css = textEffectsCss([['reflection', x]]).join(';');
  const m = /--odx-rf-blur:([\d.]+)pt/.exec(css);
  assert.ok(m, css);
  assert.ok(Number(m[1]) <= 20);
});
