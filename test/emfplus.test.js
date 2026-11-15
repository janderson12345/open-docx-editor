import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { emfToSvg } from '../src/ooxml/emf.js';

// Both fixtures were drawn by Windows' GDI+ (odx-test/mkemfplus.ps1): the same picture as an EMF+ only file (no GDI
// records) and as an EMF+ dual file (with a GDI copy).
const fixture = (name) => new Uint8Array(readFileSync(new URL(`./fixtures/${name}`, import.meta.url)));

test('an EMF+ only picture is drawn from its EMF+ records', () => {
  const svg = emfToSvg(fixture('emfplus-only.emf'));
  assert.match(svg, /<path d="M10 10h120v80h-120Z" fill="#4472c4"/, 'a solid rectangle');
  assert.match(svg, /fill="#ed7d31" fill-opacity="0\.502"/, 'a half-transparent one');
  assert.match(svg, /M220 60a80 50 0 1 0 160 0.*stroke="#70ad47" stroke-width="4"/, 'an ellipse outlined by a 4px pen');
  assert.match(svg, /stroke="#8b0000" stroke-width="2" stroke-dasharray="6 2"/, 'a dashed pen');
  assert.match(svg, /<linearGradient id="g\d+" gradientUnits="userSpaceOnUse" x1="220" y1="230" x2="380" y2="230"><stop offset="0" stop-color="#ffd700"/, 'a gradient');
  assert.match(svg, /M20 280 C60 200 140 200 180 280Z/, 'a path with a curve');
  assert.match(svg, /font-family="&quot;Arial&quot;[^"]*" font-size="16" font-weight="bold"[^>]*><tspan x="232.667" y="129.48">EMF\+ text</, 'text, its family name no longer in capitals');
  assert.match(svg, /transform="matrix\(0\.8660254 -0\.5 0\.5 0\.8660254 200 250\)"><tspan[^>]*>Turned</, 'text turned by the world transform');
  assert.match(svg, /<g transform="matrix\(5 0 0 5 150 100\)"><image width="8" height="8"[^>]*href="data:image\/png;base64,[^"]+" style="image-rendering:pixelated"/, 'a bitmap, unsmoothed');
});

test('an EMF+ dual picture is drawn from its EMF+ records too, its GDI copy left out', () => {
  const svg = emfToSvg(fixture('emfplus-dual.emf'));
  assert.match(svg, /fill="#ed7d31" fill-opacity="0\.502"/, 'transparency the GDI copy lacks');
  assert.doesNotMatch(svg, /<text x=/, 'none of the GDI copy\'s text');
});

test('a dual picture whose EMF+ records need what is not read is drawn from its GDI copy', () => {
  const b = fixture('emfplus-dual.emf');
  // Turn the first EMF+ FillRects record (0x400A) into an unknown drawing record (0x403F).
  const dv = new DataView(b.buffer);
  for (let i = 0; i + 2 < b.length; i += 2) if (dv.getUint16(i, true) === 0x400a) { dv.setUint16(i, 0x403f, true); break; }
  const svg = emfToSvg(b);
  assert.match(svg, /<text x=/, 'the GDI copy\'s text');
  assert.doesNotMatch(svg, /fill-opacity="0\.502"/);
});
