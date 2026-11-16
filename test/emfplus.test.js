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
  assert.match(svg, /<linearGradient id="g\d+" gradientUnits="userSpaceOnUse" x1="220" y1="230" x2="380" y2="230" spreadMethod="repeat"><stop offset="0" stop-color="#ffd700"/, 'a gradient');
  assert.match(svg, /M20 280 C60 200 140 200 180 280Z/, 'a path with a curve');
  assert.match(svg, /font-family="&quot;Arial&quot;[^"]*" font-size="16" font-weight="bold"[^>]*><tspan x="232.667" y="129.48">EMF\+ text</, 'text, its family name no longer in capitals');
  assert.match(svg, /transform="matrix\(0\.8660254 -0\.5 0\.5 0\.8660254 200 250\)"><tspan[^>]*>Turned</, 'text turned by the world transform');
  const pic = /<g transform="matrix\(5 0 0 5 150 100\)"><use href="#(i\d+)" style="image-rendering:pixelated"\/>/.exec(svg);
  assert.ok(pic, 'a bitmap, unsmoothed');
  assert.match(svg, new RegExp(`<defs>.*<image id="${pic[1]}" width="8" height="8"[^>]*href="data:image/png;base64,[^"]+"/>.*</defs>`), 'its picture defined once');
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

// GDI+'s own pictures of the cases a review found (odx-test/mkemfplus2.ps1), checked against the records GDI+ wrote and
// its own geometry.
test('EMF+ pies and arcs end on their angle\'s ray from the centre, as GDI+ draws them', () => {
  const svg = emfToSvg(fixture('emfplus-piearc.emf'));
  // (A 45-degree pie on a 200 x 100 ellipse centred at (110, 60): its edge where the 45-degree ray meets it.)
  assert.match(svg, /<path d="M110 60L210 60A100 50 0 0 1 154\.721 104\.721Z" fill="#4472c4"/);
});

test('EMF+ containers: transforms set inside one are relative to it', () => {
  const svg = emfToSvg(fixture('emfplus-container.emf'));
  assert.match(svg, /fill="#ed7d31"[^>]*transform="matrix\(1 0 0 1 100 220\)"/, 'SetWorldTransform inside a scaled container');
});

test('EMF+ regions and clips: an intersection, and a clip moved by OffsetClip', () => {
  const svg = emfToSvg(fixture('emfplus-clip.emf'));
  assert.match(svg, /<clipPath id="p\d+"><path d="M400 120 C400 147\.614/, 'the region\'s circle clips its rectangle');
  assert.match(svg, /<path d="M20 240h40v40h-40Z" clip-rule="nonzero" transform="matrix\(1 0 0 1 60 0\)"/, 'the clip moved 60 to the right');
});

// Settled against Windows' own GDI+ playing the pictures (emfplus-word2/WORD-CHECK.md).
test('EMF+ font sizes: a point size ignores the page scale; a pixel size follows it as world units do', () => {
  const svg = emfToSvg(fixture('emfplus-pagepoint.emf'));
  assert.match(svg, /font-size="6"[^>]*transform="matrix\(2\.666667 0 0 2\.666667 0 0\)"><tspan[^>]*>12pt x2/, '12pt under PageScale 2: still 16px');
  assert.match(svg, /font-size="12"[^>]*transform="matrix\(2\.666667 0 0 2\.666667 0 0\)"><tspan[^>]*>12px x2/, '12px under PageScale 2: 32px');
});

test('EMF+ Display units of a printer reference are 1/100 inch', () => {
  const svg = emfToSvg(fixture('emfplus-dpi600.emf'));
  assert.match(svg, /<path d="M10 10h120v80h-120Z" fill="#4472c4"[^>]*transform="matrix\(6 0 0 6 0 0\)"/);
});
