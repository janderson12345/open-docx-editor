import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { OMF, OMF_RANGES, OMF_VERSION } from '../src/fonts/omf.js';
import { OMF_CDN } from '../src/index.js';
import { fontStack, lineFactor, setOmf } from '../src/ooxml/props.js';

test('open-metric-fonts: every Microsoft family it stands in for, with its faces and their alphabets', () => {
  assert.match(OMF_VERSION, /^\d+\.\d+\.\d+$/);
  assert.equal(OMF_CDN, 'https://cdn.jsdelivr.net/npm/@open-metric-fonts/');
  for (const name of ['Aptos', 'Aptos SemiBold', 'Aptos Narrow', 'Aptos Serif', 'Aptos Mono', 'Grandview', 'Seaford', 'Tenorite', 'Bahnschrift', 'Sitka Text', 'Century Gothic', 'Garamond', 'Franklin Gothic Book', 'Verdana Pro', 'Walbaum']) {
    assert.ok(OMF[name], name);
  }
  assert.ok(Object.keys(OMF).length >= 150);
  for (const [name, [pkg, line, faces]] of Object.entries(OMF)) {
    assert.match(pkg, /^[a-z0-9-]+$/, name);
    assert.ok(line > 0.9 && line < 2.2, `${name}: line ${line}`);
    assert.ok(faces.length >= 1 && faces.length <= 4, `${name}: ${faces.length} faces`);
    for (const [weight, italic, stem, alphabets] of faces) {
      assert.ok(weight === 400 || weight === 700, `${name}: weight ${weight}`);
      assert.ok(italic === 0 || italic === 1);
      assert.match(stem, /^[a-z0-9-]+$/);
      // Every alphabet's characters are known (its own, or the package's for its "other" file).
      for (const a of alphabets.split(',')) assert.ok(OMF_RANGES[`${pkg}:${a}`] || OMF_RANGES[a], `${name}: ${a}`);
    }
    // (One face for each weight and style: no two Regulars.)
    assert.equal(new Set(faces.map(([w, i]) => `${w}${i}`)).size, faces.length, name);
  }
  // Word links a weight family's own faces: Aptos SemiBold's Regular is Morrow Sans SemiBold, and it has no Bold.
  assert.deepEqual(OMF['Aptos SemiBold'][2].map(([w, i, s]) => [w, i, s]), [[400, 0, 'morrowsans-semibold'], [400, 1, 'morrowsans-semibolditalic']]);
});

test('a font missing from the computer: the open stand-in after the font and the editor\'s own, and Word\'s line height', () => {
  // Off (fonts.omf false, or outside a browser): stacks as they were.
  setOmf(null);
  assert.equal(fontStack('Grandview'), '"Grandview", sans-serif');
  const asked = [];
  setOmf((name) => { asked.push(name); return `odx-omf ${name}`; });
  assert.equal(fontStack('Grandview'), '"Grandview", "odx-omf Grandview", sans-serif');
  // Aptos keeps the editor's own stand-in first (fitted to Word's prints); the open one draws the letters it lacks.
  assert.match(fontStack('Aptos'), /^"Aptos", "odx-aptos", "odx-omf Aptos", sans-serif$/);
  // A font named in another case, and one it doesn't cover.
  assert.match(fontStack('GRANDVIEW'), /"odx-omf Grandview"/);
  assert.equal(fontStack('Comic Sans MS'), '"Comic Sans MS", sans-serif');
  assert.ok(asked.includes('Grandview') && asked.includes('Aptos') && !asked.includes('Comic Sans MS'));
  setOmf(null);
  // Line heights as Word sets them (the open fonts copy the originals' metrics): the editor's own table first.
  assert.equal(lineFactor('Grandview'), OMF.Grandview[1]);
  assert.equal(lineFactor('Bahnschrift'), 1.2002);
  assert.equal(lineFactor('Aptos'), 1.2207);
  assert.equal(lineFactor('Trebuchet MS'), 1.1621); // (the table's, fitted to Word, not the computed 1.1611)
  assert.equal(lineFactor('Calibri'), 1.2207);
});
