import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { VERSION, FONTS_CDN } from '../src/index.js';
import { fontStack, setFontSubstitutes } from '../src/ooxml/props.js';

test('src/version.js is the package\'s version, and the stand-ins load from that release on jsDelivr', () => {
  assert.equal(VERSION, JSON.parse(fs.readFileSync(new URL('../package.json', import.meta.url), 'utf8')).version);
  assert.equal(FONTS_CDN, `https://cdn.jsdelivr.net/npm/open-docx-editor@${VERSION}/src/fonts/`);
  // The files it names are in the package (files: src) and in fonts.css.
  const css = fs.readFileSync(new URL('../src/fonts/fonts.css', import.meta.url), 'utf8');
  for (const m of css.matchAll(/url\('\.\/files\/([^']+)'\)/g)) assert.ok(fs.existsSync(new URL(`../src/fonts/files/${m[1]}`, import.meta.url)), m[1]);
  assert.match(css, /font-family: 'odx-aptos'; src: url\('\.\/files\/odx-aptos-400-normal\.woff2'\)/);
});

test('a host\'s substitute comes after the font itself and before the editor\'s stand-ins', () => {
  setFontSubstitutes(new Map([['Grandview', 'Inter'], ['Aptos', 'Source Sans 3']]));
  assert.match(fontStack('Grandview'), /^"Grandview", "Inter", sans-serif$/);
  assert.match(fontStack('Aptos'), /^"Aptos", "Source Sans 3", "odx-aptos"/);
  assert.match(fontStack('Calibri'), /^"Calibri", "Carlito"/);
});
