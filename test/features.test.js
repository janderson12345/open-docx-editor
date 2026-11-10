import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { DocxSession } from '../src/ooxml/docx.js';
import { listCounter } from '../src/ooxml/styles.js';
import { isStrict, fromStrict } from '../src/ooxml/strict.js';
import { lineNumbering } from '../src/editor/linenumbers.js';
import { isTiff, tiffToPng } from '../src/ooxml/tiff.js';

const load = (f) => new Uint8Array(fs.readFileSync(new URL(`./fixtures/${f}`, import.meta.url)));
const text = (doc) => { let t = ''; doc.descendants((n) => { if (n.isText) t += n.text; }); return t; };

test('legal numbering shows every level in Arabic numerals, its own included', () => {
  const s = new DocxSession(load('numbering-legal.docx'));
  const next = listCounter(s.numbering);
  assert.deepEqual([[0], [1], [2], [3], [1], [0], [1], [2]].map(([l]) => next('1', l)), ['I.', '1.1.', '1.1.1', 'I.1.a.1', '1.2.', 'II.', '2.1.', '2.1.1']);
});

test('picture bullets: the bullet\'s picture at its size, as an SVG data URL', () => {
  const s = new DocxSession(load('numbering-legal.docx'));
  const url = s.pictureBullets.get('0');
  assert.match(url, /^data:image\/svg\+xml;base64,/);
  const svg = atob(url.split(',')[1]);
  assert.match(svg, /width="9pt" height="9pt"/);
  assert.match(svg, /href="data:image\/png;base64,/);
});

test('Strict documents read as their Transitional form and save as Transitional', () => {
  const strict = load('strict.docx');
  assert.ok(isStrict(unzipSync(strict)));
  const s = new DocxSession(strict);
  const plain = new DocxSession(load('rich.docx'));
  assert.ok(text(s.doc).endsWith(text(plain.doc)));
  assert.match(text(s.doc), /^Half the page widein Strict percentages/);
  const out = unzipSync(s.save());
  for (const [p, b] of Object.entries(out)) if (/\.(xml|rels)$/.test(p)) assert.ok(!strFromU8(b).includes('purl.oclc.org'), p);
  // Percentages: 50% is 2500 fiftieths in Transitional WordprocessingML.
  assert.match(strFromU8(out['word/document.xml']), /<w:tblW w:w="2500" w:type="pct"\/>/);
  // A package with nothing Strict in it is left as it is.
  const files = unzipSync(load('rich.docx'));
  const same = fromStrict(files);
  for (const k of Object.keys(files)) assert.equal(same[k], files[k]);
});

test('line numbering settings', () => {
  assert.equal(lineNumbering('<w:sectPr><w:pgSz/></w:sectPr>'), null);
  assert.deepEqual(lineNumbering('<w:sectPr><w:lnNumType w:countBy="1" w:restart="newPage"/></w:sectPr>'), { countBy: 1, start: 0, distance: 24, restart: 'newPage' });
  assert.deepEqual(lineNumbering('<w:lnNumType w:countBy="2" w:distance="567" w:start="4" w:restart="continuous"/>'), { countBy: 2, start: 4, distance: 37.8, restart: 'continuous' });
});

test('a picture watermark is a washed-out picture placed on the page; text watermarks need a browser to draw', () => {
  const s = new DocxSession(load('watermark-picture.docx'), { imageUrl: () => 'blob:mark' });
  const hdr = s.story(s.pkg.rels.find((r) => /header/.test(r.type)).id);
  let img = null;
  hdr.descendants((n) => { if (n.type.name === 'image') img = n; });
  assert.ok(img);
  assert.match(img.attrs.place, /--odx-wm:1/);
  assert.match(img.attrs.place, /--odx-wm-h:center;--odx-wm-hrel:margin/);
  assert.match(img.attrs.place, /opacity:0\.3/);
  assert.equal(img.attrs.width, 400);
});

test('compatibility rules are read by name, a rule turned off left out', () => {
  const files = unzipSync(load('rich.docx'));
  const rels = strFromU8(files['word/_rels/document.xml.rels']);
  const target = /Type="[^"]*\/settings" Target="([^"]+)"/.exec(rels)?.[1];
  assert.ok(target, 'the fixture has a settings part');
  const path = 'word/' + target.replace(/^\/?word\//, '');
  files[path] = strToU8(strFromU8(files[path]).replace(/<w:compat>[\s\S]*?<\/w:compat>|(?=<\/w:settings>)/, '<w:compat><w:splitPgBreakAndParaMark/><w:useFELayout w:val="0"/><w:doNotUseIndentAsNumberingTabStop w:val="1"/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="11"/></w:compat>'));
  const s = new DocxSession(zipSync(files));
  assert.equal(s.compatMode, 11);
  assert.deepEqual([...s.compat].sort(), ['doNotUseIndentAsNumberingTabStop', 'splitPgBreakAndParaMark']);
});

test('a picture bullet\'s picture is in the stylesheet once, by class', () => {
  const s = new DocxSession(load('numbering-legal.docx'));
  const css = s.picBulletCss('.x');
  assert.equal((css.match(/--odx-picbullet:url\("data:image\/svg\+xml;base64,/g) || []).length, 1);
  assert.match(css, /^\.x p\.odx-pb-0\{/);
});

test('TIFF pictures (uncompressed, LZW, CCITT fax) are shown as PNGs; the document keeps the TIFFs', () => {
  const files = unzipSync(load('tiff.docx'));
  const tiffs = Object.keys(files).filter((k) => k.endsWith('.tif'));
  assert.equal(tiffs.length, 3);
  for (const k of tiffs) {
    assert.ok(isTiff(files[k]), k);
    const png = tiffToPng(files[k]);
    assert.ok(png && png[0] === 0x89 && png[1] === 0x50, `${k} decoded`);
  }
  const urls = [];
  const s = new DocxSession(load('tiff.docx'), { imageUrl: (bytes, mime) => { urls.push(mime); return 'blob:x'; } });
  assert.deepEqual(urls, ['image/png', 'image/png', 'image/png']);
  const out = unzipSync(s.save());
  for (const k of tiffs) assert.deepEqual(out[k], files[k], `${k} kept as it was`);
  assert.equal(tiffToPng(new Uint8Array([0x49, 0x49, 0x2a, 0, 8, 0, 0, 0, 0, 0])), null); // (a broken one: nothing, no throw)
});
