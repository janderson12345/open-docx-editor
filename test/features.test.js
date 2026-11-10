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

test('a cropped picture (a:srcRect) is shown cropped: only its part, stretched to the frame', async () => {
  const { blankDocxBytes } = await import('../src/ooxml/template.js');
  const png = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==', 'base64'));
  const s0 = new DocxSession(blankDocxBytes());
  const { xml } = s0.addImage(png, 'image/png', 100, 50);
  const cropped = xml.replace('<a:stretch>', '<a:srcRect l="25000" t="10000" r="25000" b="40000"/><a:stretch>');
  const main = s0.pkg.mainPath;
  s0.pkg.setText(main, s0.pkg.text(main).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body><w:p>${cropped}</w:p><w:p>${xml}</w:p>$1`));
  const urls = [];
  const s = new DocxSession(s0.pkg.toBytes(), { imageUrl: (b, mime) => { urls.push(mime); return 'blob:whole'; } });
  const srcs = [];
  s.doc.descendants((n) => { if (n.type.name === 'image') srcs.push(n.attrs.src); });
  assert.equal(srcs[1], 'blob:whole');
  assert.match(srcs[0], /^data:image\/svg\+xml;base64,/);
  const svg = Buffer.from(srcs[0].split(',')[1], 'base64').toString();
  // Half the width (a quarter off each side) and half the height (10% off the top, 40% off the bottom).
  assert.match(svg, /<image href="data:image\/png;base64,[^"]+" x="-50" y="-20" width="200" height="200" preserveAspectRatio="none"\/>/);
  // The document keeps the crop and the picture as they were.
  assert.match(strFromU8(unzipSync(s.save())[main]), /<a:srcRect l="25000" t="10000" r="25000" b="40000"\/>/);
});

test('floating pictures in front of (or behind) the text are stacked by their drawing order, not their place in the text', async () => {
  const { readDocument } = await import('../src/ooxml/reader.js');
  const { parseXml } = await import('../src/ooxml/xml.js');
  const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const pic = (rh, behind) => `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${rh}" behindDoc="${behind}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>`
    + '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
    + '<wp:extent cx="952500" cy="952500"/><wp:wrapNone/><wp:docPr id="1" name="P"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId9"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
  const read = (body) => {
    const json = readDocument(parseXml(`<w:document ${ns}><w:body>${body}</w:body></w:document>`), { resolveImage: () => 'blob:p' });
    return json.content.flatMap((p) => p.content || []).filter((n) => n.type === 'image').map((n) => /z-index:(-?\d+)/.exec(n.attrs.place)[1]);
  };
  // In the text's order already: left as they are.
  assert.deepEqual(read(`<w:p>${pic(100, 0)}${pic(200, 0)}</w:p>`), ['1', '1']);
  // The first one drawn over the second: it goes on top.
  assert.deepEqual(read(`<w:p>${pic(300, 0)}${pic(200, 0)}${pic(100, 1)}${pic(50, 1)}</w:p>`), ['2', '1', '-1', '-2']);
});

test('an autofit table with a placeholder grid and a merged first row takes its columns from the cells of the next row', async () => {
  const { readDocument } = await import('../src/ooxml/reader.js');
  const { parseXml } = await import('../src/ooxml/xml.js');
  const tc = (w, t, span) => `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="${w}"/>${span ? `<w:gridSpan w:val="${span}"/>` : ''}</w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
  const tbl = '<w:tbl><w:tblPr><w:tblW w:type="auto" w:w="100"/></w:tblPr><w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>'
    + `<w:tr>${tc(4500, 'Title', 3)}</w:tr><w:tr>${tc(1500, 'a')}${tc(1500, 'b')}${tc(1500, 'c')}</w:tr></w:tbl><w:p/>`;
  const json = readDocument(parseXml(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${tbl}</w:body></w:document>`));
  const cells = json.content[0].content[1].content;
  assert.deepEqual(cells.map((c) => c.attrs.colwidth[0]), [100, 100, 100]);
  assert.deepEqual(json.content[0].content[0].content[0].attrs.colwidth, [100, 100, 100]);
});

test('Chinese and Hebrew list numbers, and leading zeros kept under legal numbering', async () => {
  const { formatNumber } = await import('../src/ooxml/styles.js');
  assert.deepEqual([1, 2, 10, 11, 20, 21, 101, 110].map((n) => formatNumber(n, 'chineseCounting')), ['一', '二', '十', '十一', '二十', '二十一', '一百零一', '一百一十']);
  assert.deepEqual([1, 10, 15, 16, 20, 115].map((n) => formatNumber(n, 'hebrew1')), ['א', 'י', 'טו', 'טז', 'כ', 'קטו']);
  assert.deepEqual([1, 22, 23].map((n) => formatNumber(n, 'hebrew2')), ['א', 'ת', 'אא']);
  const s = new DocxSession(load('numbering-legal.docx'));
  const lv = s.numbering.abstract.get([...s.numbering.nums.values()][0].abstractId);
  lv[0].fmt = 'decimalZero'; lv[1].legal = true; lv[1].fmt = 'decimalZero'; lv[1].text = '%1.%2';
  const next = listCounter(s.numbering);
  const id = [...s.numbering.nums.keys()][0];
  next(id, 0);
  assert.equal(next(id, 1), '01.01');
});

test('a list number formatted as hidden text is not shown, nor the room after it', async () => {
  const { listLabel } = await import('../src/editor/pararules.js');
  const { schema } = await import('../src/editor/schema.js');
  const session = { numbering: { nums: new Map([['1', { abstractId: 'a', overrides: {} }]]), abstract: new Map([['a', [{ start: 1, fmt: 'decimal', text: '%1.', pPr: { indLeft: 720, indFirst: -360 }, jc: 'left', suff: 'tab', hidden: true }]]]) }, styleProp: () => null, defaultTabStop: 720 };
  const run = (extra) => {
    const node = schema.nodes.paragraph.create({ numId: '1', ilvl: 0, extra });
    const attrs = {}, style = [];
    listLabel(node, session, { numId: '1', ilvl: 0 }, '1.', attrs, style);
    return [attrs['data-label'], style.find((s) => s.startsWith('--odx-hang'))];
  };
  assert.deepEqual(run([]), ['', '--odx-hang:0pt']);
  session.numbering.abstract.get('a')[0].hidden = false;
  assert.deepEqual(run([]), ['1.', '--odx-hang:18pt']);
  assert.deepEqual(run([['rPr', '<w:rPr><w:vanish/></w:rPr>']]), ['', '--odx-hang:0pt']);
});

test('a pasted picture given as a data: URL is stored as the document\'s own; a pasted copy gets drawing ids of its own', async () => {
  const { pastePictures } = await import('../src/editor/pastepics.js');
  const { schema } = await import('../src/editor/schema.js');
  const { Slice, Fragment } = await import('prosemirror-model');
  const { blankDocxBytes } = await import('../src/ooxml/template.js');
  const s = new DocxSession(blankDocxBytes(), { imageUrl: () => 'blob:stored' });
  const png = 'iVBORw0KGgoAAAANSUhEUgAAAAIAAAADCAYAAAC56t6BAAAAEklEQVR42mP8z8DwnwEJMDIgAQBAGgP/8m3SOgAAAABJRU5ErkJggg==';
  const own = s.addImage(Uint8Array.from(Buffer.from(png, 'base64')), 'image/png', 20, 30);
  const p = schema.nodes.paragraph.create(null, [
    schema.nodes.image.create({ src: `data:image/png;base64,${png}`, xml: '' }),
    schema.nodes.image.create({ src: 'blob:own', xml: own.xml, width: 20, height: 30 }),
  ]);
  const out = pastePictures(s).props.transformPasted(new Slice(Fragment.from(p), 0, 0)).content.firstChild;
  const [pasted, copy] = [out.child(0), out.child(1)];
  assert.match(pasted.attrs.xml, /<a:blip r:embed="rId\d+"/);
  assert.deepEqual([pasted.attrs.src, pasted.attrs.width, pasted.attrs.height], ['blob:stored', 2, 3]); // (its own size, from the PNG)
  const id = (x) => /<wp:docPr id="(\d+)"/.exec(x)[1];
  assert.notEqual(id(copy.attrs.xml), id(own.xml));
  assert.ok(Object.keys(s.pkg.files).filter((k) => /media\//.test(k)).length >= 2);
});

test('an SVG picture is shown from its SVG, not its PNG fallback', async () => {
  const { readDocument } = await import('../src/ooxml/reader.js');
  const { parseXml } = await import('../src/ooxml/xml.js');
  const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const blip = '<a:blip r:embed="rIdPng"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdSvg"/></a:ext></a:extLst></a:blip>';
  const body = `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="1" name="Icon"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill>${blip}</pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
  const json = readDocument(parseXml(`<w:document ${ns}><w:body>${body}</w:body></w:document>`), { resolveImage: (id) => `blob:${id}` });
  assert.equal(json.content[0].content[0].attrs.src, 'blob:rIdSvg');
});

test('built-in headings used but not defined look like Word\'s and count as headings; a style hidden until used joins the gallery once used', async () => {
  const { stylesToCss, styleMenu } = await import('../src/ooxml/styles.js');
  const { outlineLevel } = await import('../src/editor/commands.js');
  const { schema } = await import('../src/editor/schema.js');
  const info = { styles: new Map(), defaults: { rPr: {}, pPr: {} }, defaultParagraph: null };
  const css = stylesToCss('.d', info, { minor: 'Calibri', major: 'Calibri Light' });
  assert.match(css, /\.d p\[data-style="Heading1"\]\{font-family:var\(--odx-font-major\)[^}]*font-size:16pt;color:#2F5496/);
  const session = { styles: info, styleProp: () => null };
  assert.equal(outlineLevel(session, schema.nodes.paragraph.create({ styleId: 'Heading2' })), 1);
  const st = (id, name, extra = {}) => [id, { id, name, type: 'paragraph', uiPriority: 9, hidden: true, unhide: true, qFormat: true, ...extra }];
  const gallery = { styles: new Map([st('Heading1', 'heading 1', { hidden: false, unhide: false }), st('Heading2', 'heading 2'), st('Heading3', 'heading 3')]) };
  assert.deepEqual(styleMenu(gallery, new Set()).map((s) => s.id), ['Heading1']);
  assert.deepEqual(styleMenu(gallery, new Set(['Heading1'])).map((s) => s.id), ['Heading1', 'Heading2']);
});
