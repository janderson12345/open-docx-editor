import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { DocxSession } from '../src/ooxml/docx.js';

const bytes = fs.readFileSync(new URL('./fixtures/textbox.docx', import.meta.url));

/** The fixture's floating text box made as wide as the column (624px): text can't go beside it, so it's a band. */
function wideBox() {
  const files = unzipSync(bytes);
  const xml = strFromU8(files['word/document.xml']).replace(/(<wp:anchor\b[\s\S]*?<wp:extent cx=")\d+(" cy="\d+")/, `$1${624 * 9525}$2`).replace(/(<wp:positionH relativeFrom="column"><wp:posOffset>)\d+/, '$10');
  files['word/document.xml'] = strToU8(xml);
  const s = new DocxSession(zipSync(files));
  let node = null;
  s.doc.descendants((n) => { if (!node && n.type.name === 'image' && /<wp:anchor/.test(n.attrs.xml)) node = n; });
  return { s, node };
}

test('a box as wide as the column stays a band (text above and below it) when it is resized or moved', () => {
  const { s, node } = wideBox();
  assert.match(node.attrs.place, /--odx-band:1/, 'read as a band');
  const bigger = s.resizeShape(node.attrs.xml, 624, node.attrs.height + 20);
  assert.match(bigger.place, /--odx-band:1/, 'after a resize');
  const moved = s.moveShape(node.attrs.xml, 0, 16);
  assert.match(moved.place, /--odx-band:1/, 'after a move');
});

test('a resize changes the frame, the shape and the VML copy alike, and nothing else of the shape', () => {
  const { s, node } = wideBox();
  const r = s.resizeShape(node.attrs.xml, 300, 120);
  assert.equal(r.width, 300);
  assert.equal(r.height, 120);
  assert.match(r.xml, new RegExp(`<wp:extent cx="${300 * 9525}" cy="${120 * 9525}"`));
  assert.match(r.xml, new RegExp(`<a:ext cx="${300 * 9525}" cy="${120 * 9525}"`));
  assert.match(r.xml, /width:225pt;height:90pt/);
  const strip = (x) => x.replace(/cx="\d+"|cy="\d+"|width:[\d.]+pt;height:[\d.]+pt/g, '');
  assert.equal(strip(r.xml), strip(node.attrs.xml));
});

test('a move shifts the offsets and the VML margins; a box placed by alignment isn\'t moved', () => {
  const { s, node } = wideBox();
  const m = s.moveShape(node.attrs.xml, 10, -4);
  const off = (x, a) => Number(new RegExp(`<wp:position${a}\\b[^>]*>\\s*<wp:posOffset>(-?\\d+)`).exec(x)[1]);
  assert.equal(off(m.xml, 'H') - off(node.attrs.xml, 'H'), 10 * 9525);
  assert.equal(off(m.xml, 'V') - off(node.attrs.xml, 'V'), -4 * 9525);
  assert.match(m.xml, /margin-left:295.5pt;margin-top:4.5pt/);
  const aligned = node.attrs.xml.replace(/<wp:posOffset>-?\d+<\/wp:posOffset>(<\/wp:positionH>)/, '<wp:align>center</wp:align>$1');
  assert.equal(s.moveShape(aligned, 10, 0), null);
});

test('a text box shows the paragraphs inside block content controls; one placed against the page is marked for its page', () => {
  const files = unzipSync(bytes);
  const xml = strFromU8(files['word/document.xml'])
    .replace(/<w:txbxContent><w:p>([\s\S]*?)<\/w:p><\/w:txbxContent>/g, '<w:txbxContent><w:sdt><w:sdtPr><w:alias w:val="Title"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Held in a control</w:t></w:r></w:p></w:sdtContent></w:sdt></w:txbxContent>')
    .replace(/<wp:positionV relativeFrom="\w+">[\s\S]*?<\/wp:positionV>/, '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wp14"><wp:positionV relativeFrom="page"><wp14:pctPosVOffset xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing">15000</wp14:pctPosVOffset></wp:positionV></mc:Choice><mc:Fallback><wp:positionV relativeFrom="page"><wp:posOffset>1524000</wp:posOffset></wp:positionV></mc:Fallback></mc:AlternateContent>')
    .replace(/<wp:wrap(Square|Tight|Through|TopAndBottom)\b[^>]*?(\/>|>[\s\S]*?<\/wp:wrap\1>)/, '<wp:wrapNone/>');
  files['word/document.xml'] = strToU8(xml);
  const s = new DocxSession(zipSync(files), { imageUrl: () => 'blob:x' });
  let node = null;
  s.doc.descendants((n) => { if (!node && n.type.name === 'image' && /<wp:anchor/.test(n.attrs.xml)) node = n; });
  assert.match(decodeURIComponent(node.attrs.src), /Held/);
  assert.match(node.attrs.place, /--odx-arel:page;--odx-ay:160px/);
});

test('a text box draws a run in all caps (w:caps) in capitals, its own or from its style', () => {
  const files = unzipSync(bytes);
  const xml = strFromU8(files['word/document.xml'])
    .replace(/<w:txbxContent><w:p>([\s\S]*?)<\/w:p><\/w:txbxContent>/g, '<w:txbxContent><w:p><w:r><w:rPr><w:caps/></w:rPr><w:t>Banner title</w:t></w:r><w:r><w:t> plain words</w:t></w:r></w:p></w:txbxContent>');
  files['word/document.xml'] = strToU8(xml);
  const s = new DocxSession(zipSync(files), { imageUrl: () => 'blob:x' });
  let node = null;
  s.doc.descendants((n) => { if (!node && n.type.name === 'image' && /<wp:anchor/.test(n.attrs.xml)) node = n; });
  const svg = decodeURIComponent(node.attrs.src);
  assert.match(svg, /BANNER TITLE plain/);
});

test('a text box takes a character style\'s theme font over Times New Roman defaults, and draws its small caps', () => {
  const files = unzipSync(bytes);
  const xml = strFromU8(files['word/document.xml'])
    .replace(/<w:txbxContent><w:p>([\s\S]*?)<\/w:p><\/w:txbxContent>/g, '<w:txbxContent><w:p><w:r><w:rPr><w:rStyle w:val="coversub"/></w:rPr><w:t>Monthly Report</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Headline</w:t></w:r></w:p></w:txbxContent>');
  files['word/document.xml'] = strToU8(xml);
  files['word/styles.xml'] = strToU8(strFromU8(files['word/styles.xml'])
    .replace(/<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"\/>/, '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>')
    .replace('</w:styles>', '<w:style w:type="character" w:customStyle="1" w:styleId="coversub"><w:name w:val="cover:sub"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:b/><w:smallCaps/><w:sz w:val="40"/></w:rPr></w:style></w:styles>'));
  const s = new DocxSession(zipSync(files), { imageUrl: () => 'blob:x' });
  let node = null;
  s.doc.descendants((n) => { if (!node && n.type.name === 'image' && /<wp:anchor/.test(n.attrs.xml)) node = n; });
  const svg = decodeURIComponent(node.attrs.src);
  const fam = (word) => new RegExp(`<tspan font-family="([^"]*)"[^>]*>[^<]*(?:<tspan[^>]*>[^<]*</tspan>[^<]*)*${word}`).exec(svg)?.[1] || '';
  assert.match(fam('M'), /^&quot;Calibri&quot;/, 'the character style\'s minorHAnsi, not the defaults\' Times');
  assert.match(fam('Headline'), new RegExp(`^&quot;${s.theme.major}&quot;`), 'the paragraph style\'s majorHAnsi');
  // 20pt: lowercase letters as capitals at 16pt (21.33px), capitals and spaces at 26.67px.
  assert.match(svg, /M<tspan font-size="21.33">ONTHLY<\/tspan> R<tspan font-size="21.33">EPORT<\/tspan>/);
});

test('the word count takes in text boxes once (not their fallback copy)', async () => {
  const { wordCount } = await import('../src/editor/commands.js');
  const s = new DocxSession(bytes);
  const withBoxes = wordCount(s.doc).words;
  const { Transform } = await import('prosemirror-transform');
  const tr = new Transform(s.doc);
  const at = [];
  s.doc.descendants((n, pos) => { if (n.type.name === 'image') at.push(pos); });
  for (const pos of at.reverse()) tr.delete(pos, pos + 1);
  const body = wordCount(tr.doc).words;
  // The fixture's boxes: "An inline text box whose text wraps when it is resized." (11) and "A floating text box, placed by offsets." (7)
  assert.equal(withBoxes - body, 18);
});
