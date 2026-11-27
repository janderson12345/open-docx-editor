import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { DocxSession } from '../src/ooxml/docx.js';
import { readDocument } from '../src/ooxml/reader.js';
import { parseXml } from '../src/ooxml/xml.js';

// A form's header and footer lines in floating VML text boxes placed against the page (mkvmlhf.mjs).
const bytes = fs.readFileSync(new URL('./fixtures/vml-hf-boxes.docx', import.meta.url));

const pictures = (doc) => { const out = []; doc.descendants((n) => { if (n.type.name === 'image') out.push(n); if (n.isText && n.text.trim()) out.push(n.text); }); return out; };

test('floating VML text boxes in a header and footer are drawn as placed boxes, not as text in their line', () => {
  const s = new DocxSession(bytes);
  const foot = pictures(s.story('rIdFt1'));
  assert.equal(foot.length, 3, 'three boxes, no text left in the paragraph');
  for (const n of foot) {
    assert.equal(typeof n, 'object');
    assert.match(n.attrs.place, /^position:absolute;/);
    assert.match(n.attrs.place, /z-index:-1/, 'behind the text (z-index below 0)');
    assert.match(n.attrs.place, /--odx-arel:page/);
    assert.match(n.attrs.xml, /<v:shape\b[\s\S]*<v:textbox\b/, 'written back as it was');
    assert.doesNotMatch(n.attrs.xml, /<w:drawing\b/);
  }
  // Its place to the hundredth of a px (736.1pt), its size, and the page number marked for each page's own.
  assert.match(foot[0].attrs.place, /--odx-ay:981\.47px/);
  assert.equal(Math.round(foot[0].attrs.width), 468);
  assert.match(foot[0].attrs.place, /left:calc\(46\.67px - var\(--odx-hf-left, 0px\)\)|left:calc\(47px - var\(--odx-hf-left, 0px\)\)/);
  assert.match(decodeURIComponent(foot[2].attrs.src), /data-odx-fld="PAGE"/);
  assert.match(decodeURIComponent(foot[1].attrs.src), /TEST 4\/20\.B[\s\S]*Page 1 of 3/);
  const head = pictures(s.story('rIdHd1'));
  assert.equal(head.length, 2);
  assert.match(decodeURIComponent(head[1].attrs.src), /Supplement B[\s\S]*Form Instructions/);
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word"';
const body = (shape) => `<w:document ${W}><w:body><w:p><w:r><w:pict>${shape}</w:pict></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`;
const box = (style, extra = '', attrs = '') => `<v:shape style="position:absolute;margin-left:10pt;margin-top:20pt;width:100pt;height:30pt;${style}"${attrs}><v:textbox><w:txbxContent><w:p><w:r><w:t>Inside</w:t></w:r></w:p></w:txbxContent></v:textbox>${extra}</v:shape>`;
const read = (xml) => {
  const shapes = [];
  const doc = readDocument(parseXml(xml), { resolveShape: (d, w, h) => { shapes.push(d); return `data:image/svg+xml,${w}x${h}`; } });
  const p = doc.content[0];
  return { content: p.content || [], shape: shapes[0] };
};

test('a floating VML text box: its place, wrapping, fill and line as the drawing it stands for', () => {
  // Placed against its paragraph by default, wrapped square on the side it says, filled and lined unless it says not.
  let r = read(body(box('mso-wrap-distance-left:4pt', '<w10:wrap type="square" side="left"/>', ' fillcolor="#ff0000" strokeweight="2pt"')));
  assert.equal(r.content.length, 1);
  assert.equal(r.content[0].type, 'image');
  assert.match(r.content[0].attrs.place, /^float:right/, 'text on its left only: floated right');
  const a = r.shape.getElementsByTagName('wp:anchor')[0];
  assert.equal(a.getAttribute('distL'), String(4 * 12700));
  assert.equal(r.shape.getElementsByTagName('wp:positionV')[0].getAttribute('relativeFrom'), 'paragraph');
  assert.equal(r.shape.getElementsByTagName('wp:positionH')[0].getAttribute('relativeFrom'), 'column');
  assert.equal(r.shape.getElementsByTagName('a:srgbClr')[0].getAttribute('val'), 'ff0000');
  assert.equal(r.shape.getElementsByTagName('a:ln')[0].getAttribute('w'), String(2 * 12700));
  // VML's default insets (0.1in, 0.05in) and no wrapping when w10:wrap gives no type; centred across the page.
  r = read(body(box('mso-position-horizontal:center;mso-position-horizontal-relative:page;z-index:5', '', ' filled="f" stroked="f"')));
  assert.match(r.content[0].attrs.place, /^position:absolute;left:calc\(0px \+ \(100% - 133px\) \/ 2\);/);
  assert.match(r.content[0].attrs.place, /z-index:1/);
  const bp = r.shape.getElementsByTagName('wps:bodyPr')[0];
  assert.deepEqual(['lIns', 'tIns', 'rIns', 'bIns'].map((k) => bp.getAttribute(k)), ['91440', '45720', '91440', '45720']);
  assert.ok(r.shape.getElementsByTagName('a:noFill').length >= 2, 'no fill, no line');
  // Turned, or of vertical text: kept as before (its text in the line).
  for (const style of ['rotation:90', 'flip:x']) {
    r = read(body(box(style)));
    assert.notEqual(r.content[0]?.type, 'image', style);
  }
});

// Inline VML text boxes and an inline VML group of two text boxes (mkvmlgroup.mjs), as Word sets them (its prints of the
// same document, ..\vmlgroup-word): each box at its place in the group, the title's box over the body's; the group's line
// its height and its outlines' outer halves; at its own size.
test('an inline VML group of text boxes is drawn as the DrawingML group Word saves for it, each box at its place', () => {
  const s = new DocxSession(fs.readFileSync(new URL('./fixtures/vml-inline-boxes.docx', import.meta.url)));
  const imgs = [];
  s.doc.descendants((n) => { if (n.type.name === 'image') imgs.push(n); });
  assert.equal(imgs.length, 3, 'two groups and one box, no text left in their lines');
  const [g] = imgs;
  assert.match(g.attrs.xml, /<v:group\b/, 'written back as it was');
  assert.equal(Math.round(g.attrs.width), 600); // 450pt
  assert.match(g.attrs.place, /max-width:none/, 'its own size, whatever the line');
  assert.match(g.attrs.place, /--odx-fxt:-1\.046875px/, '1.44pt outlines: 0.96px of room above and below, under the 2px pad');
  const svg = decodeURIComponent(g.attrs.src);
  // (Each box's text in a group moved to the box's place: translate(x y).)
  const y = (t) => { const m = new RegExp(`translate\\([\\d.]+ ([\\d.]+)\\)"><text[^>]*?y="([\\d.]+)"[^>]*>(?:(?!</text>).)*${t}`).exec(svg); return m ? Number(m[1]) + Number(m[2]) : NaN; };
  assert.ok(y('Form Instruction') < y('The Placeholders'), 'the title over the body, though it comes after it in the file');
});

// One line of text in floating VML text boxes in Arial, Times New Roman and Calibri (mktboxbase.mjs): its baseline the
// font's ascent in Word's line under the box's top and the space before (Word's prints, ..\tboxbase-word: 9.76, 11.0,
// 15.9 and 15.0px for Arial 7 and 8pt, Times New Roman 12pt and Calibri 11pt).
test('a text box line\'s baseline is its font\'s ascent in Word\'s line below the box\'s top', () => {
  const s = new DocxSession(fs.readFileSync(new URL('./fixtures/tbox-baselines.docx', import.meta.url)));
  const ys = [];
  s.doc.descendants((n) => { if (n.type.name === 'image') ys.push(Number(/<text[^>]*?y="([\d.]+)"/.exec(decodeURIComponent(n.attrs.src))?.[1])); });
  assert.equal(ys.length, 9);
  for (const [i, want] of [[0, 9.76], [1, 11.0], [6, 15.9], [8, 15.0]]) assert.ok(Math.abs(ys[i] - want) < 0.35, `box ${i}: ${ys[i]} (Word: ${want})`);
});

test('a floating VML group of text boxes is drawn as a placed DrawingML group; side-by-side bands each at their place', () => {
  const shape = (l, t, text) => `<v:shape type="#_x0000_t202" style="position:absolute;left:${l};top:${t};width:400;height:200" filled="f" strokeweight=".96pt"><v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>`;
  const group = (left) => `<v:group style="position:absolute;margin-left:${left}pt;margin-top:12pt;width:100pt;height:${left > 100 ? 60 : 50}pt;mso-position-horizontal-relative:page" coordorigin="100,100" coordsize="400,400">${shape(100, 100, 'Title')}${shape(100, 300, 'Body')}<w10:wrap type="topAndBottom" anchorx="page"/></v:group>`;
  const xml = body(group(100)).replace('</w:pict></w:r>', `</w:pict></w:r><w:r><w:pict>${group(300)}</w:pict></w:r>`);
  const r = read(xml);
  assert.equal(r.content.length, 2);
  assert.ok(r.shape.getElementsByTagName('wpg:wgp').length === 1, 'a wpg:wgp');
  assert.equal(r.shape.getElementsByTagName('a:chOff')[0].getAttribute('x'), '100');
  // The taller keeps the band; the other drawn at its place, out of the flow, first in the paragraph.
  assert.match(r.content[0].attrs.place, /^position:absolute;left:calc\(var\(--odx-page-left, 96px\) \+ /);
  assert.match(r.content[1].attrs.place, /^display:block;/);
});
