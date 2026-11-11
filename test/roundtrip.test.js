import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import mammoth from 'mammoth';
import { DocxSession } from '../src/ooxml/docx.js';
import { parseXml } from '../src/ooxml/xml.js';
import { Transform } from 'prosemirror-transform';
import { blankDocxBytes } from '../src/ooxml/template.js';
import { isStrict, fromStrict } from '../src/ooxml/strict.js';

const dir = new URL('./fixtures/', import.meta.url);
const fixtures = fs.readdirSync(dir).filter((f) => f.endsWith('.docx'));
const load = (f) => new Uint8Array(fs.readFileSync(new URL(f, dir)));
const rawText = async (bytes) => (await mammoth.extractRawText({ buffer: Buffer.from(bytes) })).value;

for (const f of [...fixtures, '(blank)']) {
  test(`round-trip: ${f}`, async () => {
    const given = f === '(blank)' ? blankDocxBytes() : load(f);
    // (A Strict document is saved in the Transitional form: compared with that.)
    const bytes = isStrict(unzipSync(given)) ? zipSync(fromStrict(unzipSync(given))) : given;
    const s1 = new DocxSession(given);
    const out = s1.save();

    // 1. Every part except the main document is untouched, byte for byte.
    const a = unzipSync(bytes), b = unzipSync(out);
    assert.deepEqual(Object.keys(b).sort(), Object.keys(a).sort());
    for (const k of Object.keys(a)) if (k !== s1.pkg.mainPath) assert.deepEqual(b[k], a[k], `part changed: ${k}`);

    // 2. The new document.xml is well-formed.
    parseXml(strFromU8(b[s1.pkg.mainPath]));

    // 3. Reading it back gives the same editor document (the writer is lossless for the model).
    const s2 = new DocxSession(out);
    assert.deepEqual(s2.doc.toJSON(), s1.doc.toJSON());

    // 4. An independent reader (mammoth) sees the same text.
    assert.equal(await rawText(out), await rawText(bytes));
  });
}

test('edits survive save: text, bold, list, link, table cell', async () => {
  const s = new DocxSession(load('rich.docx'));
  const { schema } = s.doc.type;
  let doc = s.doc;
  const para = (text, attrs, marks) => schema.nodes.paragraph.create(attrs, schema.text(text, marks));
  const bullet = s.listId('bullet');
  const link = schema.marks.link.create({ href: 'https://github.com/' });
  const added = [
    para('Brand new bold paragraph', null, [schema.marks.bold.create()]),
    para('New bullet item', { numId: bullet, ilvl: 0 }),
    schema.nodes.paragraph.create(null, [schema.text('See '), schema.text('GitHub', [link])]),
  ];
  doc = doc.copy(doc.content.append(schema.nodes.doc.create(null, added).content));
  // Replace the first table cell's text.
  let cellPara = null;
  doc.descendants((n, pos) => {
    if (cellPara == null && n.type.name === 'table_cell') cellPara = { from: pos + 2, to: pos + 2 + n.firstChild.content.size };
    return cellPara == null;
  });
  doc = new Transform(doc).replaceWith(cellPara.from, cellPara.to, schema.text('EDITED')).doc;
  const out = s.save(doc);
  const { value: html } = await mammoth.convertToHtml({ buffer: Buffer.from(out) });
  assert.match(html, /<strong>Brand new bold paragraph<\/strong>/);
  assert.match(html, /<ul><li>New bullet item<\/li><\/ul>/);
  assert.match(html, /<a href="https:\/\/github.com\/">GitHub<\/a>/);
  assert.match(html, /<td><p>EDITED<\/p><\/td>/);
  // and it reads back the same
  assert.deepEqual(new DocxSession(out).doc.toJSON().content.length, doc.toJSON().content.length);
});

test('short table rows open and save', () => {
  const files = unzipSync(blankDocxBytes());
  const s0 = new DocxSession(blankDocxBytes());
  const tc = (t) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
  const xml = strFromU8(files[s0.pkg.mainPath]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${tc('a')}${tc('b')}</w:tr><w:tr>${tc('only')}</w:tr></w:tbl><w:p/>$1`);
  files[s0.pkg.mainPath] = strToU8(xml);
  const s1 = new DocxSession(zipSync(files));
  const out = strFromU8(unzipSync(s1.save())[s1.pkg.mainPath]);
  assert.equal((out.match(/<w:tc>/g) || []).length, 3);
  assert.match(out, />only<\/w:t>/);
});

test('extension run properties (w14) stay after the w: ones when saved', () => {
  const files = unzipSync(blankDocxBytes());
  const main = new DocxSession(blankDocxBytes()).pkg.mainPath;
  const run = '<w:r><w:rPr><w:color w:val="4472C4"/><w:sz w:val="28"/><w14:shadow xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"/></w:rPr><w:t>Glow</w:t></w:r>';
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body><w:p>${run}</w:p>$1`));
  const s = new DocxSession(zipSync(files));
  const out = strFromU8(unzipSync(s.save(s.doc))[s.pkg.mainPath]);
  assert.match(out, /<w:color w:val="4472C4"\/><w:sz w:val="28"\/><w14:shadow[^>]*\/><\/w:rPr><w:t[^>]*>Glow/);
});

test('"Different Odd & Even Pages" is written into settings.xml in its place and can be turned off', () => {
  const s = new DocxSession(load('word-template.docx'));
  assert.ok(s.settingsPath);
  assert.equal(s.evenAndOddHeaders, false);
  s.evenAndOddHeaders = true;
  const on = strFromU8(unzipSync(s.save(s.doc))[s.settingsPath]);
  assert.match(on, /<w:evenAndOddHeaders\/>/);
  if (/<w:compat\b/.test(on)) assert.ok(on.indexOf('<w:evenAndOddHeaders/>') < on.indexOf('<w:compat'));
  s.evenAndOddHeaders = false;
  assert.doesNotMatch(strFromU8(unzipSync(s.save(s.doc))[s.settingsPath]), /evenAndOddHeaders/);
});

test('comments added to a document whose comments part is empty (<w:comments/>) give well-formed XML', async () => {
  const { writeComments, COMMENT_RELS } = await import('../src/ooxml/comments.js');
  const parts = new Map([['word/comments.xml', '<?xml version="1.0"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>']]);
  const pkg = {
    mainDir: 'word/',
    partByType: (t) => (t === COMMENT_RELS.comments ? 'word/comments.xml' : null),
    text: (p) => parts.get(p),
    setText: (p, s) => parts.set(p, s),
    addRel() {}, addOverride() {},
  };
  writeComments(pkg, [{ id: '0', author: 'A', initials: 'A', date: '2026-01-01T00:00:00Z', text: 'Hello', paraId: null, parentId: null, done: false, xml: null }]);
  const xml = parts.get('word/comments.xml');
  assert.match(xml, /<w:comments\b[^>]*[^/]><w:comment\b[\s\S]*Hello[\s\S]*<\/w:comments>$/);
  assert.equal(parseXml(xml).documentElement.localName, 'comments');
});

test('a resized picture is saved at its new size (frame and picture transform), others unchanged', () => {
  const s = new DocxSession(blankDocxBytes());
  const { xml, src } = s.addImage(new Uint8Array([137, 80, 78, 71]), 'image/png', 200, 100);
  let doc = s.doc;
  const img = doc.type.schema.nodes.image;
  const p = doc.type.schema.nodes.paragraph;
  doc = doc.type.create(doc.attrs, [p.create(null, [img.create({ xml, src, width: 320, height: 160 })]), p.create(null, [img.create({ xml, src, width: 200, height: 100 })])]);
  const out = strFromU8(unzipSync(s.save(doc))[s.pkg.mainPath]);
  const exts = [...out.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"/g)].map((m) => [+m[1], +m[2]]);
  const pics = [...out.matchAll(/<pic:spPr>.*?<a:ext cx="(\d+)" cy="(\d+)"/g)].map((m) => [+m[1], +m[2]]);
  assert.deepEqual(exts, [[320 * 9525, 160 * 9525], [200 * 9525, 100 * 9525]]);
  assert.deepEqual(pics, [[320 * 9525, 160 * 9525], [200 * 9525, 100 * 9525]]);
});

test('text box text is edited in the shape and its VML fallback, and saved', () => {
  const files = unzipSync(blankDocxBytes());
  const main = new DocxSession(blankDocxBytes()).pkg.mainPath;
  const ns = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml"';
  const para = (id) => `<w:p><w:bookmarkStart w:id="${id}" w:name="_GoBack"/><w:r><w:t>Old text</w:t></w:r><w:bookmarkEnd w:id="${id}"/></w:p>`;
  const box = `<w:r><mc:AlternateContent ${ns}><mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="1905000" cy="952500"/><wp:docPr id="1" name="Text Box 1"/>`
    + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>'
    + `<wps:txbx><w:txbxContent>${para(5)}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>`
    + `<mc:Fallback><w:pict><v:shape style="width:150pt;height:75pt"><v:textbox><w:txbxContent>${para(6)}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body><w:p>${box}</w:p>$1`));
  const s = new DocxSession(zipSync(files));
  let pos = null;
  s.doc.descendants((n, p) => { if (pos == null && n.type.name === 'image') pos = p; });
  assert.notEqual(pos, null);
  const node = s.doc.nodeAt(pos);
  const text = s.textBox(node.attrs.xml);
  assert.equal(text.textContent, 'Old text');
  const { schema } = s.doc.type;
  const edited = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, schema.text('New & text')));
  const { xml } = s.setTextBox(node.attrs.xml, edited);
  const doc = new Transform(s.doc).setNodeMarkup(pos, null, { ...node.attrs, xml }).doc;
  const out = strFromU8(unzipSync(s.save(doc))[main]);
  parseXml(out);
  assert.equal((out.match(/New &amp; text/g) || []).length, 2);
  assert.doesNotMatch(out, /Old text/);
  const s2 = new DocxSession(s.save(doc));
  assert.equal(s2.textBox(s2.doc.nodeAt(pos).attrs.xml).textContent, 'New & text');
});

test('a new text box saves as a Word text box (shape and VML copy) and reads back with its text', () => {
  const s = new DocxSession(blankDocxBytes());
  const { xml, src } = s.addTextBox(192, 96);
  assert.ok(src);
  const { schema } = s.doc.type;
  const text = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, schema.text('Boxed')));
  const filled = s.setTextBox(xml, text);
  const doc = s.doc.type.create(s.doc.attrs, [schema.nodes.paragraph.create(null, [schema.nodes.image.create({ xml: filled.xml, src: filled.src, width: 192, height: 96 })])]);
  const out = strFromU8(unzipSync(s.save(doc))[s.pkg.mainPath]);
  parseXml(out);
  assert.match(out, /<wps:cNvSpPr txBox="1"\/>/);
  assert.equal((out.match(/>Boxed</g) || []).length, 2);
  const s2 = new DocxSession(s.save(doc));
  let node = null;
  s2.doc.descendants((n) => { if (!node && n.type.name === 'image') node = n; });
  assert.equal(s2.textBox(node.attrs.xml).textContent, 'Boxed');
  assert.equal(node.attrs.width, 192);
});

test('each text box in a group is found where it sits and edited on its own', () => {
  const s = new DocxSession(blankDocxBytes());
  const ns = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:v="urn:schemas-microsoft-com:vml"';
  const wsp = (x, t) => `<wps:wsp><wps:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="1000" cy="500"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>`;
  const vml = (t) => `<v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>`;
  const xml = `<w:r><mc:AlternateContent ${ns}><mc:Choice Requires="wpg"><w:drawing><wp:inline><wp:extent cx="1905000" cy="476250"/><wp:docPr id="1" name="Group 1"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="476250"/><a:chOff x="0" y="0"/><a:chExt cx="2000" cy="500"/></a:xfrm></wpg:grpSpPr>${wsp(0, 'Left')}${wsp(1000, 'Right')}</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:group>${vml('Left')}${vml('Right')}</v:group></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
  const boxes = s.textBoxes(xml);
  assert.deepEqual(boxes.map((b) => b.rect), [[0, 0, 0.5, 1], [0.5, 0, 0.5, 1]]);
  assert.deepEqual(boxes.map((b) => b.copies), [[0, 2], [1, 3]]);
  const { schema } = s.doc.type;
  const out = s.setTextBox(xml, schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, schema.text('New'))), 1).xml;
  assert.deepEqual(s.textBoxes(out).map((b, k) => s.textBox(out, k).textContent), ['Left', 'New']);
  assert.equal(out.split('>New<').length - 1, 2);
  assert.equal(out.split('>Left<').length - 1, 2);
});

test('a text box sized as a percentage of the text width (wp14:pctWidth) takes that width', () => {
  const files = unzipSync(blankDocxBytes());
  const main = new DocxSession(blankDocxBytes()).pkg.mainPath;
  const ns = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"';
  const box = `<w:r><w:drawing ${ns}><wp:anchor simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>`
    + '<wp:extent cx="952500" cy="476250"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Text Box 1"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
    + '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>'
    + '<wp14:sizeRelH relativeFrom="margin"><wp14:pctWidth>40000</wp14:pctWidth></wp14:sizeRelH><wp14:sizeRelV relativeFrom="margin"><wp14:pctHeight>0</wp14:pctHeight></wp14:sizeRelV></wp:anchor></w:drawing></w:r>';
  const xml = strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body><w:p>${box}</w:p>$1`);
  files[main] = strToU8(xml);
  const s = new DocxSession(zipSync(files));
  const m = /<w:pgSz w:w="(\d+)"[\s\S]*?<w:pgMar w:top="\d+" w:right="(\d+)" w:bottom="\d+" w:left="(\d+)"/.exec(xml);
  const textWidth = (Number(m[1]) - Number(m[2]) - Number(m[3])) / 15;
  let node = null;
  s.doc.descendants((n) => { if (!node && n.type.name === 'image') node = n; });
  assert.ok(Math.abs(node.attrs.width - textWidth * 0.4) < 0.01);
  assert.equal(node.attrs.height, 50);
  // Saved unedited, the frame keeps that size, so the writer leaves it alone.
  assert.match(strFromU8(unzipSync(s.save())[main]), new RegExp(`<wp:extent cx="${Math.round(textWidth * 0.4 * 9525)}" cy="476250"`));
});

test('paragraph identities (w14:paraId, w14:textId, w:rsid*) are kept; a split paragraph does not repeat its paraId', () => {
  const files = unzipSync(blankDocxBytes());
  const main = new DocxSession(blankDocxBytes()).pkg.mainPath;
  const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
  const p = (id, t) => `<w:p w14:paraId="${id}" w14:textId="77777777" w:rsidR="00A1B2C3" w:rsidRDefault="00D4E5F6"><w:r><w:t>${t}</w:t></w:r></w:p>`;
  const xml = strFromU8(files[main])
    .replace(/<w:document\b/, `<w:document xmlns:w14="${W14}"`)
    .replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${p('1A2B3C4D', 'First paragraph')}${p('5E6F7081', 'Second')}$1`);
  files[main] = strToU8(xml);
  const s = new DocxSession(zipSync(files));
  const out = strFromU8(unzipSync(s.save())[main]);
  assert.match(out, /<w:p w14:paraId="1A2B3C4D" w14:textId="77777777" w:rsidR="00A1B2C3" w:rsidRDefault="00D4E5F6">/);
  assert.match(out, /<w:p w14:paraId="5E6F7081"/);
  // Enter in the middle of the first paragraph: both halves carry its attributes in the model; only one keeps the id.
  const tr = new Transform(s.doc).split(6);
  const split = strFromU8(unzipSync(s.save(tr.doc))[main]);
  assert.equal((split.match(/w14:paraId="1A2B3C4D"/g) || []).length, 1);
  assert.equal((split.match(/w:rsidR="00A1B2C3"/g) || []).length, 3);
  parseXml(split);
  // A part whose root does not declare w14 gets no w14 attributes (they would be unbound prefixes).
  files[main] = strToU8(xml.replace(` xmlns:w14="${W14}"`, '').replace(/ w14:\w+="[^"]*"/g, ''));
  const s2 = new DocxSession(zipSync(files));
  const t2 = new Transform(s2.doc);
  t2.setNodeMarkup(0, null, { ...s2.doc.firstChild.attrs, pattrs: [['w14:paraId', '1A2B3C4D'], ['w:rsidR', '00A1B2C3']] });
  const out2 = strFromU8(unzipSync(s2.save(t2.doc))[main]);
  assert.doesNotMatch(out2, /w14:/);
  assert.match(out2, /<w:p w:rsidR="00A1B2C3">/);
});

test('a footnote nothing refers to is kept by a save; one whose reference is deleted in the editor is dropped', () => {
  const files = unzipSync(blankDocxBytes());
  const main = new DocxSession(blankDocxBytes()).pkg.mainPath;
  const note = (id, t) => `<w:footnote w:id="${id}"><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:footnote>`;
  const notes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
    + note(1, 'Never referenced') + note(2, 'Referenced') + '</w:footnotes>';
  files['word/footnotes.xml'] = strToU8(notes);
  const ct = '[Content_Types].xml', rels = 'word/_rels/document.xml.rels';
  files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'));
  files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>'));
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, '<w:body><w:p><w:r><w:t>Text</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p>$1'));
  const s = new DocxSession(zipSync(files));
  assert.equal(strFromU8(unzipSync(s.save())['word/footnotes.xml']), notes);
  let at = null;
  s.doc.descendants((n, pos) => { if (/footnoteReference/.test(n.attrs.xml || '')) at = [pos, pos + n.nodeSize]; });
  const out = strFromU8(unzipSync(s.save(new Transform(s.doc).delete(...at).doc))['word/footnotes.xml']);
  assert.match(out, /Never referenced/);
  assert.doesNotMatch(out, />Referenced</);
});

// Saved again and again with no edits, a file settles at once: the second save writes what the first did, so nothing
// creeps in (or drops out) a little more each time a document goes through the editor.
test('five save cycles: every fixture saves the same from the first save on', () => {
  for (const f of fixtures) {
    let bytes = load(f);
    let prev = null;
    let doc = null;
    for (let i = 0; i < 5; i++) {
      const s = new DocxSession(bytes);
      if (doc) assert.deepEqual(s.doc.toJSON(), doc, `${f}: the document read back differs on cycle ${i + 1}`);
      doc = s.doc.toJSON();
      bytes = s.save();
      const parts = unzipSync(bytes);
      const now = Object.fromEntries(Object.entries(parts).map(([k, v]) => [k, strFromU8(v)]));
      if (prev) for (const k of new Set([...Object.keys(now), ...Object.keys(prev)])) assert.equal(now[k], prev[k], `${f}: ${k} changed on save ${i + 1}`);
      prev = now;
    }
  }
});
