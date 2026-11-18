// More of Word's custom XML bindings (test/databind.test.js has the rest), each checked against Word: an empty node
// shows the control's placeholder from the glossary, over cached text too; a control around a paragraph or a cell
// emptied in the editor shows it as well; a repeating section gets an item for each node it is bound to; a picture
// control shows the picture its node holds, fitted in its box. And the same with crafted values.
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { unzipSync, strFromU8 } from 'fflate';
import { EditorState } from 'prosemirror-state';
import { DocxSession } from '../src/ooxml/docx.js';
import { listFields } from '../src/editor/fields.js';
import { bindingsPlugin } from '../src/editor/bindings.js';
import { bindingDocx, bound } from './fixtures/mkbinding.mjs';

const P = (x) => `<w:p>${x}</w:p>`;
const R = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const w15 = (bytes) => {
  const f = unzipSync(bytes);
  f['word/document.xml'] = new TextEncoder().encode(strFromU8(f['word/document.xml']).replace(/<w:document\b/, `<w:document xmlns:w15="${W15}"`));
  return f;
};
const zip = async (f) => (await import('fflate')).zipSync(f);
const paras = (doc) => { const out = []; doc.descendants((n) => { if (n.type.name === 'paragraph') { out.push(n.textContent); return false; } return true; }); return out; };

test('an empty node shows the control\'s placeholder from the glossary (or Word\'s default), over cached text too', () => {
  const body = [
    bound({ xpath: '/r/e', plc: false, content: R('cached'), docPart: 'PH1' }),
    bound({ xpath: '/r/e', plc: false, content: R('cached') }),
    bound({ xpath: '/r/e', plc: false, content: R('cached'), kind: '<w:date><w:dateFormat w:val="M/d/yyyy"/></w:date>' }),
    bound({ xpath: '/r/missing', plc: false, content: R('kept') }),
  ].map(P).join('') + `<w:sdt><w:sdtPr><w:id w:val="5"/><w:placeholder><w:docPart w:val="PH1"/></w:placeholder><w:text/><w:dataBinding w:xpath="/r/e"/></w:sdtPr><w:sdtContent>${P(R('block cached'))}</w:sdtContent></w:sdt>`;
  const s = new DocxSession(bindingDocx({ body, parts: [{ xml: '<r><e></e></r>' }], glossary: { PH1: 'Type the name here' } }));
  const f = listFields(s.doc);
  assert.deepEqual(f.map((x) => x.text), ['Type the name here', 'Click or tap here to enter text.', 'Click or tap to enter a date.', 'kept', 'Type the name here']);
  assert.deepEqual(f.map((x) => x.value), ['', '', '', 'kept', '']);
  assert.match(f[0].pr, /<w:showingPlcHdr\/><w:dataBinding/);
  // In the placeholder's formatting: its building block's (italic, PlaceholderText).
  let marks = null;
  s.doc.nodesBetween(f[0].from, f[0].to, (n) => { if (!marks && n.isText) marks = n.marks.map((m) => m.type.name + (m.attrs.val ? `:${m.attrs.val}` : '')); });
  assert.ok(marks.includes('italic') && marks.includes('cstyle:PlaceholderText'), marks.join());
  const out = strFromU8(unzipSync(s.save())['word/document.xml']);
  assert.equal((out.match(/<w:showingPlcHdr\/>/g) || []).length, 4);
});

test('a control around a paragraph or a cell, emptied, shows its placeholder; typed into, it holds a value again', () => {
  const cell = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:sdt><w:sdtPr><w:id w:val="9"/><w:text/><w:dataBinding w:xpath="/r/c"/></w:sdtPr><w:sdtContent><w:tc><w:p>' + R('old cell') + '</w:p></w:tc></w:sdtContent></w:sdt></w:tr></w:tbl>';
  const block = `<w:sdt><w:sdtPr><w:id w:val="8"/><w:text/><w:dataBinding w:xpath="/r/b"/></w:sdtPr><w:sdtContent>${P(R('old block'))}</w:sdtContent></w:sdt>`;
  const s = new DocxSession(bindingDocx({ body: block + cell + P(R('end')), parts: [{ xml: '<r><b>Block value</b><c>Cell value</c></r>' }] }));
  let state = EditorState.create({ doc: s.doc, plugins: [bindingsPlugin(s)] });
  assert.deepEqual(paras(state.doc), ['Block value', 'Cell value', 'end']);
  // Emptied: the node empties, the placeholder shows, flagged.
  state = state.apply(state.tr.delete(1, 1 + 'Block value'.length));
  let cellPos = null;
  state.doc.descendants((n, pos) => { if (n.type.spec.tableRole === 'cell') cellPos = { pos, n }; });
  state = state.apply(state.tr.delete(cellPos.pos + 2, cellPos.pos + 2 + 'Cell value'.length));
  assert.deepEqual(paras(state.doc), ['Click or tap here to enter text.', 'Click or tap here to enter text.', 'end']);
  assert.equal(s.bindings.value('<w:dataBinding w:xpath="/r/b"/>'), '');
  assert.equal(s.bindings.value('<w:dataBinding w:xpath="/r/c"/>'), '');
  assert.match(state.doc.firstChild.attrs.sdt.open, /<w:showingPlcHdr\/>/);
  // Typed into (the placeholder selected, as a click selects it): a value, unflagged, out of the placeholder style.
  const first = state.doc.firstChild;
  state = state.apply(state.tr.insertText('Typed', 1, first.nodeSize - 1));
  assert.equal(state.doc.firstChild.textContent, 'Typed');
  assert.equal(s.bindings.value('<w:dataBinding w:xpath="/r/b"/>'), 'Typed');
  assert.doesNotMatch(state.doc.firstChild.attrs.sdt.open, /showingPlcHdr/);
  let ph = false;
  state.doc.firstChild.forEach((n) => { if (n.marks.some((m) => m.type.name === 'cstyle')) ph = true; });
  assert.equal(ph, false);
  const files = unzipSync(s.save(state.doc));
  assert.equal(strFromU8(files['customXml/item1.xml']), '<r><b>Typed</b><c/></r>');
});

const item = (n, label) => `<w:sdt><w:sdtPr><w:id w:val="${100 + n}"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent>${P(R(label + ': ') + bound({ xpath: `/root[1]/items[1]/item[${n}]/name[1]`, plc: false, content: R('x'), rPr: '', tag: 'Name' }))}</w:sdtContent></w:sdt>`;
const section = (items, tag = 'w15:dataBinding') => `<w:sdt><w:sdtPr><w:id w:val="99"/><${tag} w:prefixMappings="" w:xpath="/root[1]/items[1]/item" w:storeItemID=""/><w15:repeatingSection/></w:sdtPr><w:sdtContent>${items}</w:sdtContent></w:sdt>`;
const ITEMS = (...names) => `<root><items>${names.map((n) => `<item><name>${n}</name></item>`).join('')}</items></root>`;

test('a repeating section gets an item for each node, as Word makes them, and each item edits its own node', async () => {
  for (const tag of ['w15:dataBinding', 'w:dataBinding']) {
    const bytes = await zip(w15(bindingDocx({ body: P(R('Before')) + section(item(1, 'Item'), tag) + P(R('After')), parts: [{ xml: ITEMS('Alpha', 'Beta', 'Gamma') }] })));
    const s = new DocxSession(bytes);
    assert.deepEqual(paras(s.doc), ['Before', 'Item: Alpha', 'Item: Beta', 'Item: Gamma', 'After'], tag);
    // Each item's control bound to its own node, its own id; the section written around them all.
    let state = EditorState.create({ doc: s.doc, plugins: [bindingsPlugin(s)] });
    const names = listFields(state.doc, { nested: true }).filter((f) => f.tag === 'Name');
    assert.equal(names.length, 3);
    state = state.apply(state.tr.insertText('Bravo', names[1].from, names[1].to));
    assert.equal(s.bindings.value('<w:dataBinding w:xpath="/root/items/item[2]/name"/>'), 'Bravo');
    assert.deepEqual(paras(state.doc), ['Before', 'Item: Alpha', 'Item: Bravo', 'Item: Gamma', 'After']);
    const out = s.save(state.doc);
    const xml = strFromU8(unzipSync(out)['word/document.xml']);
    assert.equal((xml.match(/<w15:repeatingSection\/>/g) || []).length, 1);
    assert.equal((xml.match(/<w15:repeatingSectionItem\/>/g) || []).length, 3);
    assert.match(xml, /item\[2\]\/name\[1\]"\/><\/w:sdtPr><w:sdtContent><w:r>.*?Bravo/);
    const ids = [...xml.matchAll(/<w:id w:val="(-?\d+)"\/>/g)].map((m) => m[1]);
    assert.equal(new Set(ids).size, ids.length, 'ids repeat');
    assert.equal(strFromU8(unzipSync(out)['customXml/item1.xml']), ITEMS('Alpha', 'Bravo', 'Gamma'));
    assert.deepEqual(paras(new DocxSession(out).doc), ['Before', 'Item: Alpha', 'Item: Bravo', 'Item: Gamma', 'After']);
  }
});

test('a repeating section with more items than nodes keeps one for each; table rows repeat too', async () => {
  const s = new DocxSession(await zip(w15(bindingDocx({ body: section(item(1, 'A') + item(2, 'B') + item(3, 'C')) + P(R('After')), parts: [{ xml: ITEMS('Only') }] }))));
  assert.deepEqual(paras(s.doc), ['A: Only', 'After']);
  const cell = (t) => `<w:tc><w:p>${t}</w:p></w:tc>`;
  const row = (n) => `<w:sdt><w:sdtPr><w:id w:val="${200 + n}"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent><w:tr>${cell(bound({ xpath: `/root[1]/items[1]/item[${n}]/name[1]`, plc: false, content: R('x'), rPr: '' }))}${cell(R('fee'))}</w:tr></w:sdtContent></w:sdt>`;
  const tbl = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell(R('Name'))}${cell(R('Fee'))}</w:tr>${section(row(1))}</w:tbl>`;
  const t = new DocxSession(await zip(w15(bindingDocx({ body: tbl + P(R('After')), parts: [{ xml: ITEMS('Hosting', 'Support') }] }))));
  assert.deepEqual(paras(t.doc), ['Name', 'Fee', 'Hosting', 'fee', 'Support', 'fee', 'After']);
  const xml = strFromU8(unzipSync(t.save())['word/document.xml']);
  assert.match(xml, /<w15:repeatingSection\/><\/w:sdtPr><w:sdtContent><w:sdt>.*?Hosting.*?<\/w:sdt><w:sdt>.*?item\[2\]\/name\[1\].*?Support.*?<\/w:sdt><\/w:sdtContent><\/w:sdt><\/w:tbl>/);
});

// A PNG's first bytes, giving its size (all a picture control's binding reads of it here).
const png = (w, h) => {
  const b = new Uint8Array(33);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  new DataView(b.buffer).setUint32(16, w);
  new DataView(b.buffer).setUint32(20, h);
  return b;
};
const b64 = (u8) => Buffer.from(u8).toString('base64');
const drawing = (rId) => `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="1905000" cy="1905000"/><wp:docPr id="1" name="Picture 1"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1905000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
const picCtl = `<w:sdt><w:sdtPr><w:id w:val="3"/><w:picture/><w:dataBinding w:xpath="/r/pic"/></w:sdtPr><w:sdtContent>${P(drawing('rIdP1'))}</w:sdtContent></w:sdt>`;

test('a picture control shows the picture its node holds, fitted in its box; the picture it showed is left', () => {
  const old = png(1, 1);
  const s = new DocxSession(bindingDocx({ body: picCtl + P(R('end')), parts: [{ xml: `<r><pic>${b64(png(40, 20))}</pic></r>` }], media: { 'media/image1.png': old } }));
  const files = unzipSync(s.save());
  const xml = strFromU8(files['word/document.xml']);
  assert.match(xml, /<wp:extent cx="1905000" cy="952500"\/>/);
  assert.match(xml, /<a:ext cx="1905000" cy="952500"\/>/);
  const rId = /r:embed="([^"]+)"/.exec(xml)[1];
  assert.notEqual(rId, 'rIdP1');
  const rels = strFromU8(files['word/_rels/document.xml.rels']);
  const target = new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(rels)?.[1] || new RegExp(`Target="([^"]+)"[^>]*Id="${rId}"`).exec(rels)?.[1];
  assert.deepEqual(files[`word/${target}`], png(40, 20));
  assert.deepEqual(files['word/media/image1.png'], old);
  assert.match(strFromU8(files['[Content_Types].xml']), new RegExp(`PartName="/word/${target}" ContentType="image/png"`));
  // The same picture again: nothing to change.
  const same = new DocxSession(bindingDocx({ body: picCtl, parts: [{ xml: `<r><pic>${b64(old)}</pic></r>` }], media: { 'media/image1.png': old } }));
  assert.match(strFromU8(unzipSync(same.save())['word/document.xml']), /r:embed="rIdP1"/);
  // Not a picture, or not base64: left as it was.
  for (const v of ['not base64 !!', b64(new TextEncoder().encode('<svg/>')), '']) {
    const bad = new DocxSession(bindingDocx({ body: picCtl, parts: [{ xml: `<r><pic>${v}</pic></r>` }], media: { 'media/image1.png': old } }));
    assert.match(strFromU8(unzipSync(bad.save())['word/document.xml']), /r:embed="rIdP1"[\s\S]*cy="1905000"/);
  }
});

test('crafted repeating sections and pictures stay bounded', async () => {
  const many = `<root><items>${'<item><name>n</name></item>'.repeat(50000)}</items></root>`;
  const t0 = Date.now();
  const s = new DocxSession(await zip(w15(bindingDocx({ body: section(item(1, 'I')), parts: [{ xml: many }] }))));
  assert.equal(paras(s.doc).length, 1000);
  assert.ok(Date.now() - t0 < 20000, `${Date.now() - t0}ms`);
  // A truncated JPEG and a GIF claiming a huge size: read without running off, sized to the box.
  const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x02, 0xff, 0xe0, 0xff, 0xff]);
  const gif = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0xff, 0xff, 0x01, 0x00, 0, 0]);
  for (const bytes of [jpeg, gif]) {
    const p = new DocxSession(bindingDocx({ body: picCtl, parts: [{ xml: `<r><pic>${b64(bytes)}</pic></r>` }], media: { 'media/image1.png': png(1, 1) } }));
    const xml = strFromU8(unzipSync(p.save())['word/document.xml']);
    const [, cx, cy] = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(xml);
    assert.ok(Number(cx) <= 1905000 && Number(cy) <= 1905000, `${cx}x${cy}`);
  }
});

// The security review before 0.8.4.

test('a repeating section inside another is left as it is, however its file orders it', async () => {
  // (Each control's w:sdtPr after its content: the inner section's marker comes first in the file.)
  const itemOf = (x) => `<w:sdt><w:sdtContent>${x}</w:sdtContent><w:sdtPr><w15:repeatingSectionItem/></w:sdtPr></w:sdt>`;
  const sectionOf = (x) => `<w:sdt><w:sdtContent>${x}</w:sdtContent><w:sdtPr><w15:repeatingSection/><w15:dataBinding w:xpath="/r/i"/></w:sdtPr></w:sdt>`;
  const body = sectionOf(itemOf(sectionOf(itemOf(P(R('x')))))) + P(R('After'));
  const t0 = Date.now();
  const s = new DocxSession(await zip(w15(bindingDocx({ body, parts: [{ xml: `<r>${'<i/>'.repeat(1000)}</r>` }] }))));
  assert.ok(Date.now() - t0 < 20000, `${Date.now() - t0}ms`);
  assert.ok(paras(s.doc).length <= 1001, `${paras(s.doc).length} paragraphs`);
});

test("repeating sections' new items stay within one budget of copied elements", async () => {
  const big = P(R('x')).repeat(400);
  const many = `<root><items>${'<item><name>n</name></item>'.repeat(1000)}</items></root>`;
  const body = Array.from({ length: 20 }, () => section(`<w:sdt><w:sdtPr><w:id w:val="7"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent>${big}</w:sdtContent></w:sdt>`)).join('');
  const t0 = Date.now();
  const s = new DocxSession(await zip(w15(bindingDocx({ body, parts: [{ xml: many }] }))));
  assert.ok(Date.now() - t0 < 20000, `${Date.now() - t0}ms`);
  // (20 sections of 1,000 items of 400 paragraphs would be 8 million; the budget is 200,000 elements.)
  assert.ok(paras(s.doc).length < 80000, `${paras(s.doc).length} paragraphs`);
});

test('one picture shown by many controls is decoded and stored once', () => {
  const body = Array.from({ length: 300 }, () => picCtl).join('') + P(R('end'));
  const s = new DocxSession(bindingDocx({ body, parts: [{ xml: `<r><pic>${b64(png(40, 20))}</pic></r>` }], media: { 'media/image1.png': png(1, 1) } }));
  const files = unzipSync(s.save());
  assert.equal(Object.keys(files).filter((f) => /^word\/media\/bound/.test(f)).length, 1);
  const xml = strFromU8(files['word/document.xml']);
  assert.equal(new Set([...xml.matchAll(/r:embed="([^"]+)"/g)].map((m) => m[1])).size, 1);
});

test("a placeholder building block's long run formatting isn't copied into controls", () => {
  const s = new DocxSession(bindingDocx({ body: P(bound({ xpath: '/r/e', plc: false, content: R('cached'), docPart: 'PH1' })), parts: [{ xml: '<r><e></e></r>' }], glossary: { PH1: 'Type here' } }));
  const store = s.bindings;
  const ph = store.placeholder('<w:placeholder><w:docPart w:val="PH1"/></w:placeholder>', 'text');
  assert.equal(ph.text, 'Type here');
  assert.ok(!ph.rPr || ph.rPr.getElementsByTagName('*').length <= 64);
});
