import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { EditorState, TextSelection } from 'prosemirror-state';
import { CellSelection } from 'prosemirror-tables';
import { DocxSession } from '../src/ooxml/docx.js';
import { blankDocxBytes } from '../src/ooxml/template.js';
import { trackTransaction, resolveChanges, findChange, resolveChange } from '../src/editor/track.js';

const REV = 'w:id="7" w:author="Ann" w:date="2024-01-02T03:04:05Z"';
// One paragraph: "Keep " + inserted "new " + deleted "old " + "end" (bold by tracked formatting).
const BODY = `<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange ${REV.replace('7', '8')}><w:pPr/></w:pPrChange></w:pPr>`
  + '<w:r><w:t xml:space="preserve">Keep </w:t></w:r>'
  + `<w:ins ${REV}><w:r><w:t xml:space="preserve">new </w:t></w:r></w:ins>`
  + `<w:del ${REV.replace('7', '9')}><w:r><w:delText xml:space="preserve">old </w:delText></w:r></w:del>`
  + `<w:r><w:rPr><w:b/><w:rPrChange ${REV.replace('7', '10')}><w:rPr/></w:rPrChange></w:rPr><w:t>end</w:t></w:r></w:p>`;

function docx(body) {
  const files = unzipSync(blankDocxBytes());
  const main = new DocxSession(blankDocxBytes()).pkg.mainPath;
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
  return zipSync(files);
}
const text = (doc) => doc.textContent;
const xmlOf = (s, doc) => strFromU8(unzipSync(s.save(doc))[s.pkg.mainPath]);

test('tracked changes open and save as Word revisions', () => {
  const s = new DocxSession(docx(BODY));
  assert.equal(text(s.doc), 'Keep new old end');
  const out = xmlOf(s, s.doc);
  assert.match(out, /<w:ins w:id="7" w:author="Ann" w:date="2024-01-02T03:04:05Z"><w:r><w:t xml:space="preserve">new <\/w:t><\/w:r><\/w:ins>/);
  assert.match(out, /<w:del w:id="9"[^>]*><w:r><w:delText xml:space="preserve">old <\/w:delText><\/w:r><\/w:del>/);
  assert.match(out, /<w:rPrChange w:id="10"/);
  assert.match(out, /<w:pPrChange w:id="8"/);
});

test('accept all and reject all', () => {
  const s = new DocxSession(docx(BODY));
  const st = EditorState.create({ doc: s.doc });
  const acc = resolveChanges(st.tr, 0, st.doc.content.size, true).doc;
  assert.equal(text(acc), 'Keep new end');
  assert.equal(findChange(acc, 0, 1), null);
  assert.equal(acc.firstChild.attrs.align, 'center');
  assert.doesNotMatch(xmlOf(s, acc), /w:ins|w:del|Change/);
  const rej = resolveChanges(st.tr, 0, st.doc.content.size, false).doc;
  assert.equal(text(rej), 'Keep old end');
  assert.equal(rej.firstChild.attrs.align, null);
  assert.doesNotMatch(xmlOf(s, rej), /w:ins|w:del|Change|<w:b\/>/);
});

test('edits are recorded while tracking', () => {
  const s = new DocxSession(docx('<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>'));
  let st = EditorState.create({ doc: s.doc });
  const apply = (tr) => { st = st.apply(trackTransaction(st, tr, 'Bob')); };
  // Type at the end, then delete the original "world" and part of the insertion.
  apply(st.tr.setSelection(TextSelection.create(st.doc, 12)).insertText('!!'));
  assert.equal(text(st.doc), 'Hello world!!');
  apply(st.tr.delete(13, 14)); // Bob removes his own "!" outright
  assert.equal(text(st.doc), 'Hello world!');
  apply(st.tr.delete(7, 12)); // "world" stays, marked deleted
  assert.equal(text(st.doc), 'Hello world!');
  const out = xmlOf(s, st.doc);
  assert.match(out, /<w:del w:id="\d+" w:author="Bob"[^>]*><w:r><w:delText xml:space="preserve">world<\/w:delText><\/w:r><\/w:del><w:ins w:id="\d+" w:author="Bob"[^>]*><w:r><w:t xml:space="preserve">!<\/w:t><\/w:r><\/w:ins>/);
  // Enter inside tracked text inserts a paragraph mark.
  apply(st.tr.split(3));
  assert.equal(st.doc.childCount, 2);
  assert.match(xmlOf(s, st.doc), /<w:pPr><w:rPr><w:ins w:id="\d+" w:author="Bob"/);
  // Rejecting everything restores the original.
  const rej = resolveChanges(st.tr, 0, st.doc.content.size, false).doc;
  assert.equal(rej.childCount, 1);
  assert.equal(text(rej), 'Hello world');
});

test('formatting and paragraph changes are recorded while tracking', () => {
  const s = new DocxSession(docx('<w:p><w:r><w:t>Hello</w:t></w:r></w:p>'));
  let st = EditorState.create({ doc: s.doc });
  const apply = (tr) => { st = st.apply(trackTransaction(st, tr, 'Bob')); };
  apply(st.tr.addMark(1, 6, st.schema.marks.bold.create()));
  apply(st.tr.setNodeMarkup(0, null, { ...st.doc.firstChild.attrs, align: 'right' }));
  const out = xmlOf(s, st.doc);
  assert.match(out, /<w:b\/><w:rPrChange w:id="\d+" w:author="Bob"[^>]*><w:rPr\/><\/w:rPrChange>/);
  assert.match(out, /<w:jc w:val="right"\/><w:pPrChange w:id="\d+" w:author="Bob"[^>]*><w:pPr\/><\/w:pPrChange>/);
  const rej = resolveChanges(st.tr, 0, st.doc.content.size, false).doc;
  assert.doesNotMatch(xmlOf(s, rej), /<w:b\/>|w:jc|Change/);
});

test('Track Changes setting is saved', () => {
  const s = new DocxSession(blankDocxBytes());
  if (!s.settingsPath) return; // template without settings.xml
  s.trackRevisions = true;
  const bytes = s.save();
  assert.match(strFromU8(unzipSync(bytes)[s.settingsPath]), /<w:trackRevisions\/>/);
  assert.equal(new DocxSession(bytes).trackRevisions, true);
});

test('joining paragraphs while tracking deletes the paragraph mark', () => {
  const s = new DocxSession(docx('<w:p><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p>'));
  let st = EditorState.create({ doc: s.doc });
  st = st.apply(trackTransaction(st, st.tr.join(5), 'Bob'));
  assert.equal(st.doc.childCount, 2);
  assert.match(xmlOf(s, st.doc), /<w:pPr><w:rPr><w:del w:id="\d+" w:author="Bob"[^>]*\/><\/w:rPr><\/w:pPr><w:r><w:t xml:space="preserve">One/);
  const acc = resolveChanges(st.tr, 0, st.doc.content.size, true).doc;
  assert.equal(acc.childCount, 1);
  assert.equal(text(acc), 'OneTwo');
  const rej = resolveChanges(st.tr, 0, st.doc.content.size, false).doc;
  assert.equal(rej.childCount, 2);
  assert.doesNotMatch(xmlOf(s, rej), /w:del/);
});

// Moved text: "Alpha " moved from the start to the end, Word's w:moveFrom / w:moveTo between named range markers.
const MOVE = '<w:p><w:moveFromRangeStart w:id="20" w:author="Ann" w:date="2024-01-02T03:04:05Z" w:name="move1"/>'
  + `<w:moveFrom ${REV.replace('7', '21')}><w:r><w:t xml:space="preserve">Alpha </w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="20"/>`
  + '<w:r><w:t xml:space="preserve">Beta </w:t></w:r><w:moveToRangeStart w:id="22" w:author="Ann" w:date="2024-01-02T03:04:05Z" w:name="move1"/>'
  + `<w:moveTo ${REV.replace('7', '23')}><w:r><w:t xml:space="preserve">Alpha </w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="22"/></w:p>`;

test('moved text opens, saves as a move, and accepts or rejects as one', () => {
  const s = new DocxSession(docx(MOVE));
  assert.equal(text(s.doc), 'Alpha Beta Alpha ');
  const out = xmlOf(s, s.doc);
  assert.match(out, /<w:moveFrom w:id="21" w:author="Ann"[^>]*><w:r><w:t xml:space="preserve">Alpha <\/w:t><\/w:r><\/w:moveFrom>/); // (w:t, as Word writes it)
  assert.match(out, /<w:moveTo w:id="23" w:author="Ann"[^>]*><w:r><w:t xml:space="preserve">Alpha <\/w:t><\/w:r><\/w:moveTo>/);
  assert.match(out, /<w:moveFromRangeStart w:id="20"/);
  const st = EditorState.create({ doc: s.doc });
  const acc = resolveChanges(st.tr, 0, st.doc.content.size, true).doc;
  assert.equal(text(acc), 'Beta Alpha ');
  assert.doesNotMatch(xmlOf(s, acc), /w:move/);
  const rej = resolveChanges(st.tr, 0, st.doc.content.size, false).doc;
  assert.equal(text(rej), 'Alpha Beta ');
  assert.doesNotMatch(xmlOf(s, rej), /w:move/);
});

const row = (t, trPr = '') => `<w:tr>${trPr}<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc></w:tr>`;
const TABLE = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>'
  + row('Kept') + row('Added', `<w:trPr><w:ins ${REV.replace('7', '30')}/></w:trPr>`) + row('Removed', `<w:trPr><w:del ${REV.replace('7', '31')}/></w:trPr>`)
  + '</w:tbl><w:p/>';

test('inserted and deleted table rows open, save, and accept or reject', () => {
  const s = new DocxSession(docx(TABLE));
  const out = xmlOf(s, s.doc);
  assert.match(out, /<w:trPr><w:ins w:id="30"[^>]*\/><\/w:trPr>/);
  assert.match(out, /<w:trPr><w:del w:id="31"[^>]*\/><\/w:trPr>/);
  assert.ok(findChange(s.doc, 0, 1), 'a row revision is a change to go to');
  const st = EditorState.create({ doc: s.doc });
  const acc = resolveChanges(st.tr, 0, st.doc.content.size, true).doc;
  assert.equal(text(acc), 'KeptAdded');
  assert.doesNotMatch(xmlOf(s, acc), /<w:ins|<w:del/);
  const rej = resolveChanges(st.tr, 0, st.doc.content.size, false).doc;
  assert.equal(text(rej), 'KeptRemoved');
  assert.doesNotMatch(xmlOf(s, rej), /<w:ins|<w:del/);
});

test('a table whose rows are all rejected goes whole', () => {
  const s = new DocxSession(docx('<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' + row('New', `<w:trPr><w:ins ${REV}/></w:trPr>`) + '</w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p>'));
  const st = EditorState.create({ doc: s.doc });
  const rej = resolveChanges(st.tr, 0, st.doc.content.size, false).doc;
  assert.equal(rej.childCount, 1);
  assert.equal(text(rej), 'After');
});

test('rows inserted and deleted while tracking are recorded as row revisions', () => {
  const s = new DocxSession(docx('<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' + row('One') + row('Two') + '</w:tbl><w:p/>'));
  let st = EditorState.create({ doc: s.doc });
  const table = st.doc.firstChild;
  // Insert a copy of the first row after it, as Insert row below does.
  const after = 1 + table.child(0).nodeSize;
  st = st.apply(trackTransaction(st, st.tr.insert(after, table.child(0).type.create(null, table.child(0).content)), 'Bob'));
  assert.match(xmlOf(s, st.doc), /<w:tr><w:trPr><w:ins w:id="\d+" w:author="Bob"[^>]*\/><\/w:trPr>/);
  // Delete the last row, as Delete row does: it stays, marked deleted, its text too.
  const t2 = st.doc.firstChild;
  let pos = 1;
  for (let i = 0; i < t2.childCount - 1; i++) pos += t2.child(i).nodeSize;
  st = st.apply(trackTransaction(st, st.tr.delete(pos, pos + t2.lastChild.nodeSize), 'Bob'));
  assert.equal(st.doc.firstChild.childCount, 3);
  const out = xmlOf(s, st.doc);
  assert.match(out, /<w:trPr><w:del w:id="\d+" w:author="Bob"[^>]*\/><\/w:trPr><w:tc>.*?<w:delText xml:space="preserve">Two<\/w:delText>/);
  const acc = resolveChanges(st.tr, 0, st.doc.content.size, true).doc;
  assert.equal(text(acc), 'OneOne');
});

test('deleting a row you inserted while tracking removes it outright', () => {
  const s = new DocxSession(docx('<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' + row('One') + row('Two') + '</w:tbl><w:p/>'));
  let st = EditorState.create({ doc: s.doc });
  const first = st.doc.firstChild.child(0);
  const after = 1 + first.nodeSize;
  st = st.apply(trackTransaction(st, st.tr.insert(after, first.type.create(null, first.content)), 'Bob'));
  assert.equal(st.doc.firstChild.childCount, 3);
  const added = st.doc.firstChild.child(1);
  st = st.apply(trackTransaction(st, st.tr.delete(after, after + added.nodeSize), 'Bob'));
  assert.equal(st.doc.firstChild.childCount, 2);
  assert.equal(findChange(st.doc, 0, 1), null);
});

// A row with a deleted cell (w:cellDel) and a row with an inserted one (w:cellIns), three grid columns.
const tc = (t, rev = '') => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${rev}</w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
const CELLS = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
  + `<w:tr>${tc('A1')}${tc('B1', `<w:cellDel ${REV.replace('7', '40')}/>`)}${tc('C1')}</w:tr>`
  + `<w:tr>${tc('A2')}${tc('B2', `<w:cellIns ${REV.replace('7', '41')}/>`)}${tc('C2')}</w:tr>`
  + '</w:tbl><w:p/>';

test('tracked cell insertions and deletions accept and reject, the row shifting left', () => {
  const s = new DocxSession(docx(CELLS));
  assert.ok(findChange(s.doc, 0, 1), 'a cell revision is a change to go to');
  const st = EditorState.create({ doc: s.doc });
  const acc = resolveChanges(st.tr, 0, st.doc.content.size, true).doc;
  assert.equal(text(acc), 'A1C1A2B2C2');
  const accXml = xmlOf(s, acc);
  assert.match(accXml, /<w:tr><w:trPr><w:gridAfter w:val="1"\/><\/w:trPr><w:tc>.*?A1.*?<\/w:tc><w:tc>.*?C1.*?<\/w:tc><\/w:tr>/);
  assert.doesNotMatch(accXml, /cellIns|cellDel/);
  const rej = resolveChanges(st.tr, 0, st.doc.content.size, false).doc;
  assert.equal(text(rej), 'A1B1C1A2C2');
  const rejXml = xmlOf(s, rej);
  assert.match(rejXml, /<w:tr><w:trPr><w:gridAfter w:val="1"\/><\/w:trPr><w:tc>.*?A2.*?<\/w:tc><w:tc>.*?C2/);
  assert.doesNotMatch(rejXml, /cellIns|cellDel/);
});

test('rejecting at the caret inside a deleted row (its text deleted too) puts back the row, not only its text', () => {
  // (Two cells: the caret's run is less than the row.)
  const cell = (t, id) => `<w:tc><w:p><w:del ${REV.replace('7', id)}><w:r><w:delText>${t}</w:delText></w:r></w:del></w:p></w:tc>`;
  const del = `<w:tr><w:trPr><w:del ${REV.replace('7', '41')}/></w:trPr>${cell('Gone', '42')}${cell('Too', '43')}</w:tr>`;
  const kept = '<w:tr><w:tc><w:p><w:r><w:t>Kept</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr>';
  const s = new DocxSession(docx(`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${kept}${del}</w:tbl><w:p/>`));
  let at = null;
  s.doc.descendants((n, pos) => { if (at == null && n.isText && n.text === 'Gone') at = pos + 2; });
  let st = EditorState.create({ doc: s.doc, selection: TextSelection.create(s.doc, at) });
  assert.ok(resolveChange(false)(st, (tr) => { st = st.apply(tr); }));
  assert.equal(text(st.doc), 'KeptGoneToo');
  assert.doesNotMatch(xmlOf(s, st.doc), /<w:del\b/, 'neither the text nor the row is still deleted');
  // So accepting everything afterwards keeps the row.
  assert.equal(text(resolveChanges(st.tr, 0, st.doc.content.size, true).doc), 'KeptGoneToo');
});

test('accepting the deletion of a whole paragraph leaves the next one with its own style, list and alignment', () => {
  const s = new DocxSession(docx(`<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr><w:rPr><w:del ${REV}/></w:rPr></w:pPr>`
    + `<w:del ${REV}><w:r><w:delText>Gone</w:delText></w:r></w:del></w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Stays</w:t></w:r></w:p>`));
  const doc = resolveChanges(EditorState.create({ doc: s.doc }).tr, 0, s.doc.content.size, true).doc;
  assert.equal(doc.childCount, 1);
  assert.deepEqual([doc.firstChild.attrs.styleId, doc.firstChild.attrs.numId, doc.firstChild.attrs.align, doc.firstChild.textContent], [null, null, 'center', 'Stays']);
});

test('accepting or rejecting a paragraph moved whole leaves no empty paragraph behind', () => {
  const mv = (kind, id, t) => `<w:p><w:pPr><w:rPr><w:${kind} w:id="${id}" w:author="Ann" w:date="2024-01-02T03:04:05Z"/></w:rPr></w:pPr><w:${kind} w:id="${id}" w:author="Ann" w:date="2024-01-02T03:04:05Z"><w:r><w:t>${t}</w:t></w:r></w:${kind}></w:p>`;
  const s = new DocxSession(docx(`${mv('moveFrom', 1, 'Moved')}<w:p><w:r><w:t>Middle</w:t></w:r></w:p>${mv('moveTo', 2, 'Moved')}<w:p><w:r><w:t>End</w:t></w:r></w:p>`));
  for (const [accept, want] of [[true, ['Middle', 'Moved', 'End']], [false, ['Moved', 'Middle', 'End']]]) {
    const doc = resolveChanges(EditorState.create({ doc: s.doc }).tr, 0, s.doc.content.size, accept).doc;
    const paras = [];
    doc.forEach((p) => paras.push(p.textContent));
    assert.deepEqual(paras, want);
    assert.doesNotMatch(xmlOf(s, doc), /w:move(From|To)\b/);
  }
});

test('accepting a cell selection resolves the changes of every selected cell', () => {
  const tc = (inner) => `<w:tc><w:p>${inner}</w:p></w:tc>`;
  const ch = (id, old, now) => `<w:del w:id="${id}" w:author="Ann" w:date="2024-01-02T03:04:05Z"><w:r><w:delText>${old}</w:delText></w:r></w:del><w:ins w:id="${id + 1}" w:author="Ann" w:date="2024-01-02T03:04:05Z"><w:r><w:t>${now}</w:t></w:r></w:ins>`;
  const s = new DocxSession(docx(`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${tc(ch(1, '1,000', '900'))}${tc(ch(3, 'Net 30', 'Net 45'))}${tc('<w:r><w:t>plain</w:t></w:r>')}</w:tr></w:tbl><w:p/>`));
  const cells = [];
  s.doc.descendants((n, pos) => { if (n.type.name === 'table_cell') cells.push(pos); });
  for (const [a, b, want] of [[0, 1, ['900', 'Net 45', 'plain']], [1, 2, ['1,000900', 'Net 45', 'plain']]]) {
    const st0 = EditorState.create({ doc: s.doc });
    const st = st0.apply(st0.tr.setSelection(CellSelection.create(s.doc, cells[a], cells[b])));
    let out = st;
    resolveChange(true)(st, (tr) => { out = st.apply(tr); });
    const texts = [];
    out.doc.descendants((n) => { if (n.type.name === 'table_cell') { texts.push(n.textContent); return false; } return true; });
    assert.deepEqual(texts, want);
  }
});

test('a row inserted below takes the formatting of the row it is made from', async () => {
  const { addRow } = await import('../src/editor/commands.js');
  const tc = (t) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="right"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
  const s = new DocxSession(docx(`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="500"/></w:trPr>${tc('a')}${tc('b')}</w:tr></w:tbl><w:p/>`));
  let at = null;
  s.doc.descendants((n, pos) => { if (at == null && n.isText) at = pos; });
  const st = EditorState.create({ doc: s.doc, selection: TextSelection.create(s.doc, at) });
  let out = st;
  assert.ok(addRow(true)(st, (tr) => { out = st.apply(tr); }));
  const table = out.doc.firstChild;
  assert.equal(table.childCount, 2);
  const row = table.child(1);
  assert.match(row.attrs.trPr || '', /trHeight w:val="500"/);
  row.forEach((cell) => {
    assert.ok(cell.attrs.tcPr.some(([n, x]) => n === 'shd' && /D9D9D9/.test(x)));
    assert.equal(cell.firstChild.attrs.align, 'right');
    assert.equal(cell.textContent, '');
  });
});

test('a revision keeps its UTC time (w16du:dateUtc); dates are written and read as Word\'s wall-clock time', async () => {
  const { now, wordDate } = await import('../src/editor/util.js');
  const W16DU = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du';
  const files = unzipSync(docx(`<w:p><w:ins w:id="3" w:author="Ann" w:date="2024-01-02T09:00:00Z" w16du:dateUtc="2024-01-02T14:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins></w:p>`));
  const main = new DocxSession(blankDocxBytes()).pkg.mainPath;
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:document\b/, `<w:document xmlns:w16du="${W16DU}"`));
  const s = new DocxSession(zipSync(files));
  assert.match(xmlOf(s, s.doc), /<w:ins w:id="3" w:author="Ann" w:date="2024-01-02T09:00:00Z" w16du:dateUtc="2024-01-02T14:00:00Z">/);
  // Shown: the UTC time when there is one; else w:date as the wall-clock time it is.
  assert.equal(wordDate('2024-01-02T09:00:00Z', '2024-01-02T14:00:00Z').toISOString(), '2024-01-02T14:00:00.000Z');
  assert.equal(wordDate('2024-01-02T09:00:00Z').getHours(), 9);
  const d = new Date();
  assert.equal(Number(now().slice(11, 13)), d.getHours() === 23 && d.getMinutes() === 59 ? Number(now().slice(11, 13)) : d.getHours());
});

test('a column inserted with changes tracked is recorded as cell insertions; rejected, the table is as it was', async () => {
  const { addColumnAfter } = await import('prosemirror-tables');
  const tc = (t) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
  const s = new DocxSession(docx(`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${tc('a')}${tc('b')}</w:tr><w:tr>${tc('c')}${tc('d')}</w:tr></w:tbl><w:p/>`));
  const st = EditorState.create({ doc: s.doc, selection: TextSelection.create(s.doc, 4) });
  let out = st;
  addColumnAfter(st, (tr) => { out = st.apply(trackTransaction(st, tr, 'Me')); });
  const tracked = xmlOf(s, out.doc);
  assert.equal((tracked.match(/<w:cellIns\b/g) || []).length, 2);
  assert.equal((tracked.match(/<w:gridCol\b/g) || []).length, 3); // (the grid follows the table)
  const back = xmlOf(s, resolveChanges(out.tr, 0, out.doc.content.size, false).doc);
  assert.equal((back.match(/<w:tc>/g) || []).length, 4);
  assert.equal((back.match(/<w:gridCol\b/g) || []).length, 2);
  assert.doesNotMatch(back, /gridAfter|cellIns/);
});
