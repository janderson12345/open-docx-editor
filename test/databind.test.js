// Content controls bound to custom XML (w:dataBinding): shown with their nodes' values on opening, as Word shows them
// (checked against Word: a node's value replaces the placeholder, a single space too; an empty node or one not found
// leaves it; storeItemID picks the part when one has that id), and edits going into the node and every control bound
// to it, body and header, and into customXml/itemN.xml on saving.
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { unzipSync, strFromU8 } from 'fflate';
import { EditorState } from 'prosemirror-state';
import { DocxSession } from '../src/ooxml/docx.js';
import { compileXPath, parseXmlDate } from '../src/ooxml/databind.js';
import { listFields, formsPlugin } from '../src/editor/fields.js';
import { bindingsPlugin } from '../src/editor/bindings.js';
import { fillFields, readFieldValues } from '../src/fieldsNode.js';
import { bindingDocx, bound, FIXTURE, PH } from './fixtures/mkbinding.mjs';

const P = (x) => `<w:p>${x}</w:p>`;
const texts = (doc) => listFields(doc).map((f) => f.text);
const ID1 = '{11111111-2222-3333-4444-555555555555}', ID3 = '{33333333-2222-3333-4444-555555555555}';

test('XPaths Word writes are read; anything else is not', () => {
  assert.deepEqual(compileXPath('/root/a'), [{ ns: null, name: 'root', n: 1 }, { ns: null, name: 'a', n: 1 }]);
  assert.deepEqual(compileXPath("/ns0:root[1]/ns0:a[2]", "xmlns:ns0='urn:x'"), [{ ns: 'urn:x', name: 'root', n: 1 }, { ns: 'urn:x', name: 'a', n: 2 }]);
  assert.deepEqual(compileXPath('/root/a/@x').at(-1), { attr: true, ns: null, name: 'x' });
  for (const bad of ['root/a', '//a', '/root/*', '/root/a[@x="1"]', '/root/a[last()]', '/root/@x/a', '/ns1:root', '/root/a[0]', '/root/a b', '/' + 'a/'.repeat(100) + 'a', '/' + 'a'.repeat(3000)]) assert.equal(compileXPath(bad), null, bad);
});

test('a bound control shows its node\'s value, in its own formatting, as Word shows it on opening', () => {
  const root = '<root><a>Alpha</a><sp> </sp><empty></empty><b>Beta</b><b>Beta2</b><ml>line1\nline2</ml><dd>G</dd><cb>true</cb><att x="AttrVal"/></root>';
  const body = [
    bound({ xpath: '/root/a' }),
    bound({ xpath: '/root/sp' }),
    bound({ xpath: '/root/empty' }),
    bound({ xpath: '/root/nope' }),
    bound({ xpath: '/root/b' }),
    bound({ xpath: '/root[1]/b[2]' }),
    bound({ xpath: '/root/ml', kind: '<w:text w:multiLine="1"/>' }),
    bound({ xpath: '/root/dd', kind: '<w:dropDownList><w:listItem w:displayText="Green" w:value="G"/><w:listItem w:displayText="Red" w:value="R"/></w:dropDownList>' }),
    bound({ xpath: '/root/cb', plc: false, kind: '<w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>', content: '<w:r><w:t>☐</w:t></w:r>' }),
    bound({ xpath: '/root/att/@x' }),
    bound({ xpath: '/root/a', plc: false, content: '<w:r><w:rPr><w:i/></w:rPr><w:t>cached</w:t></w:r>' }),
    bound({ xpath: '/root/b', plc: false, content: '<w:r><w:rPr><w:i/></w:rPr><w:t>Beta</w:t></w:r>' }),
  ].map(P).join('');
  const s = new DocxSession(bindingDocx({ body, parts: [{ xml: root, id: ID1 }] }));
  assert.deepEqual(texts(s.doc), ['Alpha', ' ', '.', '.', 'Beta', 'Beta2', 'line1line2', 'Green', '☒', 'AttrVal', 'Alpha', 'Beta']);
  const fields = listFields(s.doc);
  // The placeholder gone for a value, a single space too; kept for an empty node and one not found.
  assert.deepEqual(fields.map((f) => f.mark.attrs.phText), [null, null, '.', '.', null, null, null, null, null, null, null, null]);
  assert.ok(!/showingPlcHdr/.test(fields[0].pr) && /showingPlcHdr/.test(fields[2].pr));
  // The control's own run formatting (bold, not the placeholder's style); one already showing its value keeps its own.
  const marks = (i) => { let m = null; s.doc.nodesBetween(fields[i].from, fields[i].to, (n) => { if (!m && n.isText) m = n.marks.map((x) => x.type.name); }); return m; };
  assert.ok(marks(0).includes('bold') && !marks(0).includes('cstyle'));
  assert.ok(marks(10).includes('bold') && !marks(10).includes('italic'));
  assert.ok(marks(11).includes('italic'));
  // A multi-line value's line break, a check box's state.
  let breaks = 0;
  s.doc.nodesBetween(fields[6].from, fields[6].to, (n) => { if (n.type.name === 'hard_break') breaks++; });
  assert.equal(breaks, 1);
  assert.equal(fields[8].checked, true);
});

test('the part a binding reads: the one its storeItemID names, else the first where the XPath finds a node (as Word)', () => {
  const body = [
    bound({ xpath: '/other/z' }),
    bound({ xpath: '/root/a' }),
    bound({ xpath: '/root/c' }),
    bound({ xpath: '/root/a', store: ID3 }),
    bound({ xpath: '/root/a', store: '{99999999-2222-3333-4444-555555555555}' }),
    bound({ xpath: '/root/c', store: ID1 }),
    bound({ xpath: '/root/a', store: ID3.slice(1, -1).toLowerCase() }),
  ].map(P).join('');
  const parts = [{ xml: '<root><a>Alpha</a></root>', id: ID1 }, { xml: '<other><z>Zed</z></other>' }, { xml: '<root><a>FromPart3</a><c>Gamma</c></root>', id: ID3 }];
  const s = new DocxSession(bindingDocx({ body, parts }));
  assert.deepEqual(texts(s.doc), ['Zed', 'Alpha', 'Gamma', 'FromPart3', 'Alpha', '.', 'FromPart3']);
});

test('namespaced parts are read through w:prefixMappings', () => {
  const body = [
    bound({ xpath: '/ns0:data[1]/ns0:v[1]', prefixes: "xmlns:ns0='urn:probe'" }),
    bound({ xpath: '/q:data/q:v', prefixes: 'xmlns:q=&quot;urn:probe&quot;' }),
    bound({ xpath: '/data/v' }),
  ].map(P).join('');
  const s = new DocxSession(bindingDocx({ body, parts: [{ xml: '<ns0:data xmlns:ns0="urn:probe"><ns0:v>Namespaced</ns0:v></ns0:data>' }] }));
  assert.deepEqual(texts(s.doc), ['Namespaced', 'Namespaced', '.']);
});

test('a bound date shows in the control\'s format, a UTC one in local time as Word shows it', () => {
  const d = parseXmlDate('2012-10-03T00:00:00Z');
  assert.equal(d.getTime(), Date.UTC(2012, 9, 3));
  assert.equal(parseXmlDate('2012-10-03').getDate(), 3);
  assert.equal(parseXmlDate('2012-13-03'), null);
  const kind = '<w:date><w:dateFormat w:val="yyyy-MM-dd"/><w:lid w:val="en-US"/><w:storeMappedDataAs w:val="dateTime"/></w:date>';
  const s = new DocxSession(bindingDocx({ body: P(bound({ xpath: '/r/d', kind })), parts: [{ xml: '<r><d>2012-10-03T12:00:00</d></r>' }] }));
  assert.deepEqual(texts(s.doc), ['2012-10-03']);
  assert.match(listFields(s.doc)[0].pr, /w:fullDate="2012-10-03T00:00:00Z"/);
});

test('bound controls around a table cell and in a header show their values, and the header is saved so', () => {
  const s = new DocxSession(bindingDocx(FIXTURE));
  assert.match(s.doc.textContent, /Contract: NNC-0042Order no: first orderOrder no again: first orderReference: {2}TitleBound title/);
  const hdr = s.story('rIdH1');
  assert.equal(hdr.textContent, 'Header order: first order');
  const out = unzipSync(s.save());
  const h = strFromU8(out['word/header1.xml']);
  assert.match(h, /first order/);
  assert.ok(!/showingPlcHdr/.test(h));
  const body = strFromU8(out['word/document.xml']);
  assert.match(body, /<w:dataBinding w:xpath="\/root\/title"\/><\/w:sdtPr><w:sdtContent><w:tc><w:tcPr>.*?<\/w:tcPr><w:p>.*?Bound title.*?<\/w:p><\/w:tc>/);
});

test('typing into a bound control updates its node, the controls bound to it (body and header), and the saved part', () => {
  const bytes = bindingDocx(FIXTURE);
  const s = new DocxSession(bytes);
  s.story('rIdH1'); // (the header read, as the pages read it)
  let state = EditorState.create({ doc: s.doc, plugins: [...formsPlugin(), bindingsPlugin(s)] });
  const order = listFields(state.doc).find((f) => f.tag === 'Order');
  state = state.apply(state.tr.insertText('second order', order.from, order.to));
  assert.deepEqual(Object.fromEntries(listFields(state.doc).map((f) => [f.tag, f.text])), { Number: 'NNC-0042', Order: 'second order', OrderAgain: 'second order', Ref: ' ' });
  assert.equal(s.bindings.value({ xpath: '/root/order', store: '', prefixes: '' }), 'second order');
  assert.equal(s.story('rIdH1').textContent, 'Header order: second order');
  assert.ok(s._bindRestory, 'the pages are told to draw the headers again');
  // A cell's control, typed into: its node follows.
  let cell = null;
  state.doc.descendants((n, pos) => { if (n.attrs?.sdt && n.type.spec.tableRole === 'cell') cell = { n, pos }; });
  const from = cell.pos + 2, to = cell.pos + cell.n.nodeSize - 2;
  state = state.apply(state.tr.insertText('New title', from, to));
  assert.equal(s.bindings.value({ xpath: '/root/title', store: '', prefixes: '' }), 'New title');
  const out = unzipSync(s.save(state.doc));
  const item = strFromU8(out['customXml/item1.xml']);
  assert.equal(item, '<?xml version="1.0" encoding="utf-8"?><root><number>NNC-0042</number><order>second order</order><ref> </ref><title>New title</title></root>');
  const orig = unzipSync(bytes);
  assert.deepEqual(out['customXml/itemProps1.xml'], orig['customXml/itemProps1.xml'], 'the part\'s properties are left as they were');
  assert.deepEqual(out['customXml/_rels/item1.xml.rels'], orig['customXml/_rels/item1.xml.rels']);
  const doc = strFromU8(out['word/document.xml']);
  assert.equal(doc.match(/second order/g).length, 2);
  assert.match(strFromU8(out['word/header1.xml']), /second order/);
  // Read back: the same values, and no placeholder left on them.
  const again = new DocxSession(s.save(state.doc));
  assert.match(again.doc.textContent, /Order no: second orderOrder no again: second order.*New title/);
});

test('an edit that is undone puts the node back; a part nothing edits is saved unchanged', async () => {
  const { history, undo } = await import('prosemirror-history');
  const bytes = bindingDocx({ ...FIXTURE, parts: [...FIXTURE.parts, { xml: '<?xml version="1.0"?>\n<untouched a="1">  <x/> </untouched>' }] });
  const s = new DocxSession(bytes);
  let state = EditorState.create({ doc: s.doc, plugins: [history(), bindingsPlugin(s)] });
  const num = listFields(state.doc).find((f) => f.tag === 'Number');
  state = state.apply(state.tr.insertText('X-1', num.from, num.to));
  assert.equal(s.bindings.value('<w:dataBinding w:xpath="/root/number"/>'), 'X-1');
  undo(state, (tr) => { state = state.apply(tr); });
  assert.equal(listFields(state.doc).find((f) => f.tag === 'Number').text, 'NNC-0042');
  assert.equal(s.bindings.value('<w:dataBinding w:xpath="/root/number"/>'), 'NNC-0042');
  const out = unzipSync(s.save(state.doc));
  assert.deepEqual(out['customXml/item2.xml'], unzipSync(bytes)['customXml/item2.xml']);
});

test('crafted bindings and parts stay quick and harmless', () => {
  const deep = '<root>' + '<a>'.repeat(2000) + 'v' + '</a>'.repeat(2000) + '</root>';
  const many = '<root>' + '<b>x</b>'.repeat(20000) + '</root>';
  const xp = '/root' + '/a'.repeat(63);
  const body = [
    bound({ xpath: xp }),
    bound({ xpath: '/root/b[99999]' }),
    bound({ xpath: '/' + 'x'.repeat(5000) }),
    bound({ xpath: '/__proto__/constructor' }),
    bound({ xpath: '/a:root', prefixes: 'xmlns:a=\'' + 'b'.repeat(10000) }),
  ].map(P).join('') + P(bound({ xpath: '/root/b[20000]', content: PH }));
  const t = Date.now();
  const s = new DocxSession(bindingDocx({ body, parts: [{ xml: deep }, { xml: many }, { xml: '<__proto__><constructor>p</constructor></__proto__>' }] }));
  assert.ok(Date.now() - t < 3000, `${Date.now() - t}ms`);
  assert.deepEqual(texts(s.doc), ['v', '.', '.', 'p', '.', 'x']);
  assert.equal({}.constructor, Object);
});

test('filling bound fields without an editor (fillFields) updates their nodes too', () => {
  const { bytes, missing } = fillFields(bindingDocx(FIXTURE), { Order: 'filled order' });
  assert.deepEqual(missing, []);
  const out = unzipSync(bytes);
  assert.match(strFromU8(out['customXml/item1.xml']), /<order>filled order<\/order>/);
  assert.match(strFromU8(out['word/header1.xml']), /filled order/);
  const v = readFieldValues(bytes);
  assert.equal(v.Order, 'filled order');
  assert.equal(v.OrderAgain, 'filled order');
});
