import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { EditorState, TextSelection } from 'prosemirror-state';
import { toggleMark } from 'prosemirror-commands';
import { schema } from '../src/editor/schema.js';
import { csFormatPlugin } from '../src/editor/csformat.js';
import { rPr } from '../src/ooxml/writer.js';
import { readRPr, rPrToCss } from '../src/ooxml/props.js';
import { parseXml, NS } from '../src/ooxml/xml.js';

const rtl = schema.marks.rextra.create({ items: [['rtl', '<w:rtl/>']] });
const make = (marks) => {
  const doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, [schema.text('שלום עולם', marks)])]);
  const s = EditorState.create({ doc, plugins: [csFormatPlugin] });
  return s.apply(s.tr.setSelection(TextSelection.create(s.doc, 1, 5)));
};
const apply = (s, cmd) => { let out = s; cmd(s, (tr) => { out = s.applyTransaction(tr).state; }); return out; };
const runXml = (s, at) => rPr(s.doc.nodeAt(at).marks);

test('bold, italic and a size applied to a right-to-left run also set its complex-script members', () => {
  let s = make([rtl]);
  s = apply(s, toggleMark(schema.marks.bold));
  assert.match(runXml(s, 1), /<w:b\/><w:bCs\/>/);
  assert.doesNotMatch(runXml(s, 6), /bCs/); // outside the selection
  s = apply(s, toggleMark(schema.marks.italic));
  assert.match(runXml(s, 1), /<w:i\/><w:iCs\/>/);
  s = apply(s, (st, d) => { d(st.tr.addMark(1, 5, schema.marks.size.create({ val: 32 }))); return true; });
  assert.match(runXml(s, 1), /<w:sz w:val="32"\/><w:szCs w:val="32"\/>/);
  s = apply(s, toggleMark(schema.marks.bold));
  assert.doesNotMatch(runXml(s, 1), /<w:b\b|bCs/);
});

test('a left-to-right run gets no complex-script members', () => {
  const s = apply(make([]), toggleMark(schema.marks.bold));
  assert.equal(runXml(s, 1), '<w:rPr><w:b/></w:rPr>');
});

test('style run properties carry their complex-script members for right-to-left runs', () => {
  const el = parseXml(`<w:rPr xmlns:w="${NS.w}"><w:rFonts w:ascii="Arial" w:cs="David"/><w:b/><w:bCs w:val="0"/><w:iCs/><w:sz w:val="22"/><w:szCs w:val="28"/></w:rPr>`).documentElement;
  const css = rPrToCss(readRPr(el));
  assert.match(css, /--odx-bcs:normal/);
  assert.match(css, /--odx-ics:italic/);
  assert.match(css, /--odx-szcs:14pt/);
  assert.match(css, /--odx-fcs:"David"/);
  // The run's own class, which editor.css draws with them.
  const dom = schema.marks.rextra.spec.toDOM(rtl);
  assert.match(dom[1].class, /odx-cs/);
});

test('a highlight in a colour Word has no name for is written as run shading', () => {
  const t = schema.text('x', [schema.marks.highlight.create({ val: '#ffeb3b' })]);
  assert.equal(rPr(t.marks), '<w:rPr><w:shd w:val="clear" w:color="auto" w:fill="FFEB3B"/></w:rPr>');
  assert.equal(rPr([schema.marks.highlight.create({ val: 'yellow' })]), '<w:rPr><w:highlight w:val="yellow"/></w:rPr>');
});

test('percentage and solid shading patterns are drawn as the colour they blend to', async () => {
  const { shdFill } = await import('../src/ooxml/props.js');
  const shd = (a) => parseXml(`<w:shd xmlns:w="${NS.w}" ${a}/>`).documentElement;
  assert.equal(shdFill(shd('w:val="pct10" w:color="auto" w:fill="auto"')), '#E6E6E6');
  assert.equal(shdFill(shd('w:val="solid" w:color="FF0000" w:fill="auto"')), '#FF0000');
  assert.equal(shdFill(shd('w:val="pct50" w:color="000000" w:fill="FFFFFF"')), '#808080');
  assert.equal(shdFill(shd('w:val="clear" w:color="auto" w:fill="D9D9D9"')), '#D9D9D9');
  assert.equal(shdFill(shd('w:val="clear" w:color="auto" w:fill="auto"')), null);
});

test('a distributed paragraph spreads its last line too', async () => {
  const { pPrToCss } = await import('../src/ooxml/props.js');
  assert.match(pPrToCss({ align: 'distribute' }), /text-align-last:justify/);
  assert.match(pPrToCss({ align: 'thaiDistribute' }), /text-align-last:justify/);
  assert.doesNotMatch(pPrToCss({ align: 'both' }), /text-align-last/);
});

test('Hebrew or Arabic typed into a run is made a right-to-left run (w:rtl), as Word does; Latin text is left alone', () => {
  const doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, [schema.text('ab')])]);
  let s = EditorState.create({ doc, plugins: [csFormatPlugin] });
  s = s.apply(s.tr.insertText('שלום', 3));
  s = s.apply(s.tr.insertText('cd', 1));
  assert.match(rPr(s.doc.nodeAt(s.doc.content.size - 5).marks), /<w:rtl\/>/);
  assert.equal(rPr(s.doc.nodeAt(1).marks), '');
  const t = s.doc.firstChild.textContent;
  assert.equal(t, 'cdabשלום');
});

test('Right-to-Left Text Direction sets the paragraph\'s w:bidi, and Left-to-Right takes it away', async () => {
  const { setDirection } = await import('../src/editor/commands.js');
  const { pPrXml } = await import('../src/ooxml/writer.js');
  const doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create({ align: 'left' }, [schema.text('abc')])]);
  let s = EditorState.create({ doc, selection: TextSelection.create(doc, 2) });
  setDirection(true)(s, (tr) => { s = s.apply(tr); });
  assert.equal(pPrXml(s.doc.firstChild.attrs), '<w:pPr><w:bidi/><w:jc w:val="left"/></w:pPr>');
  assert.match(schema.nodes.paragraph.spec.toDOM(s.doc.firstChild)[1].style, /direction:rtl/);
  setDirection(false)(s, (tr) => { s = s.apply(tr); });
  assert.equal(pPrXml(s.doc.firstChild.attrs), '<w:pPr><w:jc w:val="left"/></w:pPr>');
});
