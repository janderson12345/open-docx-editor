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
