import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { EditorState, TextSelection } from 'prosemirror-state';
import { joinBackward, joinForward, deleteSelection } from 'prosemirror-commands';
import { DocxSession } from '../src/ooxml/docx.js';
import { blankDocxBytes } from '../src/ooxml/template.js';
import { sectionJoinPlugin, sectionEnd } from '../src/editor/sections.js';

// Columns on selected text, as Word makes it: P1 P2 | P3 P4 in two columns | P5 P6. The first section ends with P2, the
// two-column one with P4 (continuous breaks both).
const ONE = '<w:sectPr><w:type w:val="continuous"/><w:cols w:space="720"/></w:sectPr>';
const TWO = '<w:sectPr><w:type w:val="continuous"/><w:cols w:num="2" w:space="720"/></w:sectPr>';
const para = (t, sect = '') => `<w:p>${sect ? `<w:pPr>${sect}</w:pPr>` : ''}<w:r><w:t>${t}</w:t></w:r></w:p>`;
const BODY = para('P1') + para('P2', ONE) + para('P3') + para('P4', TWO) + para('P5') + para('P6');

function session(body = BODY) {
  const files = unzipSync(blankDocxBytes());
  const s0 = new DocxSession(blankDocxBytes());
  files[s0.pkg.mainPath] = strToU8(strFromU8(files[s0.pkg.mainPath]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
  return new DocxSession(zipSync(files));
}
const stateOf = (s) => EditorState.create({ doc: s.doc, plugins: [sectionJoinPlugin] });
/** Each top-level paragraph as text, with |1 or |2 where it ends a one- or two-column section. */
const shape = (doc) => {
  const out = [];
  doc.forEach((n) => { const x = sectionEnd(n); out.push(n.textContent + (x == null ? '' : /w:num="2"/.test(x) ? '|2' : '|1')); });
  return out.join(' ');
};
const posIn = (doc, text, off) => {
  let at = null;
  doc.descendants((n, p) => { if (at == null && n.isText && n.text === text) at = p + off; });
  return at;
};
const run = (st, cmd, from, to = from) => {
  let s = st.apply(st.tr.setSelection(TextSelection.create(st.doc, from, to)));
  cmd(s, (tr) => { s = s.apply(tr); });
  return s;
};

test('section break deleted: Backspace at the start of the next section joins the text to it, in its setup', () => {
  const s = session();
  let st = stateOf(s);
  assert.equal(shape(st.doc), 'P1 P2|1 P3 P4|2 P5 P6');
  st = run(st, joinBackward, posIn(st.doc, 'P3', 0));
  assert.equal(shape(st.doc), 'P1 P2P3 P4|2 P5 P6', 'the first section went; P2 joins the two-column section');
  // Delete at the end of P4: the two-column section goes, its text joins the last section.
  st = run(stateOf(s), joinForward, posIn(stateOf(s).doc, 'P4', 2));
  assert.equal(shape(st.doc), 'P1 P2|1 P3 P4P5 P6');
});

test('section break deleted with a selection across it: the joined paragraph ends as the last one it takes in did', () => {
  const s = session();
  const st0 = stateOf(s);
  let st = run(st0, deleteSelection, posIn(st0.doc, 'P2', 1), posIn(st0.doc, 'P3', 1));
  assert.equal(shape(st.doc), 'P1 P3 P4|2 P5 P6');
  // Across both breaks: one section left before P6, the last.
  st = run(st0, deleteSelection, posIn(st0.doc, 'P2', 1), posIn(st0.doc, 'P5', 1));
  assert.equal(shape(st.doc), 'P1 P5 P6');
  // Inside the two-column section, across a paragraph and not a break: its end stays where it was.
  st = run(st0, deleteSelection, posIn(st0.doc, 'P3', 1), posIn(st0.doc, 'P4', 1));
  assert.equal(shape(st.doc), 'P1 P2|1 P4|2 P5 P6');
  // Typing over it: the same.
  let s2 = st0.apply(st0.tr.setSelection(TextSelection.create(st0.doc, posIn(st0.doc, 'P2', 1), posIn(st0.doc, 'P3', 1))));
  s2 = s2.apply(s2.tr.insertText('x'));
  assert.equal(shape(s2.doc), 'P1 Px3 P4|2 P5 P6');
});

test('Delete before an empty paragraph that ends a section: the break stays, on the paragraph the caret was in', () => {
  const s = session(para('P1') + `<w:p><w:pPr>${ONE}</w:pPr></w:p>` + para('P3') + para('P4', TWO) + para('P5'));
  const st0 = stateOf(s);
  assert.equal(shape(st0.doc), 'P1 |1 P3 P4|2 P5');
  const st = run(st0, joinForward, posIn(st0.doc, 'P1', 2));
  assert.equal(shape(st.doc), 'P1|1 P3 P4|2 P5');
});

test('a document with a section break deleted saves with one section fewer, each w:sectPr where it belongs', () => {
  const s = session();
  const st0 = stateOf(s);
  const st = run(st0, joinBackward, posIn(st0.doc, 'P3', 0));
  const xml = strFromU8(unzipSync(s.save(st.doc))[s.pkg.mainPath]);
  const sects = xml.match(/<w:sectPr\b/g) || [];
  assert.equal(sects.length, 2, 'the two-column section and the document\'s last');
  const paras = xml.split('</w:p>');
  const k = paras.findIndex((p) => p.replace(/<[^>]+>/g, '') === 'P2P3');
  assert.ok(k > 0 && !/<w:sectPr/.test(paras[k]), 'P2 and P3 joined, ending no section');
  assert.match(paras[k + 1], /<w:cols w:num="2".*P4/, 'P4 still ends the two-column section they are in');
});
