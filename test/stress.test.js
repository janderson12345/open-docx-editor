// Stress documents (as odx-test's stress files: thousands of tracked changes, thousands of Find matches) growing
// linearly, not quadratically: each step timed on a document of N paragraphs and one of 4N, the larger at most 10
// times slower (4 when linear, 16 when quadratic, as Accept All was when each change's position went through every
// step made before it). A ratio, so it doesn't depend on the machine's speed; the best of several runs, after one to
// warm up, so a pause on a busy machine doesn't count. And the faster ways the long documents take give what the
// plain ones do: Accept and Reject All a run of blocks at a time (track.js), the text spans a block at a time.
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { EditorState } from 'prosemirror-state';
import { Fragment } from 'prosemirror-model';
import { DecorationSet } from 'prosemirror-view';
import { DocxSession } from '../src/ooxml/docx.js';
import { blankDocxBytes } from '../src/ooxml/template.js';
import { resolveAll, resolveChanges, findChange } from '../src/editor/track.js';
import { findAll } from '../src/editor/find.js';
import { textSpansPlugin, spansSet, spanDecos } from '../src/editor/textspans.js';

const N = 600;
const LIMIT = 10;
const RUNS = 5;
const rev = (k) => `w:id="${k}" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"`;

/** A .docx of the body XML. */
function docx(body) {
  const files = unzipSync(blankDocxBytes());
  files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
  return zipSync(files);
}

/** `n` paragraphs, each with a tracked insertion, a tracked deletion and "apple" five times. */
function stressDocx(n) {
  let body = '';
  for (let k = 0; k < n; k++) {
    body += `<w:p><w:r><w:t xml:space="preserve">apple river ${k} apple stone </w:t></w:r><w:ins ${rev(2 * k + 1)}><w:r><w:t xml:space="preserve">added apple ${k} </w:t></w:r></w:ins>`
      + `<w:del ${rev(2 * k + 2)}><w:r><w:delText xml:space="preserve">removed apple ${k} </w:delText></w:r></w:del><w:r><w:t xml:space="preserve">garden apple.</w:t></w:r></w:p>`;
  }
  return docx(body);
}

/** The best time of `fn` over RUNS runs (ms), and its last result. */
function best(fn) {
  let t = Infinity;
  let out;
  for (let i = 0; i < RUNS; i++) {
    const s = performance.now();
    out = fn();
    t = Math.min(t, performance.now() - s);
  }
  return { t, out };
}

/** Accept All (or Reject All), its transaction applied with the text spans' plugin, as in the editor: the new doc. */
function resolveWhole(state, accept) {
  let next = null;
  resolveAll(accept)(state, (tr) => { next = state.apply(tr); });
  return next.doc;
}

const steps = {
  open: (s) => new DocxSession(s.bytes).doc,
  // (The editor's state as it opens: every text node's span made, textspans.js.)
  'text spans': (s) => EditorState.create({ doc: s.doc, plugins: [textSpansPlugin] }),
  find: (s) => findAll(s.doc, 'apple').length,
  'accept all': (s) => resolveWhole(s.state, true),
  'reject all': (s) => resolveWhole(s.state, false),
  save: (s) => s.session.save(s.doc).length,
};

test('stress documents: open, text spans, Find, Accept All, Reject All and save grow linearly with the changes', (t) => {
  const sizes = [N, 4 * N].map((n) => {
    const bytes = stressDocx(n);
    const session = new DocxSession(bytes);
    return { n, bytes, session, doc: session.doc, state: EditorState.create({ doc: session.doc, plugins: [textSpansPlugin] }) };
  });
  // (Each step once on the smaller document first: the code compiled before anything is timed.)
  for (const fn of Object.values(steps)) fn(sizes[0]);
  const report = {};
  for (const [name, fn] of Object.entries(steps)) {
    const [small, large] = sizes.map((s) => best(() => fn(s)));
    report[name] = `${small.t.toFixed(1)}ms, ${large.t.toFixed(1)}ms`;
    // (Below a few milliseconds the timer's own steps count: a step that fast isn't the trouble.)
    if (large.t > 5) assert.ok(large.t / Math.max(small.t, 0.5) < LIMIT, `${name}: ${N} paragraphs ${small.t.toFixed(1)}ms, ${4 * N} ${large.t.toFixed(1)}ms (${(large.t / small.t).toFixed(1)}x)`);
    if (name === 'find') assert.equal(large.out, 4 * N * 5);
    if (name === 'accept all' || name === 'reject all') {
      assert.equal(findChange(large.out, 0, 1), null, `${name} leaves no change`);
      assert.match(large.out.firstChild.textContent, name === 'accept all' ? /added apple 0 garden/ : /removed apple 0 garden/);
    }
  }
  t.diagnostic(JSON.stringify(report));
});

// Every kind of change Accept and Reject All take, over more than 200 blocks (so they go a run of blocks at a time):
// insertions and deletions, a paragraph mark deleted (joining the next) or inserted, a paragraph deleted whole,
// paragraph and run formatting changes, table rows inserted and deleted.
function mixedDocx(cycles) {
  let body = '';
  let id = 1;
  const r = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
  const ins = (t) => `<w:ins ${rev(id++)}>${r(t)}</w:ins>`;
  const del = (t) => `<w:del ${rev(id++)}><w:r><w:delText xml:space="preserve">${t}</w:delText></w:r></w:del>`;
  const mark = (kind) => `<w:pPr><w:rPr><w:${kind} ${rev(id++)}/></w:rPr></w:pPr>`;
  for (let k = 0; k < cycles; k++) {
    body += `<w:p>${r(`plain ${k} `)}${ins('added ')}${del('removed ')}</w:p>`;
    body += `<w:p>${mark('del')}${r(`joins ${k} `)}</w:p>`;
    body += `<w:p>${mark('del')}${del(`gone ${k}`)}</w:p>`;
    body += `<w:p>${mark('ins')}${r(`new mark ${k}`)}</w:p>`;
    body += `<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange ${rev(id++)}><w:pPr/></w:pPrChange></w:pPr><w:r><w:rPr><w:b/><w:rPrChange ${rev(id++)}><w:rPr/></w:rPrChange></w:rPr><w:t>bold ${k}</w:t></w:r></w:p>`;
    const row = (trPr, t) => `<w:tr>${trPr}<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p>${r(t)}</w:p></w:tc></w:tr>`;
    body += `<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>${row('', `row ${k}`)}${row(`<w:trPr><w:ins ${rev(id++)}/></w:trPr>`, 'row in')}${row(`<w:trPr><w:del ${rev(id++)}/></w:trPr>`, 'row out')}</w:tbl>`;
    body += `<w:p>${r(`after ${k}`)}</w:p>`;
  }
  return docx(body);
}

test('Accept and Reject All a run of blocks at a time give what the whole document at once does', () => {
  const docs = { mixed: new DocxSession(mixedDocx(60)).doc };
  // The fixtures with tracked changes, their content repeated past 200 blocks.
  for (const f of fs.readdirSync(new URL('./fixtures/', import.meta.url)).filter((n) => n.endsWith('.docx'))) {
    const doc = new DocxSession(new Uint8Array(fs.readFileSync(new URL(`./fixtures/${f}`, import.meta.url)))).doc;
    if (!findChange(doc, 0, 1)) continue;
    const kids = [];
    while (kids.length <= 200) doc.forEach((b) => kids.push(b));
    docs[f] = doc.type.create(doc.attrs, Fragment.fromArray(kids));
  }
  assert.ok(Object.keys(docs).length > 2, Object.keys(docs).join(', '));
  for (const [name, doc] of Object.entries(docs)) {
    for (const accept of [true, false]) {
      const whole = resolveChanges(EditorState.create({ doc }).tr, 0, doc.content.size, accept).doc;
      let byRuns = null;
      resolveAll(accept)(EditorState.create({ doc }), (tr) => { byRuns = tr; });
      assert.ok(byRuns.steps.length === 1 || name !== 'mixed', 'the mixed document goes a run at a time, in one step');
      assert.ok(byRuns.doc.eq(whole), `${name}: ${accept ? 'Accept' : 'Reject'} All`);
    }
  }
});

test('the text spans made a block at a time are the set DecorationSet.create makes', () => {
  const docs = [new DocxSession(mixedDocx(10)).doc, new DocxSession(stressDocx(50)).doc];
  for (const f of ['rich.docx', 'original.docx', 'reflection-effects.docx'].filter((n) => fs.existsSync(new URL(`./fixtures/${n}`, import.meta.url)))) {
    docs.push(new DocxSession(new Uint8Array(fs.readFileSync(new URL(`./fixtures/${f}`, import.meta.url)))).doc);
  }
  const flat = (set) => set.find().map((d) => [d.from, d.to, d.spec.key ?? '', JSON.stringify(d.type.attrs ?? null)]).sort((a, b) => a[0] - b[0] || a[1] - b[1] || (a[3] < b[3] ? -1 : 1));
  for (const doc of docs) {
    const fast = spansSet(doc);
    const plain = DecorationSet.create(doc, spanDecos(doc));
    assert.ok(fast.find().length > 0);
    assert.deepEqual(flat(fast), flat(plain));
  }
});
