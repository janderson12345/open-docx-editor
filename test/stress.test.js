// Stress documents (as odx-test's stress files: thousands of tracked changes, thousands of Find matches) growing
// linearly, not quadratically: each step timed on a document of N paragraphs and one of 4N, the larger at most 10
// times slower (4 when linear, 16 when quadratic, as Accept All was when each change's position went through every
// step made before it). A ratio, so it doesn't depend on the machine's speed; the best of several runs, after one to
// warm up, so a pause on a busy machine doesn't count. The same for the decorations (comments, formatting marks...) as a
// document opens and as a key is typed. And the faster ways the long documents take give what the plain ones do:
// Accept and Reject All a run of blocks at a time (track.js), the text spans a block at a time, the other decorations'
// sets from sorted decorations (decoset.js), the soft hyphens mapped through an edit.
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { EditorState } from 'prosemirror-state';
import { Fragment } from 'prosemirror-model';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { DocxSession } from '../src/ooxml/docx.js';
import { blankDocxBytes } from '../src/ooxml/template.js';
import { resolveAll, resolveChanges, findChange } from '../src/editor/track.js';
import { findAll } from '../src/editor/find.js';
import { textSpansPlugin, spansSet, spanDecos } from '../src/editor/textspans.js';
import { decoSet } from '../src/editor/decoset.js';
import { commentsPlugin } from '../src/editor/comments.js';
import { marksPlugin } from '../src/editor/marks.js';
import { mailMergePlugin } from '../src/editor/mailmerge.js';
import { joinsPlugin } from '../src/editor/joins.js';
import { hyphenatePlugin, loadHyphenation, hyphenSet, mapHyphens } from '../src/editor/hyphenate.js';
import { mergeTargets, fieldMarks, mergeCursor } from '../src/merge.js';

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

/** The best time of `fn` over `runs` runs (ms), and its last result. */
function best(fn, runs = RUNS) {
  let t = Infinity;
  let out;
  for (let i = 0; i < runs; i++) {
    const s = performance.now();
    out = fn();
    t = Math.min(t, performance.now() - s);
  }
  return { t, out };
}

/**
 * The smaller and larger documents' best times (`best`), measured again up to three more times while the larger looks
 * LIMIT times as slow or more: a step that is really quadratic is about 16 times as slow every time, while a pause from
 * the other test files running beside this one (npm test runs them in parallel) comes and goes.
 */
function timedPair(small, large, runs = RUNS) {
  let a = best(small, runs), b = best(large, runs);
  for (let i = 0; i < 3 && b.t > 5 && b.t / Math.max(a.t, 0.5) >= LIMIT; i++) {
    const a2 = best(small, runs), b2 = best(large, runs);
    if (b2.t / Math.max(a2.t, 0.5) < b.t / Math.max(a.t, 0.5)) { a = a2; b = b2; }
  }
  return [a, b];
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
    const [small, large] = timedPair(() => fn(sizes[0]), () => fn(sizes[1]));
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

// --- Decorations: comments, formatting marks, merge fields, joined paragraphs, hyphenation ---------------------------

const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const fld = (t) => `<w:r><w:fldChar w:fldCharType="${t}"/></w:r>`;
// (The decorations' documents: as many repeats as the stress documents' paragraphs, the larger made from the smaller.)
const DN = N;
/** Body XML of `n` repeats of what each decoration is made for. */
const decoDocs = {
  comments: (n) => Array.from({ length: n }, (_, k) => `<w:p>${run('before ')}<w:commentRangeStart w:id="${k}"/>${run(`commented ${k}`)}<w:commentRangeEnd w:id="${k}"/><w:r><w:commentReference w:id="${k}"/></w:r>${run(' after')}</w:p>`).join(''),
  marks: (n) => Array.from({ length: n }, (_, k) => `<w:p>${run(`text ${k}`)}</w:p><w:p/>`).join(''),
  'merge fields': (n) => Array.from({ length: n }, (_, k) => `<w:p>${run(`Dear ${k} `)}${fld('begin')}<w:r><w:instrText xml:space="preserve"> MERGEFIELD Name </w:instrText></w:r>${fld('separate')}${run('«Name»')}${fld('end')}${run(',')}</w:p>`).join(''),
  joins: (n) => Array.from({ length: n }, (_, k) => `<w:p><w:pPr><w:rPr><w:del ${rev(2 * k + 1)}/></w:rPr></w:pPr>${run(`joins ${k} `)}</w:p><w:p><w:pPr><w:rPr><w:ins ${rev(2 * k + 2)}/></w:rPr></w:pPr>${run(`into ${k}`)}</w:p><w:p>${run(`plain ${k}`)}</w:p>`).join(''),
  hyphenation: (n) => Array.from({ length: n }, (_, k) => `<w:p>${run(`Characterization of hyphenation considerations ${k} demonstrates extraordinary international communication`)}</w:p>`).join(''),
};
/** Each decoration's plugin for a session (Preview Results on, automatic hyphenation on). */
const decoPlugins = {
  comments: (s) => commentsPlugin(s),
  marks: () => marksPlugin,
  'merge fields': (s) => { s.merge = { records: [{ Name: 'Ada' }], fields: ['Name'], index: 0, preview: true }; return mailMergePlugin(s); },
  joins: () => joinsPlugin,
  hyphenation: (s) => { s.autoHyphenation = true; return hyphenatePlugin(s); },
};
/** The plugin's decorations in `state`. */
const decosOf = (plugin, state) => { const v = plugin.getState(state); return v.set || v; };
/** A state with the plugin, turned on as the editor does (formatting marks and Preview Results by a meta, the others
 *  made with the state). */
function opened(plugin, doc) {
  const st = EditorState.create({ doc, plugins: [plugin] });
  return /^odx-(marks|mailmerge)\$/.test(plugin.spec.key?.key ?? '') ? st.apply(st.tr.setMeta(plugin.spec.key, true)) : st;
}

/**
 * The document's blocks `times` times over (a copy's comment markers given ids of their own): a larger document without
 * reading one, which takes longer than all the timing.
 */
function repeated(doc, times) {
  let top = 0;
  doc.descendants((n) => { const id = /w:id="(\d+)"/.exec(n.attrs?.xml || '')?.[1]; if (id) top = Math.max(top, Number(id) + 1); });
  const copy = (node, k) => {
    if (node.type.name === 'opaque_inline' && /w:id="\d+"/.test(node.attrs.xml || '')) return node.type.create({ ...node.attrs, xml: node.attrs.xml.replace(/w:id="(\d+)"/, (_, id) => `w:id="${Number(id) + k * top}"`) }, null, node.marks);
    if (node.isLeaf || node.isText || !k) return node;
    const kids = [];
    node.forEach((c) => kids.push(copy(c, k)));
    return node.copy(Fragment.fromArray(kids));
  };
  const blocks = [];
  for (let k = 0; k < times; k++) doc.forEach((b) => blocks.push(copy(b, k)));
  return doc.type.create(doc.attrs, Fragment.fromArray(blocks));
}

test('decorations: comments, formatting marks, merge fields, joins and hyphenation grow linearly, as a document opens and as a key is typed', async (t) => {
  await loadHyphenation();
  const report = {};
  for (const [name, body] of Object.entries(decoDocs)) {
    const session = new DocxSession(docx(body(DN)));
    const plugin = decoPlugins[name](session);
    const sizes = [1, 4].map((times) => {
      const doc = repeated(session.doc, times);
      const open = () => opened(plugin, doc);
      const state = open();
      // (A character typed at the end of the first paragraph.)
      const typed = () => state.apply(state.tr.insertText('x', state.doc.firstChild.nodeSize - 1));
      return { open, typed };
    });
    for (const s of sizes) { s.open(); s.typed(); }
    for (const step of ['open', 'typed']) {
      // (More runs: a millisecond or two each, where one pause would count.)
      const [small, large] = timedPair(sizes[0][step], sizes[1][step], 15);
      report[`${name} ${step}`] = `${small.t.toFixed(1)}ms, ${large.t.toFixed(1)}ms`;
      assert.ok(decosOf(plugin, large.out).find().length >= 4 * DN, `${name}: decorations made`);
      if (large.t > 5) assert.ok(large.t / Math.max(small.t, 0.5) < LIMIT, `${name} ${step}: ${DN} ${small.t.toFixed(1)}ms, ${4 * DN} ${large.t.toFixed(1)}ms (${(large.t / small.t).toFixed(1)}x)`);
    }
  }
  t.diagnostic(JSON.stringify(report));
});

/** A set's tree as it is: each level's own decorations (as their index in `decos`) and its children's spans. */
function tree(set, decos) {
  const index = new Map(decos.map((d, i) => [d.type, i]));
  const walk = (s) => ({ local: s.local.map((d) => [d.from, d.to, index.get(d.type)]), children: s.children.map((c, i) => (i % 3 === 2 ? walk(c) : c)) });
  return walk(set);
}
const fixtureSession = (f) => new DocxSession(new Uint8Array(fs.readFileSync(new URL(`./fixtures/${f}`, import.meta.url))));

test('the decorations\' sets made from sorted decorations are the sets DecorationSet.create makes', async () => {
  await loadHyphenation();
  const cases = [];
  // Each plugin's own decorations, on its generated document and on the fixtures it is for.
  const fixtures = { comments: ['find-all.docx', 'rich.docx'], marks: ['rich.docx', 'lists-more.docx'], 'merge fields': ['mailmerge.docx', 'mergeif.docx', 'mergenext.docx'], joins: ['original.docx', 'rich.docx'], hyphenation: ['hyphenation.docx', 'hyphenation-zone.docx', 'other-app.docx'] };
  for (const [name, body] of Object.entries(decoDocs)) {
    for (const session of [new DocxSession(docx(body(40))), ...fixtures[name].map(fixtureSession)]) {
      const plugin = decoPlugins[name](session);
      if (name === 'merge fields') session.merge.records = [Object.fromEntries(['Name', 'FirstName', 'Last Name', 'City', 'Group', 'Amount'].map((f) => [f, `${f} value`]))];
      cases.push([name, session.doc, decosOf(plugin, opened(plugin, session.doc)).find()]);
    }
  }
  // And made-up ones over a document of tables: widgets anywhere, inline decorations within blocks and across them (and
  // empty ones, dropped), node decorations on nodes and off them, several at the same place.
  const doc = new DocxSession(mixedDocx(8)).doc;
  let seed = 7;
  const rand = (n) => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed % n; };
  const nodes = [];
  doc.descendants((node, pos) => { nodes.push([pos, pos + node.nodeSize]); });
  const made = [];
  for (let k = 0; k < 600; k++) {
    const a = rand(doc.content.size);
    const kind = rand(5);
    const [f, e] = nodes[rand(nodes.length)];
    if (kind === 0) made.push(Decoration.widget(a, () => null, { k }));
    else if (kind === 1) made.push(Decoration.inline(a, Math.min(doc.content.size, a + rand(40)), { class: `i${k}` }));
    else if (kind === 2) made.push(Decoration.node(f, e, { class: `n${k}` }));
    else if (kind === 3) made.push(Decoration.node(f, Math.max(f + 1, e + rand(3) - 1), { class: `o${k}` }));
    else if (a > 0 && a < doc.content.size - 2) made.push(Decoration.node(a, a + 1 + rand(2), { class: `m${k}` }));
  }
  made.push(Decoration.inline(2, 6, { class: 'same' }), Decoration.inline(2, 6, { class: 'same2' }), Decoration.widget(0, () => null, { at: 0 }), Decoration.widget(doc.content.size, () => null, { at: 'end' }));
  cases.push(['made up', doc, made]);
  for (const name of [...Object.keys(decoDocs), 'made up']) assert.ok(cases.some(([n, , decos]) => n === name && decos.length), name);
  for (const [name, d, decos] of cases) {
    // (create takes the decorations out of the array it is given: a copy.)
    assert.deepEqual(tree(decoSet(d, decos), decos), tree(DecorationSet.create(d, decos.slice()), decos), name);
  }
});

test('merge field previews: a field\'s paragraph and formatting found from its position give what resolving it does', () => {
  for (const f of ['mailmerge.docx', 'mergeif.docx', 'mergenext.docx', 'mergeskip.docx']) {
    const session = fixtureSession(f);
    const { doc } = session;
    const targets = mergeTargets(doc, mergeCursor([{ Name: 'Ada' }], 0, { preview: true }));
    assert.ok(targets.length, f);
    session.merge = { records: [{ Name: 'Ada' }], fields: ['Name'], index: 0, preview: true };
    const plugin = mailMergePlugin(session);
    const shown = decosOf(plugin, opened(plugin, doc)).find().filter((d) => d.spec.key).map((d) => [d.from, d.spec.marks.map((m) => m.type.name)]);
    const plain = targets.filter(({ field: x }) => doc.resolve(x.from).parent === doc.resolve(x.to).parent).map(({ field: x }) => [x.from, fieldMarks(doc, x).map((m) => m.type.name)]);
    assert.deepEqual(shown, plain, f);
  }
});

test('soft hyphens mapped through edits, the paragraphs touched made again, are the set made whole', async () => {
  await loadHyphenation();
  const flat = (set) => set.find().map((d) => [d.from, d.to, d.spec.key ?? '', JSON.stringify(d.type.attrs ?? null)]).sort((a, b) => a[0] - b[0] || a[1] - b[1] || (a[2] + a[3] < b[2] + b[3] ? -1 : 1));
  const tables = strFromU8(unzipSync(mixedDocx(3))['word/document.xml']).match(/<w:body>([\s\S]*?)<w:sectPr/)[1];
  const sessions = [new DocxSession(docx(decoDocs.hyphenation(30) + tables)), ...['hyphenation-zone.docx', 'hyphenation-limit.docx', 'hyphenation-free.docx'].map(fixtureSession)];
  for (const session of sessions) {
    for (const own of [false, true]) {
      let state = EditorState.create({ doc: session.doc });
      let set = hyphenSet(state.doc, session, own);
      assert.ok(set.find().length > 0);
      // (A position 6 letters into the first long text node after `from`; the paragraphs, [pos, node].)
      const inText = (from) => { let at = null; state.doc.nodesBetween(from, state.doc.content.size, (n, p) => { if (at == null && n.isText && n.text.length > 12) at = Math.max(p, from) + 6; return at == null; }); return at; };
      const para = (i) => { const ps = []; state.doc.descendants((n, p) => { if (n.type.name !== 'paragraph') return true; ps.push([p, n]); return false; }); return ps[1 + (i % (ps.length - 1))]; };
      const bold = state.schema.marks.strong || state.schema.marks.bold || Object.values(state.schema.marks)[0];
      const edits = [
        (tr) => tr.insertText('hyphenationx', inText(0)), // typed into a word
        (tr) => tr.split(inText(Math.floor(tr.doc.content.size / 3))), // Enter in a word
        (tr) => { const [p] = para(4); return tr.delete(p - 1, p + 1); }, // two paragraphs joined
        (tr) => { const [p, n] = para(2); return tr.delete(p, p + n.nodeSize); }, // a paragraph deleted
        (tr) => { const [p, n] = para(1); return tr.setNodeMarkup(p, null, { ...n.attrs, view: { ...n.attrs.view, noHyph: true } }); }, // left unhyphenated
        (tr) => tr.addMark(5, Math.min(tr.doc.content.size, 60), bold.create()), // formatted
        (tr) => tr.insert(para(5)[0], [para(1)[1], para(2)[1]]), // paragraphs pasted
        (tr) => { const at = inText(0); return tr.insertText('extraordinarily', at).insertText('characterization', at + 30); }, // two steps
        (tr) => tr.setDocAttribute('odxTest', 1), // a step with no range
      ];
      for (const [i, edit] of edits.entries()) {
        const tr = state.tr;
        edit(tr);
        assert.ok(tr.docChanged, `edit ${i}`);
        set = mapHyphens(tr, set, session, own);
        state = state.apply(tr);
        assert.deepEqual(flat(set), flat(hyphenSet(state.doc, session, own)), `edit ${i}, ${own ? 'the browser\'s hyphens' : 'soft hyphens'}`);
      }
    }
  }
});
