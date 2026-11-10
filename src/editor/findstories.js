// Find and Replace through every part of the document, as Word searches it: the main text, then the text boxes, the
// headers and footers, the footnotes and the endnotes. Moving to a match in another part opens that part's editor
// (hfedit.js, notes.js, txbxedit.js) with the match selected; Replace All changes them all.
import { EditorState, TextSelection } from 'prosemirror-state';
import { findAll, replacementFor } from './find.js';
import { editingStory, closeStory, editStory } from './hfedit.js';
import { noteRefs, editingNote, closeNote, editNote } from './notes.js';
import { editingTextBox, closeTextBox, editTextBox } from './txbxedit.js';

/** The document's view for a view that may be a header, note or text box editor over it. */
export const mainOf = (view) => view?.odx?.main || view;

const hasBox = (n) => n.type.name === 'image' && /txbxContent/.test(n.attrs.xml || '');

/** The parts to search in order: [{ key, kind, doc, ... }], the main text first. */
export function stories(main) {
  const { session } = main.odx;
  const out = [{ key: 'main', kind: 'main', doc: main.state.doc }];
  if (main.odx.readOnly || !session) return out;
  main.state.doc.descendants((n, pos) => {
    if (!hasBox(n)) return;
    (session.textBoxes(n.attrs.xml) || []).forEach((b, k) => {
      const doc = session.textBox(n.attrs.xml, k);
      if (doc) out.push({ key: `txbx:${pos}:${k}`, kind: 'textbox', pos, k, doc });
    });
  });
  // Headers and footers in the order the sections name them (each part once).
  const seen = new Set();
  const sects = [];
  main.state.doc.forEach((n) => { const x = (n.attrs.extra || []).find((e) => e[0] === 'sectPr')?.[1]; if (x) sects.push(x); });
  sects.push(main.state.doc.attrs.sectPr || '');
  for (const x of sects) {
    for (const m of x.matchAll(/<w:(header|footer)Reference\b[^>]*>/g)) {
      const rId = /r:id="([^"]+)"/.exec(m[0])?.[1];
      if (!rId || seen.has(rId)) continue;
      seen.add(rId);
      const doc = session.story(rId);
      if (doc) out.push({ key: `hf:${rId}`, kind: m[1], rId, doc });
    }
  }
  const refs = noteRefs(main.state.doc);
  for (const kind of ['footnote', 'endnote']) {
    for (const r of refs) {
      const note = r.kind === kind && session.notes?.[kind]?.get(r.id);
      if (note) out.push({ key: `note:${kind}:${r.id}`, kind: 'note', noteKind: kind, id: r.id, note, doc: note.doc });
    }
  }
  return out;
}

/** The key of the part being edited (see stories()). */
function currentKey(main) {
  const s = editingStory(); if (s?.main === main) return `hf:${s.rId}`;
  const t = editingTextBox(); if (t?.main === main) return `txbx:${t.pos}:${t.k}`;
  const n = editingNote(); if (n?.main === main) return `note:${n.kind}:${n.note.id}`;
  return 'main';
}

function closeAll() { closeStory(); closeNote(); closeTextBox(); }

/** Open part `s` for editing; returns its view, or null when it can't be shown. */
function openPart(main, s) {
  closeAll();
  if (s.kind === 'main') { main.focus(); return main; }
  if (s.kind === 'textbox') return editTextBox(main, s.pos, null, null, s.k);
  if (s.kind === 'note') {
    const el = main.dom.parentElement.querySelector(`.odx-note[data-kind="${s.noteKind}"][data-id="${CSS.escape(s.id)}"]`);
    if (!el) return null;
    el.scrollIntoView({ block: 'center' });
    return editNote(main, s.noteKind, s.id, el);
  }
  // A header or footer: on the first page that shows it.
  const { pages } = main.odx;
  for (let p = 0; p < pages.count; p++) {
    const at = pages.pageAt(pages.pageTop(p) + 1);
    if (at.hf?.[s.kind]?.[at.type] === s.doc) return editStory(main, s.kind, p);
  }
  return null;
}

const select = (v, m) => v.dispatch(v.state.tr.setSelection(TextSelection.create(v.state.doc, m.from, m.to)).scrollIntoView());

/**
 * Select the next (dir 1) or previous (dir -1) match from the selection of `view` (the document or a part's editor),
 * going on into the following parts. Returns the count of matches in all parts.
 */
export function findNextAll(view, query, opts, dir = 1) {
  const main = mainOf(view);
  const here = currentKey(main);
  const cur = here === 'main' ? main : main.odx.hfView || main;
  const list = stories(main);
  let i = list.findIndex((s) => s.key === here);
  if (i < 0) i = 0;
  list[i] = { ...list[i], doc: cur.state.doc };
  const found = list.map((s) => findAll(s.doc, query, opts));
  const total = found.reduce((n, f) => n + f.length, 0);
  if (!total) return 0;
  const { from, to } = cur.state.selection;
  const mine = found[i];
  const m = dir > 0 ? mine.find((x) => x.from >= to) : [...mine].reverse().find((x) => x.to <= from);
  if (m) { select(cur, m); return total; }
  for (let step = 1; step <= list.length; step++) {
    const j = (i + dir * step + list.length * step) % list.length;
    if (!found[j].length) continue;
    const v = openPart(main, list[j]);
    if (!v) continue;
    const ms = findAll(v.state.doc, query, opts);
    if (!ms.length) continue;
    select(v, dir > 0 ? ms[0] : ms[ms.length - 1]);
    return total;
  }
  return total;
}

/** Replace the selected match (if it is one) where it is, and move to the next. */
export function replaceOneAll(view, query, replacement, opts) {
  const { from, to } = view.state.selection;
  const hit = findAll(view.state.doc, query, opts).find((x) => x.from === from && x.to === to);
  if (hit) view.dispatch(view.state.tr.insertText(replacementFor(hit, replacement), from, to));
  return findNextAll(view, query, opts, 1);
}

/** `doc` with every match replaced, and how many: { doc, n }, or null when it has none. */
function replacedIn(doc, query, replacement, opts) {
  const all = findAll(doc, query, opts);
  if (!all.length) return null;
  const tr = EditorState.create({ doc }).tr;
  for (const m of [...all].reverse()) tr.insertText(replacementFor(m, replacement), m.from, m.to);
  return { doc: tr.doc, n: all.length };
}

/** Replace every match in every part; returns how many were replaced. */
export function replaceAllStories(view, query, replacement, opts) {
  const main = mainOf(view);
  closeAll();
  const { session } = main.odx;
  const list = stories(main);
  let total = 0;
  const tr = main.state.tr;
  for (const m of [...findAll(main.state.doc, query, opts)].reverse()) { tr.insertText(replacementFor(m, replacement), m.from, m.to); total++; }
  // Text boxes: each drawing's run XML takes its boxes' new text, one undoable step with the main text.
  const boxes = new Map();
  for (const s of list) {
    if (s.kind !== 'textbox') continue;
    const r = replacedIn(s.doc, query, replacement, opts);
    if (!r) continue;
    total += r.n;
    if (!boxes.has(s.pos)) boxes.set(s.pos, []);
    boxes.get(s.pos).push([s.k, r.doc]);
  }
  for (const [pos, changed] of boxes) {
    const node = main.state.doc.nodeAt(pos);
    let { xml, src } = node.attrs;
    for (const [k, doc] of changed) { const out = session.setTextBox(xml, doc, k); xml = out.xml; src = out.src || src; }
    tr.setNodeMarkup(tr.mapping.map(pos), null, { ...node.attrs, xml, src });
  }
  let parts = false;
  for (const s of list) {
    if (s.kind !== 'header' && s.kind !== 'footer' && s.kind !== 'note') continue;
    const r = replacedIn(s.doc, query, replacement, opts);
    if (!r) continue;
    total += r.n;
    parts = true;
    if (s.kind === 'note') {
      s.note.doc = r.doc;
      s.note.dirty = true;
      main.odx.stories?.note(s.noteKind, s.id, r.doc); // (shared: src/collab.js)
    } else {
      session.setStory(s.rId, r.doc);
      main.odx.stories?.story(s.rId, r.doc);
    }
  }
  if (tr.docChanged) main.dispatch(tr);
  if (parts) { main.odx.pages.restory(); main.odx.notify?.(main); }
  return total;
}
