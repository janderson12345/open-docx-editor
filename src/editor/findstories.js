// Find and Replace through every part of the document, as Word searches it: the main text, then the text boxes, the
// headers and footers, the footnotes and the endnotes, then the comments. Moving to a match in another part opens that
// part's editor (hfedit.js, notes.js, txbxedit.js, or the comment's box in the margin) with the match selected;
// Replace All changes them all.
import { EditorState, TextSelection } from 'prosemirror-state';
import { findAll, replacementFor } from './find.js';
import { editingStory, closeStory, editStory } from './hfedit.js';
import { noteRefs, editingNote, closeNote, editNote } from './notes.js';
import { editingTextBox, closeTextBox, editTextBox } from './txbxedit.js';
import { schema } from './schema.js';
import { refresh, setText } from './comments.js';
import { commentsAllowed } from './protect.js';

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
  // The comments (and replies), each as text: a paragraph a line.
  if (main.odx.features?.comments !== false) {
    for (const c of session.comments || []) if (c.text && !c.rich) out.push({ key: `cmt:${c.id}`, kind: 'comment', id: c.id, doc: commentDoc(c.text) });
  }
  return out;
}

const commentDoc = (text) => schema.nodes.doc.create(null, text.split('\n').map((l) => schema.nodes.paragraph.create(null, l ? schema.text(l) : null)));
/** A position in a comment's doc as an offset in its text. */
const offset = (doc, pos) => doc.textBetween(0, pos, '\n').length;
/** The comment open for editing in the margin, if any: its box and the comment's id. (Its selection stays while the
 * find bar has the focus.) */
function commentBox(main) {
  const id = main.odx.session?.editingComment;
  if (id == null) return null;
  const ta = main.dom.parentElement?.querySelector(`textarea[data-odx-box="e${CSS.escape(String(id))}"]`);
  return ta ? { ta, id: String(id) } : null;
}

/** Open comment `id` for editing in the margin and select [from, to] of its doc there. */
function selectInComment(main, id, doc, m) {
  closeAll();
  const { session } = main.odx;
  session.activeComment = id;
  session.editingComment = id;
  session.findComment = id; // (closed with the find bar: findbar.js)
  // (The box selects it as it opens: balloons.js.)
  session.commentSelect = { id: String(id), from: offset(doc, m.from), to: offset(doc, m.to) };
  refresh(main);
}

/** The key of the part being edited (see stories()). */
function currentKey(main) {
  const c = commentBox(main); if (c) return `cmt:${c.id}`;
  const s = editingStory(); if (s?.main === main) return `hf:${s.rId}`;
  const t = editingTextBox(); if (t?.main === main) return `txbx:${t.pos}:${t.k}`;
  const n = editingNote(); if (n?.main === main) return `note:${n.kind}:${n.note.id}`;
  return 'main';
}

function closeAll() { closeStory(); closeNote(); closeTextBox(); }

/** Open part `s` for editing; returns its view, or null when it can't be shown. */
function openPart(main, s) {
  closeAll();
  // (Out of a comment's box, as the search moves on.)
  if (main.odx.session?.editingComment != null) { main.odx.session.editingComment = null; refresh(main); }
  if (s.kind === 'main') { main.focus(); return main; }
  if (s.kind === 'comment') return null; // (opened by selectInComment)
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

// Where the last Find went: { index (from 1), total, part } (for the find bar to say, as screen readers need to hear it).
let lastHit = null;
export const findPosition = () => lastHit;
const PART = { main: '', textbox: 'in a text box', header: 'in the header', footer: 'in the footer', comment: 'in a comment' };
const partName = (s) => (s.kind === 'note' ? `in ${s.noteKind === 'endnote' ? 'an endnote' : 'a footnote'}` : PART[s.kind] ?? '');
const hit = (found, list, j, k, total) => { lastHit = { index: found.slice(0, j).reduce((n, f) => n + f.length, 0) + k + 1, total, part: partName(list[j]) }; };

/**
 * Select the next (dir 1) or previous (dir -1) match from the selection of `view` (the document or a part's editor),
 * going on into the following parts. Returns the count of matches in all parts.
 */
export function findNextAll(view, query, opts, dir = 1) {
  const main = mainOf(view);
  const here = currentKey(main);
  const box = here.startsWith('cmt:') ? commentBox(main) : null;
  const cur = here === 'main' || box ? main : main.odx.hfView || main;
  const list = stories(main);
  let i = list.findIndex((s) => s.key === here);
  if (i < 0) i = 0;
  if (box) list[i] = { ...list[i], doc: commentDoc(box.ta.value) };
  else list[i] = { ...list[i], doc: cur.state.doc };
  const found = list.map((s) => findAll(s.doc, query, opts));
  const total = found.reduce((n, f) => n + f.length, 0);
  lastHit = null;
  if (!total) return 0;
  const mine = found[i];
  if (box) {
    const d = list[i].doc;
    const m = dir > 0 ? mine.find((x) => offset(d, x.from) >= box.ta.selectionEnd) : [...mine].reverse().find((x) => offset(d, x.to) <= box.ta.selectionStart);
    if (m) { box.ta.setSelectionRange(offset(d, m.from), offset(d, m.to)); hit(found, list, i, mine.indexOf(m), total); return total; }
  } else {
    const { from, to } = cur.state.selection;
    const m = dir > 0 ? mine.find((x) => x.from >= to) : [...mine].reverse().find((x) => x.to <= from);
    if (m) { select(cur, m); hit(found, list, i, mine.indexOf(m), total); return total; }
  }
  for (let step = 1; step <= list.length; step++) {
    const j = (i + dir * step + list.length * step) % list.length;
    if (!found[j].length) continue;
    if (list[j].kind === 'comment') {
      const ms = found[j];
      selectInComment(main, list[j].id, list[j].doc, dir > 0 ? ms[0] : ms[ms.length - 1]);
      hit(found, list, j, dir > 0 ? 0 : ms.length - 1, total);
      return total;
    }
    const v = openPart(main, list[j]);
    if (!v) continue;
    const ms = findAll(v.state.doc, query, opts);
    if (!ms.length) continue;
    select(v, dir > 0 ? ms[0] : ms[ms.length - 1]);
    hit(found, list, j, dir > 0 ? 0 : ms.length - 1, total);
    return total;
  }
  return total;
}

/** Replace the selected match (if it is one) where it is, and move to the next. */
export function replaceOneAll(view, query, replacement, opts) {
  // In a comment's box: the selected match replaced in its text, then on to the next.
  const box = commentBox(mainOf(view));
  if (box) {
    const { ta } = box;
    const d = commentDoc(ta.value);
    const hit = findAll(d, query, opts).find((x) => offset(d, x.from) === ta.selectionStart && offset(d, x.to) === ta.selectionEnd);
    if (hit) { const at = ta.selectionStart; ta.setRangeText(replacementFor(hit, replacement), at, ta.selectionEnd, 'end'); ta.dispatchEvent(new Event('input')); }
    return findNextAll(view, query, opts, 1);
  }
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
  lastHit = null; // (a count, not a place)
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
  // Comments: each one's text with its matches replaced (not where the document can only be read).
  if (commentsAllowed(main.state)) {
    for (const s of list) {
      if (s.kind !== 'comment') continue;
      const r = replacedIn(s.doc, query, replacement, opts);
      if (!r) continue;
      total += r.n;
      setText(main, s.id, r.doc.textBetween(0, r.doc.content.size, '\n'));
    }
  }
  return total;
}
