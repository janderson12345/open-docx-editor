// Tracked changes (Word's "Track Changes"), stored the way Word stores them:
// - inserted text carries an `ins` mark, deleted text a `del` mark (it stays in the
//   document, struck through, until accepted);
// - formatting changes keep the old run properties in a w:rPrChange (in the `rextra`
//   mark) and paragraph changes the old w:pPr in a w:pPrChange (paragraph `extra`);
// - an inserted or deleted paragraph mark is a w:ins / w:del in the paragraph mark's
//   w:rPr (paragraph `extra`);
// - an inserted or deleted table row is a w:ins / w:del in its w:trPr (row `trPr`);
// - moved text is a deletion and an insertion with `move` set (Word's w:moveFrom / w:moveTo).
// While tracking is on, `trackTransaction` rewrites each edit into that form.
import { announce } from './live.js';
import { Plugin, PluginKey, TextSelection } from 'prosemirror-state';
import { Mapping, ReplaceStep, ReplaceAroundStep, AddMarkStep, RemoveMarkStep } from 'prosemirror-transform';
import { schema } from './schema.js';
import { rPr as rPrXml, pPrXml } from '../ooxml/writer.js';
import { marksFromRPr, paragraphAttrsFromPPr } from '../ooxml/reader.js';
import { esc } from '../ooxml/xml.js';
import { now } from './util.js';
import { resolvingAllowed } from './protect.js';

export const trackKey = new PluginKey('odx-track');

/** Tracking state: { on, author }. Start it from the document's own setting (w:trackRevisions). */
export function trackPlugin({ on = false, author = 'Author' } = {}) {
  return new Plugin({
    key: trackKey,
    state: {
      init: () => ({ on, author }),
      apply: (tr, value) => {
        const m = tr.getMeta(trackKey);
        return m ? { ...value, ...m } : value;
      },
    },
  });
}

export const tracking = (state) => trackKey.getState(state) || { on: false, author: 'Author' };

/** Turn tracking on or off (toggles when `on` is omitted). */
export const setTracking = (on) => (state, dispatch) => {
  if (dispatch) dispatch(state.tr.setMeta(trackKey, { on: on ?? !tracking(state).on }).setMeta('addToHistory', false));
  return true;
};

// --- Reading revisions ------------------------------------------------------

const REV_MARKS = ['ins', 'del'];
const fmtChange = (node) => node.marks.find((m) => m.type === schema.marks.rextra && m.attrs.items.some((it) => it[0] === 'rPrChange')) || null;
const extraItem = (attrs, name) => (attrs.extra || []).find((e) => e[0] === name);
/** The paragraph mark's own revision ('ins' | 'del' | null), from w:rPr in the paragraph properties. */
export const paragraphMarkRev = (attrs) => markRev(attrs);
function markRev(attrs) {
  const r = extraItem(attrs, 'rPr');
  // (A paragraph moved whole: its mark moved too, a deletion where it was and an insertion where it went.)
  return r ? (/<w:(ins|moveTo)\b/.test(r[1]) ? 'ins' : /<w:(del|moveFrom)\b/.test(r[1]) ? 'del' : null) : null;
}
/** A table row's own revision ('ins' | 'del' | null), from its w:trPr. */
export const rowRev = (row) => (row.attrs.trPr ? (/<w:ins\b/.test(row.attrs.trPr) ? 'ins' : /<w:del\b/.test(row.attrs.trPr) ? 'del' : null) : null);
/** `trPr` with its revision replaced by `rev` (XML, or '' for none); null when nothing is left. */
function withRowRev(trPr, rev) {
  let inner = trPr ? trPr.replace(/^<w:trPr\b[^>]*>|<\/w:trPr>$|^<w:trPr\b[^>]*\/>$/g, '') : '';
  inner = inner.replace(/<w:(ins|del)\b[^>]*\/>|<w:(ins|del)\b[^>]*>[\s\S]*?<\/w:\2>/g, '');
  // (Before a w:trPrChange, the last of w:trPr's children.)
  const at = inner.search(/<w:trPrChange\b/);
  inner = at < 0 ? inner + (rev || '') : inner.slice(0, at) + (rev || '') + inner.slice(at);
  return inner ? `<w:trPr>${inner}</w:trPr>` : null;
}
/** `trPr` with its w:gridAfter set to `n` grid columns (the row's empty columns at its end; w:wAfter dropped with it). */
function withGridAfter(trPr, n) {
  let inner = trPr ? trPr.replace(/^<w:trPr\b[^>]*>|<\/w:trPr>$|^<w:trPr\b[^>]*\/>$/g, '') : '';
  inner = inner.replace(/<w:gridAfter\b[^>]*\/>|<w:wAfter\b[^>]*\/>/g, '');
  if (n) {
    const el = `<w:gridAfter w:val="${n}"/>`;
    // (After w:cnfStyle, w:divId and w:gridBefore, w:trPr's first children.)
    const m = /^((?:<w:cnfStyle\b[^>]*\/>|<w:divId\b[^>]*\/>|<w:gridBefore\b[^>]*\/>)*)/.exec(inner);
    inner = m[1] + el + inner.slice(m[1].length);
  }
  return inner ? `<w:trPr>${inner}</w:trPr>` : null;
}

/** A table cell's own revision ('ins' | 'del' | null): w:cellIns / w:cellDel in its w:tcPr. */
export const cellRev = (cell) => {
  const items = cell.attrs.tcPr || [];
  return items.some((t) => t[0] === 'cellIns') ? 'ins' : items.some((t) => t[0] === 'cellDel') ? 'del' : null;
};
const isCell = (node) => node.type === schema.nodes.table_cell || node.type === schema.nodes.table_header;
const hasChange = (node) => (isCell(node) && !!cellRev(node)) || (node.isInline
  ? node.marks.some((m) => REV_MARKS.includes(m.type.name)) || !!fmtChange(node)
  : (node.type === schema.nodes.paragraph && (!!extraItem(node.attrs, 'pPrChange') || !!markRev(node.attrs))) || (node.type === schema.nodes.table_row && !!rowRev(node)));

/** Whether a paragraph has any tracked change (for the change bar). */
export function paragraphChanged(node) {
  if (hasChange(node)) return true;
  let found = false;
  node.forEach((c) => { if (!found && hasChange(c)) found = true; });
  return found;
}

// Maps positions through the steps of `tr` from index `start` on.
function since(tr, start) {
  const m = new Mapping();
  for (let k = start; k < tr.mapping.maps.length; k++) m.appendMap(tr.mapping.maps[k]);
  return m;
}

let idSeq = 0;
function nextId(doc) {
  if (!idSeq) {
    // w:id values must be unique among the document's annotations; start above the largest seen.
    let max = 0;
    const see = (s) => { for (const m of String(s).matchAll(/w:id="(\d+)"/g)) max = Math.max(max, Number(m[1])); };
    doc.descendants((n) => {
      for (const m of n.marks) if (m.attrs.id != null && /^\d+$/.test(m.attrs.id)) max = Math.max(max, Number(m.attrs.id));
      if (n.attrs.extra) see(JSON.stringify(n.attrs.extra));
      if (n.attrs.xml) see(n.attrs.xml);
      const items = n.marks.find((m) => m.type === schema.marks.rextra);
      if (items) see(JSON.stringify(items.attrs.items));
    });
    idSeq = Math.max(max, 1000);
  }
  return String(++idSeq);
}

const revXml = (name, attrs) => `<w:${name} w:id="${attrs.id}" w:author="${esc(attrs.author)}" w:date="${attrs.date}"/>`;

// Paragraph mark revisions live in the mark's w:rPr, before its other children.
function withMarkRev(attrs, rev) {
  const extra = attrs.extra.filter((e) => e[0] !== 'rPr');
  const old = extraItem(attrs, 'rPr');
  let inner = old ? old[1].replace(/^<w:rPr\b[^>]*>|<\/w:rPr>$|^<w:rPr\b[^>]*\/>$/g, '') : '';
  inner = inner.replace(/<w:(ins|del|moveFrom|moveTo)\b[^>]*\/>|<w:(ins|del|moveFrom|moveTo)\b[^>]*>[\s\S]*?<\/w:\2>/g, '');
  const xml = (rev || '') + inner;
  if (xml) extra.push(['rPr', `<w:rPr>${xml}</w:rPr>`]);
  return { ...attrs, extra };
}

// --- Recording edits ------------------------------------------------------

/**
 * Rewrite an edit made while tracking is on: deleted content is marked deleted
 * instead of removed (unless it was the same author's own tracked insertion),
 * inserted content is marked inserted, and formatting or paragraph changes keep
 * the old properties. Returns a new transaction on the same starting state.
 */
export function trackTransaction(state, tr, author) {
  const out = state.tr;
  const date = now();
  // (And the true UTC time, as Word 2023 and later write it beside w:date, which is on the author's clock.)
  const dateUtc = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const own = (m) => m && m.attrs.author === author;
  let cursor = null;
  // One id per transaction and kind, reused when continuing the author's own adjacent change.
  const revAttrs = {};
  const attrsFor = (kind) => (revAttrs[kind] ||= { id: nextId(state.doc), author, date, dateUtc });

  // Maps positions in the document before original step i onto `out`.
  const mapFor = (i) => {
    const m = new Mapping();
    for (let k = i - 1; k >= 0; k--) m.appendMap(tr.mapping.maps[k].invert());
    for (const map of out.mapping.maps) m.appendMap(map);
    return m;
  };
  const toOut = (i, pos, assoc) => mapFor(i).map(pos, assoc);

  tr.steps.forEach((step, i) => {
    if (step instanceof ReplaceStep) {
      const a = toOut(i, step.from, 1);
      const b = Math.max(a, toOut(i, step.to, -1));
      let range = { start: a, end: a };
      if (b > a) range = trackDelete(out, a, b, attrsFor, own);
      if (step.slice.size) cursor = trackInsert(out, range.end, step.slice, attrsFor, own);
      else {
        // Backspace leaves the caret before the deleted text, Delete after it.
        const sel = state.selection;
        const back = tr.getMeta('odx-delete') ? tr.getMeta('odx-delete') < 0 : sel.empty && sel.head === step.to && i === 0;
        cursor = back ? range.start : range.end;
      }
    } else if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
      trackFormat(out, toOut(i, step.from, 1), toOut(i, step.to, -1), step, attrsFor);
    } else if (step instanceof ReplaceAroundStep && step.insert === 1 && step.gapFrom === step.from + 1 && step.gapTo === step.to - 1 && step.slice.content.childCount === 1 && out.doc.nodeAt(toOut(i, step.from, 1))?.type === step.slice.content.firstChild.type) {
      trackAttrs(out, toOut(i, step.from, 1), step.slice.content.firstChild, attrsFor);
    } else {
      // Anything else (tables, lists lifted or wrapped, ...) is applied untracked.
      const mapped = step.map(mapFor(i));
      if (mapped) out.maybeStep(mapped);
    }
  });

  if (cursor != null) out.setSelection(TextSelection.near(out.doc.resolve(Math.max(0, Math.min(cursor, out.doc.content.size)))));
  else if (tr.selectionSet) {
    const s = tr.selection;
    const back = tr.mapping.invert();
    const map = (p) => Math.max(0, Math.min(out.mapping.map(back.map(p)), out.doc.content.size));
    try { out.setSelection(TextSelection.between(out.doc.resolve(map(s.anchor)), out.doc.resolve(map(s.head)))); } catch { /* keep the mapped selection */ }
  }
  if (tr.storedMarksSet) out.setStoredMarks(tr.storedMarks);
  if (tr.scrolledIntoView) out.scrollIntoView();
  for (const k of ['addToHistory', 'paste', 'uiEvent']) if (tr.getMeta(k) !== undefined) out.setMeta(k, tr.getMeta(k));
  return out;
}

/** Mark [a, b) deleted; returns where the range now starts and ends. */
function trackDelete(out, a, b, attrsFor, own) {
  const del = schema.marks.del;
  const remove = []; // the author's own insertions: really deleted
  const mark = [];
  const joins = []; // paragraph boundaries whose mark was the author's own insertion
  const marks = []; // paragraph marks to record as deleted
  const rows = []; // whole table rows in the range
  out.doc.nodesBetween(a, b, (node, pos) => {
    if (node.type === schema.nodes.table_row && pos >= a && pos + node.nodeSize <= b && rowRev(node) !== 'del') {
      // The author's own inserted row goes outright, as their own inserted text does.
      const by = rowRev(node) === 'ins' && /<w:ins\b[^>]*\bw:author="([^"]*)"/.exec(node.attrs.trPr);
      if (by && own({ attrs: { author: by[1].replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>') } })) { remove.push([pos, pos + node.nodeSize]); return false; }
      rows.push(pos);
    }
    if (node.isInline) {
      const s = Math.max(pos, a);
      const e = Math.min(pos + node.nodeSize, b);
      const ins = node.marks.find((m) => m.type.name === 'ins');
      if (own(ins)) remove.push([s, e]);
      else if (!node.marks.some((m) => m.type === del)) mark.push([s, e]);
      return false;
    }
    if (node.type === schema.nodes.paragraph) {
      const end = pos + node.nodeSize; // just after the paragraph's closing token
      // The paragraph mark is deleted when the range crosses into the next paragraph.
      if (end - 1 >= a && end < b) {
        const next = out.doc.resolve(end).nodeAfter;
        if (next && next.type === schema.nodes.paragraph) {
          if (markRev(node.attrs) === 'ins' && own(insOfMark(node.attrs))) joins.push(end);
          else if (markRev(node.attrs) !== 'del') marks.push(pos);
        }
      }
    }
    return true;
  });
  // Continue the author's own deletion right next to this one, so held-down Backspace makes one change.
  const $a = out.doc.resolve(a);
  const $b = out.doc.resolve(b);
  const near = [$a.nodeBefore, $b.nodeAfter].map((n) => n && n.marks.find((m) => m.type === del)).find((m) => own(m));
  const delMark = near || del.create(attrsFor('del'));
  const start = out.steps.length;
  for (const [s, e] of mark) out.addMark(s, e, delMark);
  for (const pos of marks) {
    const node = out.doc.nodeAt(pos);
    out.setNodeMarkup(pos, null, withMarkRev(node.attrs, revXml('del', { ...delMark.attrs, id: nextId(out.doc) })));
  }
  for (const pos of rows) {
    const node = out.doc.nodeAt(pos);
    out.setNodeMarkup(pos, null, { ...node.attrs, trPr: withRowRev(node.attrs.trPr, revXml('del', { ...delMark.attrs, id: nextId(out.doc) })) });
  }
  // Remove from the end so earlier positions stay valid.
  const cuts = [...remove.map(([s, e]) => ({ s, e })), ...joins.map((p) => ({ join: p }))].sort((x, y) => (y.s ?? y.join) - (x.s ?? x.join));
  for (const c of cuts) {
    if (c.join != null) {
      const p = since(out, start).map(c.join);
      try { out.join(p); } catch { /* not joinable */ }
    } else out.delete(since(out, start).map(c.s), since(out, start).map(c.e, -1));
  }
  return { start: since(out, start).map(a, 1), end: since(out, start).map(b, -1) };
}

function insOfMark(attrs) {
  const r = extraItem(attrs, 'rPr');
  const m = r && /<w:ins\b[^>]*\bw:author="([^"]*)"/.exec(r[1]);
  return m ? { attrs: { author: m[1].replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>') } } : null;
}

/** Insert `slice` at `at`, marked inserted; returns the position after it. */
function trackInsert(out, at, slice, attrsFor, own) {
  const { ins, del } = schema.marks;
  // Continue the author's own insertion right before, so typing makes one change.
  const before = out.doc.resolve(at).nodeBefore;
  const prev = before && before.marks.find((m) => m.type === ins);
  const insMark = own(prev) ? prev : ins.create(attrsFor('ins'));
  const start = out.steps.length;
  out.replace(at, at, slice);
  const end = since(out, start).map(at, 1);
  const from = since(out, start).map(at, -1);
  if (end > from) {
    out.removeMark(from, end, del);
    out.addMark(from, end, insMark);
    // Paragraph marks created by the insertion (Enter, pasted paragraphs) are inserted too.
    const marks = [];
    const rows = []; // whole table rows the insertion made (Insert row above / below)
    const cells = []; // whole cells it made in rows already there
    out.doc.nodesBetween(from, end, (node, pos) => {
      if (node.type === schema.nodes.paragraph && pos + node.nodeSize - 1 >= from && pos + node.nodeSize - 1 < end) marks.push(pos);
      if (node.type === schema.nodes.table_row && pos >= from && pos + node.nodeSize <= end) rows.push(pos);
      // Whole cells the insertion made in rows that were there (Insert column left / right): a cell insertion each,
      // as Word records a column inserted with changes tracked (rejected, the cells go again).
      else if (isCell(node) && pos >= from && pos + node.nodeSize <= end && !rows.some((r) => pos > r && pos < r + out.doc.nodeAt(r).nodeSize)) cells.push(pos);
      return !node.isInline;
    });
    for (const pos of cells) {
      const node = out.doc.nodeAt(pos);
      if (!cellRev(node)) out.setNodeMarkup(pos, null, { ...node.attrs, tcPr: [...(node.attrs.tcPr || []), ['cellIns', revXml('cellIns', { ...insMark.attrs, id: nextId(out.doc) })]] });
    }
    for (const pos of marks) {
      const node = out.doc.nodeAt(pos);
      if (!markRev(node.attrs)) out.setNodeMarkup(pos, null, withMarkRev(node.attrs, revXml('ins', { ...insMark.attrs, id: nextId(out.doc) })));
    }
    for (const pos of rows) {
      const node = out.doc.nodeAt(pos);
      out.setNodeMarkup(pos, null, { ...node.attrs, trPr: withRowRev(node.attrs.trPr, revXml('ins', { ...insMark.attrs, id: nextId(out.doc) })) });
    }
  }
  return end;
}

/** Apply a formatting step, recording each run's old formatting in a w:rPrChange. */
function trackFormat(out, from, to, step, attrsFor) {
  if (to <= from) return;
  const type = step.mark.type;
  if (['ins', 'del', 'link'].includes(type.name)) {
    out.step(step instanceof AddMarkStep ? new AddMarkStep(from, to, step.mark) : new RemoveMarkStep(from, to, step.mark));
    return;
  }
  const olds = [];
  out.doc.nodesBetween(from, to, (node, pos) => {
    // Formatting the author's own insertion isn't a separate change.
    const ins = node.marks.find((m) => m.type.name === 'ins');
    if (node.isText && !(ins && ins.attrs.author === attrsFor('fmt').author)) olds.push({ s: Math.max(pos, from), e: Math.min(pos + node.nodeSize, to), marks: node.marks });
    return !node.isInline;
  });
  if (step instanceof AddMarkStep) out.addMark(from, to, step.mark);
  else out.removeMark(from, to, step.mark);
  for (const { s, e, marks } of olds) {
    if (fmtChange({ marks })) continue; // already has its original formatting recorded
    const formatting = marks.filter((m) => !['ins', 'del', 'link', 'rextra'].includes(m.type.name));
    const rextra = marks.find((m) => m.type === schema.marks.rextra);
    const old = rPrXml([...formatting, ...(rextra ? [rextra] : [])]) || '<w:rPr/>';
    const a = { ...attrsFor('fmt'), id: nextId(out.doc) };
    const change = ['rPrChange', `<w:rPrChange w:id="${a.id}" w:author="${esc(a.author)}" w:date="${a.date}">${old}</w:rPrChange>`];
    // The node's current rextra (after the step) gains the change record.
    out.doc.nodesBetween(s, e, (node, pos) => {
      if (!node.isText) return true;
      const cur = node.marks.find((m) => m.type === schema.marks.rextra);
      const items = [...(cur ? cur.attrs.items.filter((it) => it[0] !== 'rPrChange') : []), change];
      out.addMark(Math.max(pos, s), Math.min(pos + node.nodeSize, e), schema.marks.rextra.create({ items }));
      return false;
    });
  }
}

/** Apply a paragraph-properties change, keeping the old properties in a w:pPrChange. */
function trackAttrs(out, pos, next, attrsFor) {
  const old = out.doc.nodeAt(pos);
  let attrs = next.attrs;
  if (old.type === schema.nodes.paragraph && !extraItem(old.attrs, 'pPrChange')) {
    const a = attrsFor('fmt');
    // The old properties, without the parts a w:pPrChange can't hold.
    const prev = pPrXml(old.attrs, ['rPr', 'sectPr', 'pPrChange']) || '<w:pPr/>';
    attrs = { ...attrs, extra: [...attrs.extra.filter((e) => e[0] !== 'pPrChange'), ['pPrChange', `<w:pPrChange w:id="${a.id}" w:author="${esc(a.author)}" w:date="${a.date}">${prev}</w:pPrChange>`]] };
  }
  out.setNodeMarkup(pos, null, attrs);
}

// --- Accepting and rejecting ---------------------------------------------------

/** The range of the change at the selection: the selection itself, or the change around the caret. */
function changeRange(state) {
  const { from, to, empty, $from } = state.selection;
  if (!empty) return { from, to };
  const parent = $from.parent;
  const pos = $from.start();
  // In an inserted or deleted row (or cell): the whole row, before its runs. (Its text carries the same deletion; taken
  // alone, rejecting it would leave the row's own w:del, and Accept All would then delete the text put back.)
  for (let d = $from.depth; d > 0; d--) {
    const n = $from.node(d);
    if ((n.type === schema.nodes.table_row && rowRev(n)) || (isCell(n) && cellRev(n))) return { from: $from.before(d), to: $from.after(d) };
  }
  // The run under (or just before) the caret, extended over neighbours with the same revision.
  let hit = null;
  parent.forEach((node, off, idx) => {
    if (!hit && pos + off <= from && from <= pos + off + node.nodeSize && hasChange(node)) hit = idx;
  });
  if (hit == null) {
    if (parent.type === schema.nodes.paragraph && hasChange(parent)) return { from: $from.before(), to: $from.after() };
    return null;
  }
  const key = (n) => n.marks.filter((m) => REV_MARKS.includes(m.type.name)).map((m) => m.attrs.id + m.type.name).join() + (fmtChange(n) ? 'f' : '');
  const k = key(parent.child(hit));
  let a = hit;
  let b = hit;
  while (a > 0 && key(parent.child(a - 1)) === k) a--;
  while (b < parent.childCount - 1 && key(parent.child(b + 1)) === k) b++;
  let start = pos;
  for (let i = 0; i < a; i++) start += parent.child(i).nodeSize;
  let end = start;
  for (let i = a; i <= b; i++) end += parent.child(i).nodeSize;
  return { from: start, to: end };
}

/**
 * Accept (`accept` true) or reject the tracked changes in [from, to].
 * Accepting keeps insertions and removes deletions; rejecting does the reverse.
 * Formatting changes are kept (accept) or their old formatting restored (reject).
 */
export function resolveChanges(tr, from, to, accept) {
  tr.setMeta('odx-resolve', true); // (refused in a document protected for tracked changes: protect.js)
  const doc = tr.doc;
  // Positions in `doc` mapped through the steps made here (the transaction may hold others before them).
  const begin = tr.steps.length;
  const mapIn = (p, assoc) => since(tr, begin).map(p, assoc);
  const removes = [];
  const unmark = [];
  const restores = [];
  const paras = [];
  const rowsGone = []; // whole rows the decision removes
  const rowsKept = []; // rows that keep their content and lose their revision
  const cellsGone = []; // cells the decision takes out of their rows (the row's cells after them shift left)
  const cellsKept = []; // cells that stay and lose their revision
  doc.nodesBetween(from, to, (node, pos) => {
    if (isCell(node) && cellRev(node) && pos >= from - 2 && pos + node.nodeSize <= to + 2) {
      const rev = cellRev(node);
      // (Not a cell merged down over rows: it only loses its mark.)
      if (((accept && rev === 'del') || (!accept && rev === 'ins')) && node.attrs.rowspan === 1) { cellsGone.push({ pos, node, row: doc.resolve(pos).before() }); return false; }
      cellsKept.push({ pos, node });
      return true;
    }
    if (node.type === schema.nodes.table_row) {
      const rev = rowRev(node);
      // (A row's own revision only when the range takes in the whole row: a run accepted inside it leaves it.)
      if (rev && pos >= from - 3 && pos + node.nodeSize <= to + 3) { // (the row's text selected counts: findChange's range)
        if ((accept && rev === 'del') || (!accept && rev === 'ins')) { rowsGone.push({ pos, node }); return false; }
        rowsKept.push({ pos, node });
      }
      return true;
    }
    if (node.type === schema.nodes.paragraph) {
      if (hasChange(node) && pos >= from - 1) paras.push({ pos, node });
      return true;
    }
    if (!node.isInline) return true;
    const s = Math.max(pos, from);
    const e = Math.min(pos + node.nodeSize, to);
    const ins = node.marks.find((m) => m.type.name === 'ins');
    const del = node.marks.find((m) => m.type.name === 'del');
    if ((accept && del) || (!accept && ins)) removes.push([s, e]);
    else {
      if (ins) unmark.push([s, e, ins]);
      if (del) unmark.push([s, e, del]);
      const f = fmtChange(node);
      if (f) restores.push({ s, e, node, f });
    }
    return false;
  });
  for (const [s, e, m] of unmark) tr.removeMark(s, e, m);
  for (const { s, e, node, f } of restores) {
    const items = f.attrs.items.filter((it) => it[0] !== 'rPrChange');
    if (accept) {
      tr.removeMark(s, e, schema.marks.rextra);
      if (items.length) tr.addMark(s, e, schema.marks.rextra.create({ items }));
    } else {
      const xml = f.attrs.items.find((it) => it[0] === 'rPrChange')[1];
      const inner = (/^<w:rPrChange\b[^>]*>([\s\S]*)<\/w:rPrChange>$/.exec(xml) || [])[1] || '';
      const old = marksFromRPr(inner.startsWith('<w:rPr') ? inner : `<w:rPr>${inner}</w:rPr>`).map((m) => schema.markFromJSON(m));
      for (const m of node.marks) if (!['ins', 'del', 'link'].includes(m.type.name)) tr.removeMark(s, e, m.type);
      for (const m of old) tr.addMark(s, e, m);
    }
  }
  // Paragraph properties and paragraph marks.
  const joins = [];
  for (const { pos, node } of paras) {
    let attrs = node.attrs;
    const change = extraItem(attrs, 'pPrChange');
    if (change) {
      if (accept) attrs = { ...attrs, extra: attrs.extra.filter((e) => e[0] !== 'pPrChange') };
      else {
        const inner = (/^<w:pPrChange\b[^>]*>([\s\S]*)<\/w:pPrChange>$/.exec(change[1]) || [])[1] || '';
        const old = paragraphAttrsFromPPr(inner.startsWith('<w:pPr') ? inner : `<w:pPr>${inner}</w:pPr>`);
        const keep = attrs.extra.filter((e) => e[0] === 'rPr' || e[0] === 'sectPr');
        attrs = { ...attrs, ...old, extra: [...old.extra.filter((e) => e[0] !== 'rPr' && e[0] !== 'sectPr'), ...keep], deco: attrs.deco };
      }
    }
    const rev = markRev(attrs);
    const end = pos + node.nodeSize;
    if (rev && end <= to + 1) {
      const removed = (accept && rev === 'del') || (!accept && rev === 'ins');
      attrs = withMarkRev(attrs, '');
      if (removed) joins.push(end);
    }
    if (attrs !== node.attrs) tr.setNodeMarkup(mapIn(pos), null, attrs);
  }
  for (const { pos, node } of rowsKept) tr.setNodeMarkup(mapIn(pos), null, { ...node.attrs, trPr: withRowRev(node.attrs.trPr, '') });
  const noCellRev = (attrs) => ({ ...attrs, tcPr: (attrs.tcPr || []).filter((t) => t[0] !== 'cellIns' && t[0] !== 'cellDel') });
  for (const { pos, node } of cellsKept) tr.setNodeMarkup(mapIn(pos), null, noCellRev(node.attrs));
  // Rows that go: a table left with none goes whole.
  const gone = [];
  const byTable = new Map();
  for (const r of rowsGone) {
    const $r = doc.resolve(r.pos);
    const tpos = $r.before($r.depth);
    if (!byTable.has(tpos)) byTable.set(tpos, { table: $r.parent, rows: [] });
    byTable.get(tpos).rows.push(r);
  }
  for (const [tpos, { table, rows }] of byTable) {
    if (rows.length === table.childCount) gone.push({ at: tpos, s: tpos, e: tpos + table.nodeSize });
    else for (const r of rows) gone.push({ at: r.pos, s: r.pos, e: r.pos + r.node.nodeSize });
  }
  // Remove content last to first, then join paragraphs whose mark went away.
  const start = tr.steps.length;
  // (Text first, so a paragraph that goes whole is empty by the time its mark goes.)
  const desc = (x, y) => y.at - x.at;
  const cuts = [...[...removes.map(([s, e]) => ({ at: s, s, e })), ...gone].sort(desc), ...joins.map((p) => ({ at: p, join: p })).sort(desc)];
  for (const c of cuts) {
    const map = (p, assoc) => since(tr, start).map(p, assoc);
    if (c.join != null) {
      const p = map(c.join);
      const $p = tr.doc.resolve(p);
      if ($p.nodeBefore && $p.nodeAfter && $p.nodeBefore.type === schema.nodes.paragraph && $p.nodeAfter.type === schema.nodes.paragraph) {
        // A paragraph that went whole (its text and its mark): the next one stays as it was, its own style, list and
        // alignment kept (the join would give it the removed paragraph's).
        const gone = !$p.nodeBefore.content.size && $p.nodeAfter;
        tr.join(p);
        if (gone) tr.setNodeMarkup(tr.doc.resolve(p - 1).before(), null, gone.attrs);
      }
    } else tr.delete(map(c.s), map(c.e, -1));
  }
  // Cells that go, last first: the row's cells after each shift left, and the grid column it leaves is a w:gridAfter
  // at the row's end (a phantom cell here, as Word's "shift cells left").
  for (const c of cellsGone.sort((x, y) => y.pos - x.pos)) {
    const at = mapIn(c.pos);
    const cell = tr.doc.nodeAt(at);
    if (!cell || !isCell(cell)) continue;
    tr.delete(at, at + cell.nodeSize);
    const rowAt = mapIn(c.row);
    const row = tr.doc.nodeAt(rowAt);
    if (!row) continue;
    const phantom = schema.nodes.table_cell.create({ colspan: cell.attrs.colspan, rowspan: 1, colwidth: cell.attrs.colwidth, colwidth0: cell.attrs.colwidth, phantom: true }, schema.nodes.paragraph.create());
    tr.insert(rowAt + row.nodeSize - 1, phantom);
    const after = [];
    tr.doc.nodeAt(rowAt).forEach((x) => after.push(x));
    let trailing = 0;
    for (let k = after.length - 1; k >= 0 && after[k].attrs.phantom; k--) trailing += after[k].attrs.colspan;
    tr.setNodeMarkup(rowAt, null, { ...row.attrs, trPr: withGridAfter(row.attrs.trPr, trailing) });
  }
  // A whole column gone (a tracked column insertion rejected): every row ends in empty grid columns, which go.
  const tables = new Set(cellsGone.map((c) => { const $c = doc.resolve(c.pos); return $c.before($c.depth - 1); }));
  for (const t0 of [...tables].sort((a, b) => b - a)) {
    const at = mapIn(t0);
    const table = tr.doc.nodeAt(at);
    if (!table || table.type !== schema.nodes.table) continue;
    const trailing = [];
    table.forEach((row) => { let k = 0; for (let i = row.childCount - 1; i >= 0 && row.child(i).attrs.phantom; i--) k++; trailing.push(k); });
    const k = Math.min(...trailing);
    if (!k) continue;
    // (Last row first, each row's last cells first: what is cut never moves what is still to cut.)
    const rowStarts = [];
    let rowAt = at + 1;
    table.forEach((row) => { rowStarts.push([rowAt, row]); rowAt += row.nodeSize; });
    for (const [start, row] of rowStarts.reverse()) {
      let p = start + row.nodeSize - 1;
      for (let i = row.childCount - 1; i >= row.childCount - k; i--) { p -= row.child(i).nodeSize; tr.delete(p, p + row.child(i).nodeSize); }
    }
    rowAt = mapIn(t0) + 1;
    tr.doc.nodeAt(mapIn(t0)).forEach((row) => {
      let left = 0;
      for (let i = row.childCount - 1; i >= 0 && row.child(i).attrs.phantom; i--) left += row.child(i).attrs.colspan;
      tr.setNodeMarkup(rowAt, null, { ...row.attrs, trPr: withGridAfter(row.attrs.trPr, left) });
      rowAt += row.nodeSize;
    });
  }
  // With no moved text left, the moves' range markers go too (Word removes them on accepting or rejecting a move).
  let moves = false;
  const markers = [];
  tr.doc.descendants((node, pos) => {
    if (node.isInline && node.marks.some((m) => REV_MARKS.includes(m.type.name) && m.attrs.move)) moves = true;
    if (node.type === schema.nodes.opaque_inline && /^move(From|To)Range(Start|End)$/.test(node.attrs.kind)) markers.push(pos);
    return !moves;
  });
  if (!moves) for (const pos of markers.reverse()) tr.delete(pos, pos + 1);
  return tr;
}

const refused = (dispatch) => { if (dispatch) announce("Changes can't be accepted or rejected: this document is protected for tracked changes"); return true; };

/** Accept or reject the change at the caret, or every change in the selection; then select the next one. */
export const resolveChange = (accept) => (state, dispatch) => {
  // Selected table cells: each cell's changes (the cells between them in the document aren't selected).
  const cells = state.selection.ranges.length > 1 ? state.selection.ranges.map((g) => ({ from: g.$from.pos, to: g.$to.pos })) : null;
  const r = cells ? { from: cells[0].from, to: cells[cells.length - 1].to } : changeRange(state);
  if (!r) return false;
  if (!resolvingAllowed(state)) return refused(dispatch);
  if (dispatch) {
    let tr = state.tr.setMeta('odx-untracked', true);
    for (const g of cells || [r]) tr = resolveChanges(tr, tr.mapping.map(g.from, 1), tr.mapping.map(g.to, -1), accept);
    const next = findChange(tr.doc, tr.mapping.map(r.to), 1);
    if (next) tr.setSelection(TextSelection.create(tr.doc, next.from, next.to));
    dispatch(tr.scrollIntoView());
    announce(`Change ${accept ? 'accepted' : 'rejected'}${next ? '' : '. No more changes'}`);
  }
  return true;
};

/** Accept or reject every change in the document. */
export const resolveAll = (accept) => (state, dispatch) => {
  if (!findChange(state.doc, 0, 1)) return false;
  if (!resolvingAllowed(state)) return refused(dispatch);
  if (dispatch) {
    dispatch(resolveChanges(state.tr.setMeta('odx-untracked', true), 0, state.doc.content.size, accept).scrollIntoView());
    announce(`All changes ${accept ? 'accepted' : 'rejected'}`);
  }
  return true;
};

/** The next (dir 1) or previous (dir -1) change from `pos`, as { from, to }. */
export function findChange(doc, pos, dir) {
  const found = [];
  doc.descendants((node, p) => {
    if (node.isInline) {
      if (hasChange(node)) found.push({ from: p, to: p + node.nodeSize });
      return false;
    }
    if (node.type === schema.nodes.paragraph && hasChange(node) && !paragraphHasInlineChange(node)) found.push({ from: p + 1, to: p + node.nodeSize - 1 });
    // (A row's: its first cell's text to its last's, so it can be selected.)
    if (node.type === schema.nodes.table_row && rowRev(node)) found.push({ from: p + 3, to: p + node.nodeSize - 3 });
    if (isCell(node) && cellRev(node)) found.push({ from: p + 2, to: p + node.nodeSize - 2 });
    return true;
  });
  // Merge touching runs into one change each.
  const merged = [];
  for (const f of found) {
    const last = merged[merged.length - 1];
    if (last && f.from <= last.to) last.to = Math.max(last.to, f.to);
    else merged.push({ ...f });
  }
  return dir > 0 ? merged.find((c) => c.from >= pos) || null : [...merged].reverse().find((c) => c.to <= pos) || null;
}
const paragraphHasInlineChange = (node) => {
  let found = false;
  node.forEach((c) => { if (hasChange(c)) found = true; });
  return found;
};

/** Select the next or previous change. */
export const gotoChange = (dir) => (state, dispatch) => {
  const sel = state.selection;
  const c = findChange(state.doc, dir > 0 ? sel.to : sel.from, dir) || findChange(state.doc, dir > 0 ? 0 : state.doc.content.size, dir);
  if (!c) return false;
  if (dispatch) dispatch(state.tr.setSelection(TextSelection.create(state.doc, c.from, c.to)).scrollIntoView());
  return true;
};

/**
 * Backspace (dir -1) and Delete (dir 1) while tracking: mark one character deleted,
 * stepping over text already marked deleted, as Word does. (Left to the browser, the
 * edit would be read back from the DOM, which loses keystrokes next to deleted text.)
 */
export const trackedDelete = (dir) => (state, dispatch) => {
  if (!tracking(state).on) return false;
  const sel = state.selection;
  if (!sel.empty) {
    if (dispatch) dispatch(state.tr.deleteSelection().scrollIntoView());
    return true;
  }
  const $p = sel.$head;
  const del = schema.marks.del;
  let pos = sel.head;
  const start = $p.start();
  const end = $p.end();
  // Skip over already-deleted text.
  for (;;) {
    const node = dir < 0 ? state.doc.resolve(pos).nodeBefore : state.doc.resolve(pos).nodeAfter;
    if (!node || !node.marks.some((m) => m.type === del)) break;
    pos += dir * (node.isText ? 1 : node.nodeSize);
  }
  if ((dir < 0 && pos <= start) || (dir > 0 && pos >= end)) {
    if (pos === sel.head) return false; // at the paragraph edge: joining is handled elsewhere
    if (dispatch) dispatch(state.tr.setSelection(TextSelection.create(state.doc, pos)));
    return true;
  }
  if (dispatch) {
    const node = dir < 0 ? state.doc.resolve(pos).nodeBefore : state.doc.resolve(pos).nodeAfter;
    const size = node.isText ? 1 : node.nodeSize;
    const tr = state.tr.setSelection(TextSelection.create(state.doc, pos));
    dispatch((dir < 0 ? tr.delete(pos - size, pos) : tr.delete(pos, pos + size)).setMeta('odx-delete', dir).scrollIntoView());
  }
  return true;
};
