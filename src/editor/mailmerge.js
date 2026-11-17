// Mailings, as Word's tab: the recipients (records the host gives, api.setMergeData, or a CSV file picked here), Insert
// Merge Field (a MERGEFIELD as Word writes one, showing «Name»), Preview Results (each record's values shown in place of
// the fields, on screen only: the document keeps its fields) and Finish & Merge (merge.js: one document of every record).
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { decoSet } from './decoset.js';
import { Node, Slice, Fragment } from 'prosemirror-model';
import { schema } from './schema.js';
import { blocksFromXml } from '../ooxml/reader.js';
import { mergeTargets, fieldMarks, mergeRecord, mergeCursor } from '../merge.js';
import { announce } from './live.js';

const key = new PluginKey('odx-mailmerge');

/** The document's mail merge state: { records, fields, index, preview } (kept on its session). */
const stateOf = (session) => (session.merge ||= { records: [], fields: [], index: 0, preview: false });
export function mergeState(view) {
  return stateOf((view.odx?.main || view).odx.session);
}

/** The recipients: `records` (objects, field name -> value), and the field names (`fields`, else the records' keys). */
export function setRecipients(view, records, fields = null) {
  const m = mergeState(view);
  m.records = Array.isArray(records) ? records : [];
  m.fields = fields || [...new Set(m.records.flatMap((r) => Object.keys(r || {})))];
  m.index = 0;
  if (!m.records.length) m.preview = false;
  refresh(view);
}

/** A CSV file's rows as records, by its header row (RFC 4180: quoted fields, doubled quotes, line breaks in quotes). */
export function parseCsv(text) {
  const rows = [];
  let row = [], cell = '', q = false;
  const s = String(text).replace(/^\uFEFF/, '');
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (q) {
      if (c === '"' && s[i + 1] === '"') { cell += '"'; i++; } else if (c === '"') q = false; else cell += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(cell); cell = ''; } else if (c === '\n' || c === '\r') {
      if (c === '\r' && s[i + 1] === '\n') i++;
      row.push(cell); rows.push(row); row = []; cell = '';
    } else cell += c;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  const [head, ...body] = rows.filter((r) => r.some((x) => x !== ''));
  if (!head) return { fields: [], records: [] };
  const fields = head.map((h) => h.trim());
  return { fields, records: body.map((r) => Object.fromEntries(fields.map((f, i) => [f, r[i] ?? '']))) };
}

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Insert Merge Field: a MERGEFIELD at the cursor, as Word writes it (its result «Name»), in the formatting there. */
export function insertMergeField(view, name) {
  const main = view.odx?.main || view;
  const { state } = main;
  const instr = /\s/.test(name) ? `"${name}"` : name;
  const fld = (t) => `<w:r><w:fldChar w:fldCharType="${t}"/></w:r>`;
  const xml = `<w:p>${fld('begin')}<w:r><w:instrText xml:space="preserve"> MERGEFIELD ${esc(instr)} </w:instrText></w:r>${fld('separate')}<w:r><w:t>${esc(`«${name}»`)}</w:t></w:r>${fld('end')}</w:p>`;
  const para = Node.fromJSON(schema, blocksFromXml(xml)[0]);
  const marks = state.storedMarks || state.selection.$from.marks();
  const kids = [];
  para.forEach((n) => kids.push(n.mark(marks)));
  main.dispatch(state.tr.replaceSelection(new Slice(Fragment.fromArray(kids), 0, 0)).scrollIntoView());
  main.focus();
  return true;
}

/** Preview Results on or off (`on` omitted: toggled), at record `index`. */
export function setPreview(view, on, index) {
  const m = mergeState(view);
  m.preview = !!(on ?? !m.preview) && m.records.length > 0;
  if (index != null) m.index = Math.max(0, Math.min(m.records.length - 1, index));
  refresh(view);
  if (m.records.length) announce(m.preview ? `Previewing record ${m.index + 1} of ${m.records.length}` : 'Preview off');
  return true;
}

function refresh(view) {
  const main = view.odx?.main || view;
  main.dispatch(main.state.tr.setMeta(key, true).setMeta('addToHistory', false));
  // Headers and footers are drawn again too, with the record's values (previewStory).
  main.odx?.pages?.restory();
}

/**
 * A header or footer as Preview Results shows it: its merge and IF fields filled for the record the body's NEXT and
 * NEXTIF fields moved on to, as Word's preview (else as it is).
 */
export function previewStory(session, doc) {
  const m = session.merge;
  return doc && m?.preview && m.records.length ? mergeRecord(doc, mergeCursor(m.records, m.last ?? m.index, { preview: true })) : doc;
}

/** The textblocks of `doc` (each { node, start: its content's start }) and a lookup: the one whose content holds a position. */
function textblocks(doc) {
  const list = [];
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    list.push({ node, start: pos + 1 });
    return false;
  });
  return (p) => {
    let lo = 0, hi = list.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const b = list[mid];
      if (p < b.start) hi = mid - 1;
      else if (p > b.start + b.node.content.size) lo = mid + 1;
      else return b;
    }
    return null;
  };
}

/**
 * Each field shown as the record's value: its own content hidden, the value drawn in its place in its formatting. As
 * Word's preview: from the record shown on, NEXT and NEXTIF moving on to the records after it, SKIPIF and MERGESEQ
 * ignored (they count only in the merge).
 */
function previewDecos(doc, m) {
  const decos = [];
  const ctx = mergeCursor(m.records, m.index, { preview: true });
  // (IF fields too, their text for the record as in the merged document: merge.js mergeTargets.)
  const blockAt = textblocks(doc);
  for (const { field: f, value } of mergeTargets(doc, ctx)) {
    // (A field in one paragraph. Its textblock found by its position, not resolved from the document's start, and its
    // formatting looked for there alone: a position resolved goes through the blocks before it, quadratic in a long
    // letter's fields.)
    const block = blockAt(f.from);
    if (block ? f.to > block.start + block.node.content.size : doc.resolve(f.from).parent !== doc.resolve(f.to).parent) continue;
    const marks = fieldMarks(doc, f, block);
    decos.push(Decoration.inline(f.from, f.to, { class: 'odx-mm-hide' }));
    decos.push(Decoration.widget(f.from, () => { const s = document.createElement('span'); s.className = 'odx-mm-val'; s.textContent = value; return s; },
      { side: -1, marks: marks || [], key: `mm:${m.index}:${f.from}:${value}` }));
  }
  // (The record the body ended on, for the headers and footers: previewStory.)
  m.last = ctx.i;
  return decoSet(doc, decos);
}

/** The preview's decorations, rebuilt as the document or the record changes. */
export function mailMergePlugin(session) {
  return new Plugin({
    key,
    state: {
      init: () => DecorationSet.empty,
      apply(tr, set, _old, state) {
        if (!tr.docChanged && !tr.getMeta(key)) return set;
        const m = stateOf(session);
        return m.preview && m.records.length ? previewDecos(state.doc, m) : DecorationSet.empty;
      },
    },
    props: { decorations: (state) => key.getState(state) },
  });
}
