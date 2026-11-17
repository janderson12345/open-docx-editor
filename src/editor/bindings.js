// Content controls bound to custom XML (ooxml/databind.js) in the editor: typing into one puts its value in the node it
// is bound to, and every control bound to the same node, in the body and in the headers and footers, shows it.
import { Plugin, PluginKey } from 'prosemirror-state';
import { Transform } from 'prosemirror-transform';
import { schema } from './schema.js';
import { listFields, formsKey, writeField } from './fields.js';
import { fieldInfo } from '../ooxml/reader.js';
import { bindingOf, boundDisplay, boundPr, storeValue } from '../ooxml/databind.js';

const PLC = /<w:showingPlcHdr(?![^>]*w:val="(?:0|false)")[^>]*\/>/;
const leafText = (leaf) => (leaf.type === schema.nodes.hard_break ? '\n' : '');
const isPh = (m) => m.type === schema.marks.cstyle && m.attrs.val === 'PlaceholderText';

/**
 * The bound controls in `doc` that find a node in `store`: [{ key, kind, pr, from, to, text, value, checked, cell?,
 * block?, field?, b, node }] (text: shown, '\n' for a line break; value: what goes in the node, '' while the
 * placeholder shows).
 */
export function boundFields(doc, store) {
  const out = [];
  if (!store || store.empty) return out;
  const add = (f) => {
    const b = bindingOf(f.pr);
    const node = b && store.target(b);
    if (!node) return;
    f.b = b;
    f.node = node;
    f.value = f.placeholder ? (f.kind === 'checkbox' ? f.value : '') : storeValue(f.pr, f.kind, f);
    out.push(f);
  };
  for (const f of listFields(doc)) {
    if (f.legacy || !/<w:dataBinding\b/.test(f.pr || '')) continue;
    const text = doc.textBetween(f.from, f.to, '\n', leafText);
    const placeholder = f.block ? PLC.test(f.pr) : f.mark?.attrs.phText != null && f.text === f.mark.attrs.phText;
    add({ key: f.key, kind: f.kind, pr: f.pr, from: f.from, to: f.to, text, checked: f.checked, value: f.value, placeholder, block: f.block, field: f });
  }
  // Controls around a table cell (their wrapper on the cell: reader.js readTable).
  doc.descendants((node, pos) => {
    if (node.type.spec.tableRole !== 'cell' && node.type.spec.tableRole !== 'header_cell') return true;
    const open = node.attrs.sdt?.open;
    if (open && /<w:dataBinding\b/.test(open)) {
      const kind = fieldInfo(open).kind;
      if (kind !== 'checkbox') add({ key: `c${node.attrs.sdt.g}`, kind, pr: open, from: pos, to: pos + node.nodeSize, text: node.textBetween(0, node.content.size, '\n', leafText), placeholder: PLC.test(open), cell: node });
    }
    return true;
  });
  return out;
}

/** Text with its line breaks as hard breaks, in `marks`. */
function textNodes(text, marks) {
  const out = [];
  text.split('\n').forEach((line, i) => {
    if (i) out.push(schema.nodes.hard_break.create(null, null, marks));
    if (line) out.push(schema.text(line, marks));
  });
  return out;
}

/** The marks of the first text from..to (or none). */
function firstMarks(doc, from, to) {
  let marks = null;
  doc.nodesBetween(from, to, (n) => { if (!marks && n.isText) marks = n.marks; return !marks; });
  return marks || [];
}

/** One control made to show `value` (its node's) inside `tr`; false when it already does. */
function show(tr, f, value) {
  const shown = boundDisplay(f.pr, f.kind, value);
  if (!shown) {
    // An emptied node: the control shows its placeholder (an inline one; one around a paragraph or a cell is left).
    if (f.placeholder || f.block || f.cell || f.kind === 'checkbox') return false;
    writeField(tr, f.field, '');
    return true;
  }
  if (!f.placeholder && f.text === shown.text && (shown.checked == null || shown.checked === f.checked)) return false;
  const pr = boundPr(f.pr, shown);
  if (f.cell) {
    const first = f.cell.firstChild;
    const marks = firstMarks(tr.doc, f.from, f.to).filter((m) => !isPh(m));
    const para = schema.nodes.paragraph.create(first?.type === schema.nodes.paragraph ? first.attrs : null, textNodes(shown.text, marks));
    tr.replaceWith(f.from + 1, f.to - 1, para);
    tr.setNodeMarkup(f.from, null, { ...f.cell.attrs, sdt: { ...f.cell.attrs.sdt, open: pr } });
    return true;
  }
  if (f.block) {
    const first = tr.doc.nodeAt(f.from);
    const marks = firstMarks(tr.doc, f.from, f.to).filter((m) => !isPh(m));
    const para = schema.nodes.paragraph.create({ ...first.attrs, sdt: { ...first.attrs.sdt, open: pr } }, textNodes(shown.text, marks));
    tr.replaceWith(f.from, f.to, para);
    return true;
  }
  const m = f.field.mark;
  const mark = schema.marks.sdt.create({ ...m.attrs, pr, phText: null });
  const marks = mark.addToSet(firstMarks(tr.doc, f.from, f.to).filter((x) => !(x.type === schema.marks.sdt && x.attrs.key === m.attrs.key) && !isPh(x)));
  tr.replaceWith(f.from, f.to, textNodes(shown.text, marks));
  return true;
}

/** Every bound control in `tr`'s doc made to show its node's value; true when any changed. */
export function syncInto(tr, store) {
  let changed = false;
  // (Last first, so the positions of those before stay as they were.)
  for (const f of boundFields(tr.doc, store).reverse()) {
    if (show(tr, f, store.value(f.b))) changed = true;
  }
  return changed;
}

/** `doc` with its bound controls showing their nodes' values (the same doc when they do). */
export function syncBound(doc, store) {
  if (!store || store.empty) return doc;
  const tr = new Transform(doc);
  return syncInto(tr, store) ? tr.doc : doc;
}

export const bindingsKey = new PluginKey('odx-bindings');

/**
 * Edits to bound controls go into their nodes, and the controls bound to the same nodes follow: in this editor at
 * once, in the headers and footers (or, from a header's editor, the body) through the session. `story`: the rId of the
 * header or footer this editor edits (null for the body); `main`, for a header's editor, the body's view.
 */
export function bindingsPlugin(session, { story = null } = {}) {
  const store = session?.bindings;
  return new Plugin({
    key: bindingsKey,
    appendTransaction(trs, oldState, state) {
      if (!store || store.empty || !trs.some((t) => t.docChanged) || trs.every((t) => t.getMeta(bindingsKey))) return null;
      const was = new Map(boundFields(oldState.doc, store).map((f) => [f.key, f.value]));
      let edited = false;
      for (const f of boundFields(state.doc, store)) {
        if (was.has(f.key) && was.get(f.key) !== f.value && store.set(f.b, f.value)) edited = true;
      }
      if (!edited) return null;
      const tr = state.tr;
      syncInto(tr, store);
      // The headers and footers (but this one) follow; a header's edit reaches the body once it's drawn (view below).
      session.syncBindings(null, story);
      if (story) session._bindBody = true;
      return tr.setMeta(bindingsKey, true).setMeta('odx-untracked', true).setMeta(formsKey, { allow: true }).setMeta('addToHistory', true);
    },
    view: () => ({
      update(view) {
        if (story) {
          // (From a header's editor: the body brought up to date, then the pages redrawn.)
          const main = view.odx?.main;
          if (session._bindBody && main && !main.isDestroyed) {
            session._bindBody = false;
            const tr = main.state.tr;
            if (syncInto(tr, store)) main.dispatch(tr.setMeta(bindingsKey, true).setMeta('odx-untracked', true).setMeta(formsKey, { allow: true }));
          }
        }
        if (session._bindRestory && !story && view.odx?.pages) {
          session._bindRestory = false;
          view.odx.pages.restory();
        }
      },
    }),
  });
}
