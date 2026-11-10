// Footnotes and endnotes in the editor. References are atoms in the text,
// numbered in document order (1, 2, 3... for footnotes; i, ii, iii... for
// endnotes). Footnotes are drawn at the bottom of the page their reference is
// on (pages.js reserves the room); endnotes after the last text.
import { textSpansPlugin } from './textspans.js';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history, undo, redo } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import { baseKeymap } from 'prosemirror-commands';
import { schema } from './schema.js';
import { NOTE_KINDS } from '../ooxml/notes.js';
import { rPr } from '../ooxml/writer.js';
import * as cmd from './commands.js';
import { xmlId } from './util.js';
import { csFormatPlugin } from './csformat.js';

const REF = /^(footnote|endnote)Reference$/;

/** References in document order: [{ kind, id, pos, n }] (n: 1-based number within its kind). */
export function noteRefs(doc) {
  const out = [];
  const count = { footnote: 0, endnote: 0 };
  doc.descendants((node, pos) => {
    const m = node.type === schema.nodes.opaque_inline && REF.exec(node.attrs.kind);
    if (m) out.push({ kind: m[1], id: xmlId(node), pos, n: ++count[m[1]] });
  });
  return out;
}

const ROMAN = [['m', 1000], ['cm', 900], ['d', 500], ['cd', 400], ['c', 100], ['xc', 90], ['l', 50], ['xl', 40], ['x', 10], ['ix', 9], ['v', 5], ['iv', 4], ['i', 1]];
export const noteLabel = (kind, n) => {
  if (kind === 'footnote') return String(n);
  let s = '';
  for (const [r, v] of ROMAN) while (n >= v) { s += r; n -= v; }
  return s;
};

// (`around`: the formatting of the text it goes into (its font, size, colour...), which Word gives the reference too,
// under the reference style.)
const refRun = (kind, id, around = []) => schema.nodes.opaque_inline.create({
  kind: `${kind}Reference`,
  xml: `<w:r>${rPr([schema.marks.cstyle.create({ val: NOTE_KINDS[kind].refStyle }), ...around.filter((m) => !['cstyle', 'vertAlign', 'link', 'ins', 'del', 'sdt'].includes(m.type.name)), schema.marks.vertAlign.create({ val: 'superscript' })])}<w:${kind}Reference w:id="${id}"/></w:r>`,
});
const markRun = (kind) => schema.nodes.opaque_inline.create({
  kind: `${kind}Ref`,
  xml: `<w:r><w:rPr><w:rStyle w:val="${NOTE_KINDS[kind].refStyle}"/><w:vertAlign w:val="superscript"/></w:rPr><w:${kind}Ref/></w:r>`,
});

/** Insert a footnote or endnote at the cursor and open it for typing. */
export function insertNote(view, kind) {
  const { session } = view.odx;
  const map = session.notes[kind];
  let id = 1;
  for (const k of map.keys()) id = Math.max(id, Number(k) + 1);
  // (Shared, a number nobody else in the room will pick at the same moment.)
  if (session.uniqueIds) do id = 1e6 + Math.floor(Math.random() * 9e8); while (map.has(String(id)));
  id = String(id);
  const style = NOTE_KINDS[kind].textStyle;
  const para = schema.nodes.paragraph.create({ styleId: session.styles.styles.has(style) ? style : null }, [markRun(kind), schema.text(' ')]);
  map.set(id, { id, type: 'normal', xml: null, doc: schema.nodes.doc.create(null, para), dirty: true });
  view.odx.stories?.note(kind, id, map.get(id).doc); // (shared: src/collab.js)
  // (Text typed after the reference goes on in the formatting it had: the body's font, not the reference's.)
  const around = view.state.storedMarks || view.state.selection.$from.marks();
  const tr = view.state.tr.replaceSelectionWith(refRun(kind, id, around), false).setMeta('odx-untracked', true);
  tr.setStoredMarks(around);
  view.dispatch(tr.scrollIntoView());
  view.odx.pages.restory();
  // Open it once the layout has drawn it.
  setTimeout(() => {
    const el = view.dom.parentElement.querySelector(`.odx-note[data-kind="${kind}"][data-id="${id}"]`);
    if (el) editNote(view, kind, id, el);
  }, 400);
  return true;
}

let open = null;

/** The open note editor's text put into its note, the editor left open (for a save meanwhile). */
export function flushNote() {
  if (!open) return;
  open.note.doc = open.nv.state.doc;
  open.note.dirty = true;
}

export function closeNote() {
  if (!open) return;
  const { main, nv, wrap, note, kind } = open;
  open = null;
  note.doc = nv.state.doc;
  note.dirty = true;
  main.odx.stories?.note(kind, note.id, note.doc); // (shared: src/collab.js)
  main.odx.hfView = null;
  nv.destroy();
  wrap.remove();
  main.odx.pages.restory();
  main.odx.notify?.(main);
}

/** Edit a note in place over its drawn copy at the bottom of the page. */
export function editNote(main, kind, id, el) {
  closeNote();
  const note = main.odx.session.notes[kind].get(id);
  if (!note) return;
  const sheet = main.dom.parentElement;
  const box = el.getBoundingClientRect();
  const origin = sheet.getBoundingClientRect();
  const wrap = document.createElement('div');
  wrap.className = 'odx-noteedit';
  wrap.style.cssText = `top:${box.top - origin.top}px;left:${box.left - origin.left}px;width:${box.width}px`;
  sheet.append(wrap);
  const keys = keymap({
    'Mod-z': undo, 'Mod-y': redo, 'Mod-b': cmd.toggle('bold'), 'Mod-i': cmd.toggle('italic'), 'Mod-u': cmd.toggle('underline'),
    Escape: () => { closeNote(); main.focus(); return true; },
  });
  const nv = new EditorView(wrap, {
    state: EditorState.create({ doc: note.doc, plugins: [history(), keys, keymap(baseKeymap), textSpansPlugin, csFormatPlugin] }),
    attributes: { class: 'odx-note-content' },
    dispatchTransaction(tr) {
      nv.updateState(nv.state.apply(tr));
      main.odx.notify?.(nv);
    },
  });
  nv.odx = { ...main.odx, isStory: true };
  main.odx.hfView = nv;
  // The number of the note, as on the page.
  const ref = noteRefs(main.state.doc).find((r) => r.kind === kind && r.id === id);
  nv.dom.querySelectorAll(`[data-kind="${kind}Ref"]`).forEach((m) => { m.textContent = ref ? noteLabel(kind, ref.n) : ''; });
  open = { main, nv, wrap, note, kind };
  nv.focus();
  nv.dispatch(nv.state.tr.setSelection(TextSelection.atEnd(nv.state.doc)));
}

/** Clicking a drawn note edits it; clicking the body text closes the note being edited. */
export function noteClicks(main) {
  const sheet = main.dom.parentElement;
  sheet.addEventListener('mousedown', (e) => {
    if (main.odx.readOnly) return;
    if (open && open.wrap.contains(e.target)) return;
    const hit = [...sheet.querySelectorAll('.odx-note')].find((n) => {
      const r = n.getBoundingClientRect();
      return e.clientX >= r.left && e.clientX <= r.right && e.clientY >= r.top && e.clientY <= r.bottom;
    });
    if (hit) {
      e.preventDefault();
      editNote(main, hit.dataset.kind, hit.dataset.id, hit);
    } else if (open) closeNote();
  }, true);
}

/** Go to the next footnote/endnote reference after the cursor. */
export function nextNote(view, kind) {
  const refs = noteRefs(view.state.doc).filter((r) => r.kind === kind);
  if (!refs.length) return false;
  const here = view.state.selection.from;
  const r = refs.find((x) => x.pos >= here) || refs[0];
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, r.pos + 1)).scrollIntoView());
  return true;
}
