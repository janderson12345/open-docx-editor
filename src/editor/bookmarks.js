// Bookmarks: named places in the document (w:bookmarkStart / w:bookmarkEnd atoms).
import { TextSelection } from 'prosemirror-state';
import { schema } from './schema.js';

const attrOf = (node, a) => (new RegExp(`w:${a}="([^"]*)"`).exec(node.attrs.xml) || [])[1];
const isKind = (node, kind) => node.type === schema.nodes.opaque_inline && node.attrs.kind === kind;

/** Visible bookmarks (Word hides ones starting with "_"): [{ name, id, pos }] in document order. */
export function listBookmarks(doc, hidden = false) {
  const out = [];
  doc.descendants((node, pos) => {
    if (!isKind(node, 'bookmarkStart')) return;
    const name = attrOf(node, 'name') || '';
    if (hidden || !name.startsWith('_')) out.push({ name, id: attrOf(node, 'id'), pos });
  });
  return out;
}

export const validName = (name) => /^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(name);

/** Add (or move) bookmark `name` around the selection. */
export function addBookmark(view, name) {
  if (!validName(name)) return 'Bookmark names start with a letter and use only letters, numbers and _ (up to 40).';
  deleteBookmark(view, name);
  const { state } = view;
  let id = 0;
  state.doc.descendants((n) => { if (isKind(n, 'bookmarkStart') || isKind(n, 'bookmarkEnd')) id = Math.max(id, Number(attrOf(n, 'id')) + 1 || 0); });
  const { from, to } = state.selection;
  const mk = (kind, xml) => schema.nodes.opaque_inline.create({ kind, xml });
  const tr = state.tr.setMeta('odx-untracked', true);
  tr.insert(to, mk('bookmarkEnd', `<w:bookmarkEnd w:id="${id}"/>`));
  tr.insert(from, mk('bookmarkStart', `<w:bookmarkStart w:id="${id}" w:name="${name}"/>`));
  view.dispatch(tr);
  return null;
}

export function deleteBookmark(view, name) {
  const doc = view.state.doc;
  const b = listBookmarks(doc, true).find((x) => x.name === name);
  if (!b) return;
  const del = [];
  doc.descendants((n, pos) => {
    if ((isKind(n, 'bookmarkStart') || isKind(n, 'bookmarkEnd')) && attrOf(n, 'id') === b.id) del.push(pos);
  });
  const tr = view.state.tr.setMeta('odx-untracked', true);
  for (const pos of del.reverse()) tr.delete(pos, pos + 1);
  view.dispatch(tr);
}

export function goToBookmark(view, name) {
  const b = listBookmarks(view.state.doc, true).find((x) => x.name === name);
  if (!b) return;
  view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(b.pos + 1))).scrollIntoView());
  view.focus();
}
