// Comments in the editor: the commented text is highlighted, and the comment
// threads are drawn in the markup column (see balloons.js). The document keeps
// Word's markers (commentRangeStart/End and a commentReference run) as atoms;
// the comment text lives in session.comments.
import { announce } from './live.js';
import { Plugin, PluginKey, TextSelection } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { schema } from './schema.js';
import { now, xmlId } from './util.js';
import { commentsAllowed } from './protect.js';

const KINDS = new Set(['commentRangeStart', 'commentRangeEnd', 'commentReference']);

/** Every comment marker in the document: Map id -> { start, end, ref } (positions). */
export function commentMarkers(doc) {
  const out = new Map();
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.opaque_inline || !KINDS.has(node.attrs.kind)) return;
    const id = xmlId(node);
    const m = out.get(id) || {};
    if (node.attrs.kind === 'commentRangeStart') m.start = pos;
    else if (node.attrs.kind === 'commentRangeEnd') m.end = pos;
    else m.ref = pos;
    out.set(id, m);
  });
  return out;
}

const key = new PluginKey('odx-comments');

/** Highlights commented text; `data-cid` lets a click find its comment. */
export function commentsPlugin(session) {
  const build = (doc) => {
    const decos = [];
    const resolved = new Set(session.comments.filter((c) => c.done).map((c) => c.id));
    for (const [id, m] of commentMarkers(doc)) {
      if (m.start == null || m.end == null || m.end <= m.start) continue;
      const active = session.activeComment === id;
      decos.push(Decoration.inline(m.start, m.end, { class: `odx-cmt${resolved.has(id) ? ' odx-cmt-done' : ''}${active ? ' odx-cmt-active' : ''}`, 'data-cid': id }));
    }
    return DecorationSet.create(doc, decos);
  };
  return new Plugin({
    key,
    state: {
      init: (_, state) => build(state.doc),
      apply: (tr, old, _, state) => (tr.docChanged || tr.getMeta(key) ? build(state.doc) : old),
    },
    props: {
      decorations: (state) => key.getState(state),
      handleClick(view, pos, event) {
        const hit = event.target.closest && event.target.closest('[data-cid]');
        const id = hit ? hit.getAttribute('data-cid') : null;
        if (id !== session.activeComment) setActive(view, id);
        return false;
      },
    },
  });
}

/** In a shared document (src/collab.js), the comments go to the room after each change here. */
const share = (view) => view.odx.comments?.push(view.odx.session);

/** Redraw highlights and balloons after a comment change. */
export function refresh(view) {
  share(view);
  view.dispatch(view.state.tr.setMeta(key, true).setMeta('addToHistory', false));
  view.odx.pages.update();
  view.odx.pages.balloons();
}

export function setActive(view, id) {
  view.odx.session.activeComment = id;
  refresh(view);
}

export const marker = (kind, id) => schema.nodes.opaque_inline.create({
  kind,
  xml: kind === 'commentReference'
    ? `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`
    : `<w:${kind} w:id="${id}"/>`,
});

function nextId(session, doc) {
  // (Shared, a number nobody else in the room will pick at the same moment: the next one up, two editors would.)
  if (session.uniqueIds) {
    const used = new Set([...session.comments.map((c) => String(c.id)), ...commentMarkers(doc).keys()].map(String));
    let id;
    do id = String(1e6 + Math.floor(Math.random() * 9e8)); while (used.has(id));
    return id;
  }
  let n = -1;
  for (const c of session.comments) n = Math.max(n, Number(c.id) || 0);
  for (const id of commentMarkers(doc).keys()) n = Math.max(n, Number(id) || 0);
  return String(n + 1);
}

const initials = (name) => name.split(/\s+/).filter(Boolean).map((w) => w[0].toUpperCase()).join('').slice(0, 3);

export function newComment(session, doc, author, extra = {}) {
  const c = { id: nextId(session, doc), author, initials: initials(author), date: now(), text: '', paraId: null, parentId: null, done: false, xml: null, ...extra };
  session.comments.push(c);
  session.commentsDirty = true;
  return c;
}

/** New comment on the selection (or the word at the cursor). Opens it for typing. */
export function addComment(view) {
  if (view.odx.readOnly || view.odx.features?.comments === false) return null;
  if (!commentsAllowed(view.state)) { announce('This document is protected: it can only be read'); return null; }
  const { session } = view.odx;
  const { state } = view;
  let { from, to } = state.selection;
  if (from === to) {
    const $p = state.selection.$from;
    const text = $p.parent.textBetween(0, $p.parent.content.size, '￼', '￼');
    let a = $p.parentOffset, b = a;
    while (a > 0 && /\w/.test(text[a - 1])) a--;
    while (b < text.length && /\w/.test(text[b])) b++;
    from = $p.start() + a;
    to = $p.start() + b;
  }
  const c = newComment(session, state.doc, view.odx.author || 'Author');
  const tr = state.tr;
  tr.insert(to, [marker('commentRangeEnd', c.id), marker('commentReference', c.id)]);
  tr.insert(from, marker('commentRangeStart', c.id));
  tr.setSelection(TextSelection.create(tr.doc, from + 1, to + 1));
  tr.setMeta('odx-untracked', true);
  session.activeComment = c.id;
  session.editingComment = c.id;
  view.dispatch(tr);
  view.odx.pages.update();
  view.odx.pages.balloons();
  share(view);
  announce('Comment added. Type it in the comment box');
  return c;
}

/** Reply to a thread; the reply gets its own markers next to the parent's, as Word writes them. */
export function replyTo(view, parentId, text) {
  if (!commentsAllowed(view.state)) return null;
  const { session } = view.odx;
  const m = commentMarkers(view.state.doc).get(parentId) || {};
  const c = newComment(session, view.state.doc, view.odx.author || 'Author', { parentId, text });
  const tr = view.state.tr.setMeta('odx-untracked', true);
  const ins = [];
  if (m.ref != null) ins.push([m.ref + 1, [marker('commentRangeEnd', c.id), marker('commentReference', c.id)]]);
  else if (m.end != null) ins.push([m.end + 1, [marker('commentRangeEnd', c.id), marker('commentReference', c.id)]]);
  if (m.start != null) ins.push([m.start + 1, [marker('commentRangeStart', c.id)]]);
  ins.sort((a, b) => b[0] - a[0]).forEach(([pos, nodes]) => tr.insert(pos, nodes));
  view.dispatch(tr);
  view.odx.pages.update();
  view.odx.pages.balloons();
  share(view);
  announce('Reply posted');
  return c;
}

/** The ids of a comment and its replies. */
const thread = (session, id) => [id, ...session.comments.filter((c) => c.parentId === id).map((c) => c.id)];

export function setText(view, id, text) {
  if (!commentsAllowed(view.state)) return;
  const c = view.odx.session.comments.find((x) => x.id === id);
  if (!c) return;
  c.text = text;
  c.xml = null;
  view.odx.session.commentsDirty = true;
  refresh(view);
}

export function setResolved(view, id, done) {
  if (!commentsAllowed(view.state)) return;
  const { session } = view.odx;
  for (const tid of thread(session, id)) {
    const c = session.comments.find((x) => x.id === tid);
    c.done = done;
  }
  session.commentsDirty = true;
  refresh(view);
  announce(done ? 'Comment resolved' : 'Comment reopened');
}

/** Delete a comment (with its replies when it starts a thread) and its markers. */
export function deleteComment(view, id) {
  if (!commentsAllowed(view.state)) return;
  const { session } = view.odx;
  const ids = new Set(thread(session, id));
  session.comments = session.comments.filter((c) => !ids.has(c.id));
  session.commentsDirty = true;
  if (ids.has(session.activeComment)) session.activeComment = null;
  const tr = view.state.tr.setMeta('odx-untracked', true);
  const doomed = [];
  view.state.doc.descendants((node, pos) => {
    if (node.type === schema.nodes.opaque_inline && KINDS.has(node.attrs.kind) && ids.has(xmlId(node))) doomed.push(pos);
  });
  for (const pos of doomed.reverse()) tr.delete(pos, pos + 1);
  view.dispatch(tr);
  view.odx.pages.update();
  view.odx.pages.balloons();
  share(view);
  announce(ids.size > 1 ? 'Comment and its replies deleted' : 'Comment deleted');
}

/** Top-level comments in document order (by their start marker). */
export function orderedThreads(view) {
  const { session } = view.odx;
  const marks = commentMarkers(view.state.doc);
  const at = (c) => { const m = marks.get(c.id) || {}; return m.start ?? m.end ?? m.ref ?? Infinity; };
  return session.comments.filter((c) => c.parentId == null && marks.has(c.id)).sort((a, b) => at(a) - at(b));
}

/** Move to the previous/next comment from the cursor. */
export function gotoComment(view, dir) {
  const list = orderedThreads(view);
  if (!list.length) return false;
  const marks = commentMarkers(view.state.doc);
  const pos = (c) => { const m = marks.get(c.id); return m.start ?? m.end ?? m.ref; };
  const here = view.state.selection.from;
  const c = dir > 0 ? list.find((x) => pos(x) > here) || list[0] : [...list].reverse().find((x) => pos(x) < here - 1) || list[list.length - 1];
  const p = pos(c) + 1;
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, Math.min(p, view.state.doc.content.size))).scrollIntoView());
  setActive(view, c.id);
  view.focus();
  return true;
}
