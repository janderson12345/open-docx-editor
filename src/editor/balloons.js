// Word-style markup area: a column to the right of the page with a balloon for
// each deletion and formatting change, level with the line it belongs to.
import { schema, revXmlAttrs } from './schema.js';
import { paragraphAttrsFromPPr } from '../ooxml/reader.js';
import { commentMarkers, orderedThreads, setActive, setText, replyTo, setResolved, deleteComment } from './comments.js';
import { announce } from './live.js';
import { wordDate } from './util.js';

export const MARKUP_WIDTH = 280;

const PPR_LABELS = {
  before: 'Space Before', after: 'Space After', line: 'Line spacing', indLeft: 'Indent: Left', indRight: 'Indent: Right',
  indFirst: 'First line', keepNext: 'Keep with next', keepLines: 'Keep lines together', borders: 'Border', shading: 'Pattern', tabs: 'Tab stops',
  pageBreakBefore: 'Page break before', contextualSpacing: "Don't add space between paragraphs of the same style",
};
const ALIGN = { left: 'Left', center: 'Centered', right: 'Right', both: 'Justified' };

function describePara(now, before, session) {
  const out = [];
  if (now.styleId !== before.styleId) out.push(session.styles.styles.get(now.styleId || session.styles.defaultParagraph)?.name || now.styleId || 'Normal');
  if (now.align !== before.align) out.push(ALIGN[now.align] || 'Left');
  const a = now.view || {}, b = before.view || {};
  for (const k of Object.keys(PPR_LABELS)) {
    if (JSON.stringify(a[k]) === JSON.stringify(b[k])) continue;
    if (k === 'line') out.push(`Line spacing: ${a.lineRule && a.lineRule !== 'auto' ? `${a.lineRule} ${(a.line || 0) / 20} pt` : a.line == null || a.line === 240 ? 'single' : a.line === 480 ? 'double' : `Multiple ${(a.line / 240).toFixed(2)} li`}`);
    else if (['before', 'after', 'indLeft', 'indRight', 'indFirst'].includes(k)) out.push(`${PPR_LABELS[k]}: ${(a[k] || 0) / 20} pt`);
    else out.push(PPR_LABELS[k]);
  }
  if (now.numId !== before.numId) out.push(now.numId ? 'Numbered + Level: ' + (now.ilvl + 1) : 'No bullets or numbering');
  return out.join(', ');
}

/** The balloons for the document, top to bottom: { el, author, label, text, kind }. */
function collect(view, session) {
  const items = [];
  const nodeEl = (pos) => { const d = view.nodeDOM(pos); return d && d.nodeType === 1 ? d : null; };
  view.state.doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.paragraph) return true;
    const ch = (node.attrs.extra || []).find((e) => e[0] === 'pPrChange');
    if (ch) {
      const old = paragraphAttrsFromPPr((/^<w:pPrChange\b[^>]*>([\s\S]*)<\/w:pPrChange>$/.exec(ch[1]) || [])[1] || '');
      const r = revXmlAttrs(ch[1]);
      items.push({ el: nodeEl(pos), author: r.author, kind: 'fmt', label: 'Formatted', text: describePara(node.attrs, old, session) });
    }
    // Runs of deleted text and of formatting changes, merged while adjacent and by the same author.
    let run = null;
    node.forEach((c, off) => {
      const del = c.marks.find((m) => m.type === schema.marks.del);
      const fmt = c.marks.map((m) => m.type === schema.marks.rextra && (m.attrs.items || []).find((e) => e[0] === 'rPrChange')).find(Boolean);
      const kind = del ? 'del' : fmt ? 'fmt' : null;
      const author = del ? del.attrs.author : fmt ? revXmlAttrs(fmt[1]).author : null;
      if (kind && run && run.kind === kind && run.author === author) { run.text += c.textContent; return; }
      if (run) items.push(run);
      run = kind ? { kind, author, label: kind === 'del' ? 'Deleted' : 'Formatted', text: kind === 'del' ? c.textContent : 'Font', pos: pos + 1 + off } : null;
    });
    if (run) items.push(run);
    return false;
  });
  for (const it of items) if (it.pos != null) {
    const { node, offset } = view.domAtPos(it.pos + 1);
    const at = node.nodeType === 1 ? node.childNodes[offset] || node : node;
    it.el = at.nodeType === 1 ? at : at.parentElement;
    if (it.kind === 'fmt') it.text = '';
  }
  return items.filter((i) => i.el);
}

const drawnSig = new WeakMap(); // markup column -> what its balloons were drawn for

/** Draw the markup column (or remove it when there is nothing to show). */
export function drawBalloons(view, session, sheet, height) {
  let col = sheet.querySelector(':scope > .odx-markup');
  const items = [...collect(view, session), ...commentItems(view)];
  const top0 = sheet.getBoundingClientRect().top;
  for (const it of items) it.y = it.el.getBoundingClientRect().top - top0;
  items.sort((a, b) => a.y - b.y);
  sheet.classList.toggle('odx-has-markup', items.length > 0);
  if (!items.length) { col?.remove(); return; }
  if (!col) {
    col = document.createElement('div');
    col.className = 'odx-markup';
    col.contentEditable = 'false';
    sheet.append(col);
  }
  col.style.height = `${height}px`;
  // (Tracked changes' balloons where they were: left as drawn. Comments' have boxes being typed in: drawn each time.)
  const sig = items.some((it) => it.kind === 'comment') ? null : items.map((it) => `${it.kind}|${it.author}|${it.label}|${it.text}|${Math.round(it.y)}`).join('\n');
  if (sig != null && drawnSig.get(col) === sig) return;
  drawnSig.set(col, sig);
  // A comment box being typed in keeps the keyboard through the redraw (else focus fell to the page for a moment, and
  // NVDA said "document" between "Add a comment" and itself again).
  const had = col.contains(document.activeElement) && document.activeElement.dataset.odxBox ? document.activeElement : null;
  // All drawn first, then all measured, then placed: reading each one's height as it went in laid the page out again
  // for each (7,200 changes' balloons took seconds a keystroke).
  const balloons = items.map((it) => {
    if (it.kind === 'comment') return commentBalloon(view, it.thread);
    const b = document.createElement('div');
    b.className = `odx-balloon odx-balloon-${it.kind}`;
    const who = document.createElement('b');
    who.textContent = it.author || 'Unknown';
    const what = document.createElement('div');
    const lab = document.createElement('b');
    lab.textContent = `${it.label}${it.text ? ':' : ''} `;
    what.append(lab, it.text || '');
    b.append(who, what);
    return b;
  });
  col.replaceChildren(...balloons);
  const heights = balloons.map((b) => b.offsetHeight);
  let next = 0;
  items.forEach((it, k) => {
    const y = Math.max(next, it.y);
    balloons[k].style.top = `${y}px`;
    next = y + heights[k] + 4;
  });
  const again = had && col.querySelector(`textarea[data-odx-box="${had.dataset.odxBox}"]`);
  if (again) { again.focus(); again.setSelectionRange(had.selectionStart, had.selectionEnd); }
}

function commentItems(view) {
  const marks = commentMarkers(view.state.doc);
  const out = [];
  for (const c of orderedThreads(view)) {
    const m = marks.get(c.id);
    const dom = view.nodeDOM(m.start ?? m.end ?? m.ref);
    const el = dom && (dom.nodeType === 1 ? dom : dom.parentElement);
    if (el) out.push({ kind: 'comment', el, thread: c });
  }
  return out;
}

const h = (tag, cls, text) => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text != null) e.textContent = text;
  return e;
};
const when = (iso, utc) => {
  const d = wordDate(iso, utc);
  return !d ? '' : d.toLocaleString(undefined, { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' });
};

/** A text box with Post / Cancel. */
// Unposted text survives the column being redrawn.
const drafts = new Map();
// Boxes that have had the focus once: drawn again (the column redrawn on an edit elsewhere), they don't take it back.
const focused = new Set();
// After a comment or reply is posted or cancelled, the keyboard goes back to the document (the box it was typed in is
// redrawn away: focus fell to the page, NVDA saying only "document").
const back = (view) => setTimeout(() => { if (!view.isDestroyed && !view.dom.contains(document.activeElement)) view.focus(); }, 0);

function composer(key, placeholder, initial, onPost, onCancel, findSel = null) {
  const wrap = h('div', 'odx-cmt-compose');
  const ta = h('textarea');
  ta.dataset.odxBox = key;
  ta.placeholder = placeholder;
  ta.value = drafts.has(key) ? drafts.get(key) : initial || '';
  ta.oninput = () => { drafts.set(key, ta.value); findSel?.clear(); };
  ta.addEventListener('mousedown', () => findSel?.clear());
  const done = (fn) => () => { drafts.delete(key); focused.delete(key); fn(); };
  onCancel = done(onCancel);
  const post0 = onPost;
  onPost = (t) => { drafts.delete(key); focused.delete(key); post0(t); };
  ta.rows = 2;
  const post = h('button', 'odx-cmt-post', 'Post');
  const cancel = h('button', '', 'Cancel');
  post.type = cancel.type = 'button';
  post.onclick = () => ta.value.trim() && onPost(ta.value.trim());
  cancel.onclick = onCancel;
  ta.onkeydown = (e) => {
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); post.onclick(); }
    if (e.key === 'Escape') onCancel();
    e.stopPropagation();
  };
  const row = h('div', 'odx-cmt-actions');
  row.append(post, cancel);
  wrap.append(ta, row);
  // (The caret at the end, or the text Find chose selected.)
  setTimeout(() => {
    if (!ta.isConnected) return;
    // Find's match stays selected (through redraws of the column) until the comment is typed in or clicked, and the
    // keys stay in the find bar while it has them.
    let select = findSel?.get();
    if (!select && document.activeElement === ta) return; // (focused, with its selection, by a redraw of the column)
    // (Not while someone is in the document: Find's match there is over, and the comment mustn't take the keys.)
    if (select && document.activeElement?.closest?.('.odx-content, .odx-hfedit')) { findSel.clear(); return; }
    // (A match Find opened here is selected, the keys left where they are: in the find bar. A click edits it.)
    if (select) { focused.add(key); ta.setSelectionRange(select.from, select.to); ta.scrollIntoView({ block: 'center' }); return; }
    if (document.activeElement?.closest?.('.odx-findbar') || focused.has(key)) return;
    focused.add(key);
    ta.focus();
    ta.selectionStart = ta.selectionEnd = ta.value.length;
  }, 0);
  return wrap;
}

/** One comment thread: the comment, its replies, and Reply / Resolve / Delete. */
function commentBalloon(view, c) {
  const { session } = view.odx;
  const active = session.activeComment === c.id;
  const b = h('div', `odx-balloon odx-balloon-comment${c.done ? ' odx-resolved' : ''}${active ? ' odx-active' : ''}`);
  b.onmousedown = (e) => { if (!active && e.target.tagName !== 'TEXTAREA') { e.preventDefault(); setActive(view, c.id); } };
  b.tabIndex = 0;
  b.setAttribute('role', 'group');
  b.setAttribute('aria-label', 'Comment');
  b.onkeydown = (e) => { if (e.key === 'Enter' && e.target === b && !active) { e.preventDefault(); setActive(view, c.id); } };
  const entry = (x) => {
    const e = h('div', 'odx-cmt-entry');
    const head = h('div', 'odx-cmt-head');
    head.append(h('span', 'odx-cmt-avatar', x.initials || (x.author || '?')[0]), h('b', '', x.author || 'Unknown'), h('small', '', when(x.date, x.dateUtc)));
    e.append(head);
    if (session.editingComment === x.id) {
      // (Kept on the session: the column can be drawn again, a new box in this one's place.)
      const findSel = { get: () => (session.commentSelect?.id === String(x.id) ? session.commentSelect : null), clear: () => { session.commentSelect = null; } };
      e.append(composer('e' + x.id, 'Add a comment', x.text, (t) => { session.editingComment = null; setText(view, x.id, t); announce('Comment posted'); back(view); },
        () => { session.editingComment = null; if (!x.text) deleteComment(view, x.id); else setActive(view, x.id); back(view); }, findSel));
    } else e.append(h('div', 'odx-cmt-text', x.text));
    if (x.rich) e.append(h('div', 'odx-cmt-note', 'Has a picture or table: edit it in Word'));
    return e;
  };
  b.append(entry(c));
  const replies = session.comments.filter((x) => x.parentId === c.id);
  if (c.done && !active) {
    b.append(h('div', 'odx-cmt-note', `Resolved${replies.length ? ` · ${replies.length} ${replies.length === 1 ? 'reply' : 'replies'}` : ''}`));
    return b;
  }
  for (const r of replies) b.append(entry(r));
  if (active && session.editingComment !== c.id) {
    if (session.replying === c.id) {
      b.append(composer('r' + c.id, 'Reply', '', (t) => { session.replying = null; replyTo(view, c.id, t); back(view); }, () => { session.replying = null; setActive(view, c.id); back(view); }));
    } else {
      const row = h('div', 'odx-cmt-actions');
      const btn = (label, fn) => { const x = h('button', '', label); x.type = 'button'; x.onmousedown = (e) => { e.preventDefault(); e.stopPropagation(); fn(); }; x.onclick = (e) => { if (e.detail === 0) fn(); /* Enter or Space */ }; row.append(x); };
      if (!c.done) btn('Reply', () => { session.replying = c.id; setActive(view, c.id); });
      btn(c.done ? 'Reopen' : 'Resolve', () => setResolved(view, c.id, !c.done));
      if (!c.rich) btn('Edit', () => { session.editingComment = c.id; setActive(view, c.id); });
      btn('Delete', () => deleteComment(view, c.id));
      b.append(row);
    }
  }
  return b;
}
