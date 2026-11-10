// A .docx for an AI agent (or any program) to read and edit without a browser: its paragraphs with stable ids, its
// outline, a Markdown view, search, and edits in plain terms (replace text, rewrite a paragraph, insert or delete one,
// comment, fill form fields), each made as a tracked change under the agent's own name unless asked otherwise, so a
// person reviews them in Word or the editor and accepts or rejects them. Saving keeps everything else in the file as
// it was. In Node, call setXmlImpl(...) with a DOMParser/XMLSerializer first (as for DocxSession).
import { EditorState } from 'prosemirror-state';
import { DocxSession } from './ooxml/docx.js';
import { listCounter } from './ooxml/styles.js';
import { schema } from './editor/schema.js';
import { trackTransaction, resolveChanges, paragraphMarkRev } from './editor/track.js';
import { findAll, replacementFor } from './editor/find.js';
import { commentMarkers, newComment, marker } from './editor/comments.js';
import { fieldValues, listFields, setFieldValues } from './editor/fields.js';

/** Opens a .docx (bytes) for reading and editing. `author`: the name its edits and comments go under ('AI'). */
export function openDocument(bytes, { author = 'AI' } = {}) {
  return new AgentDocument(bytes, author);
}

// Paragraph ids ('p0', 'p1'...) are given in order when the document is opened and follow their paragraph through
// every edit after that, so an id read once still names the same paragraph after others are inserted or deleted. A
// new paragraph gets the next unused id.

/**
 * The same, on the document open in an editor (mountDocxEditor's `api.agent()`): its edits go into the editor as they're
 * made, tracked under the agent's name, for the person at the editor to see, accept or reject.
 */
export function editorAgent(view, { author = 'AI' } = {}) {
  return new AgentDocument(null, author, view);
}

export class AgentDocument {
  constructor(bytes, author, view = null) {
    this.view = view;
    this.session = view ? view.odx.session : new DocxSession(bytes);
    this.author = author;
    this._state = view ? null : EditorState.create({ doc: this.session.doc });
    this.idAt = new Map(); // paragraph position -> its id
    this.nextId = 0;
  }

  /** The editor state as it is now (the editor's own, for an editor's agent). */
  get state() { return this.view ? this.view.state : this._state; }

  /** The editor document as it is now. */
  get doc() { return this.state.doc; }

  // --- Reading --------------------------------------------------------------------------------------------------

  /**
   * Every paragraph, in order: [{ id, style, heading (outline level 1-9, or null), list (its number or bullet, as Word
   * shows it), table ({ table, row, cell } when in a table cell), text }]. Text is the document as it would read with
   * its tracked changes accepted.
   */
  paragraphs() {
    const out = [];
    const ids = this.paragraphRanges();
    const counter = listCounter(this.session.numbering);
    let table = -1;
    const walk = (node, cellAt) => {
      node.forEach((child, _, i) => {
        if (child.type === schema.nodes.table) {
          table++;
          const t = table;
          child.forEach((row, __, r) => row.forEach((cell, ___, c) => walk(cell, { table: t, row: r, cell: c })));
        } else if (child.type === schema.nodes.paragraph) {
          const numId = child.attrs.numId ?? this.session.styleProp(child.attrs.styleId, 'numId');
          const ilvl = child.attrs.ilvl ?? this.session.styleProp(child.attrs.styleId, 'ilvl') ?? 0;
          const list = numId && numId !== '0' ? counter(numId, Number(ilvl)) : null;
          out.push({ id: ids[out.length].id, style: this.styleName(child.attrs.styleId), heading: this.headingLevel(child), list: list || null, table: cellAt || null, text: finalText(child) });
        } else if (!child.isLeaf && !child.isTextblock) walk(child, cellAt);
        void i;
      });
    };
    walk(this.doc, null);
    return out;
  }

  /** The headings: [{ id, level (1-9), text }]. */
  outline() {
    return this.paragraphs().filter((p) => p.heading && p.text.trim()).map(({ id, heading, text }) => ({ id, level: heading, text }));
  }

  /**
   * The document as Markdown, for reading: headings, list items, tables, and each paragraph's id in a comment before
   * it (`<!-- p12 -->`) so edits can name it. `from` and `to` (ids) give part of it.
   */
  markdown({ from, to, ids = true } = {}) {
    const list = this.paragraphs();
    const at = (id) => { const k = list.findIndex((p) => p.id === id); if (k < 0) throw new Error(`No paragraph ${id}.`); return k; };
    const a = from ? at(from) : 0, b = to ? at(to) : list.length - 1;
    const lines = [];
    let rows = null;
    const flush = () => {
      if (!rows) return;
      const width = Math.max(...rows.map((r) => r.length));
      rows.forEach((r, i) => {
        lines.push(`| ${Array.from({ length: width }, (_, k) => (r[k] || '').replace(/\|/g, '\\|')).join(' | ')} |`);
        if (i === 0) lines.push(`|${' --- |'.repeat(width)}`);
      });
      lines.push('');
      rows = null;
    };
    let tableAt = null;
    for (const p of list.slice(a, b + 1)) {
      if (p.table) {
        if (!rows || tableAt !== p.table.table) { flush(); rows = []; tableAt = p.table.table; }
        const row = (rows[p.table.row] ||= []);
        row[p.table.cell] = [row[p.table.cell], `${ids ? `(${p.id}) ` : ''}${p.text}`].filter(Boolean).join(' / ');
        continue;
      }
      flush();
      if (!p.text.trim() && !p.list) continue;
      if (ids) lines.push(`<!-- ${p.id} -->`);
      if (p.heading) lines.push(`${'#'.repeat(Math.min(6, p.heading))} ${p.list ? `${p.list} ` : ''}${p.text}`);
      else if (p.list) lines.push(`- ${/^[•●○▪■✓•-]/.test(p.list) ? '' : `${p.list} `}${p.text}`);
      else lines.push(p.text);
      lines.push('');
    }
    flush();
    return lines.join('\n').trim() + '\n';
  }

  /** Where text appears: [{ id, text (the paragraph's), match, start, end (in the paragraph's text) }]. */
  search(query, { regex = false, matchCase = false, wholeWord = false } = {}) {
    const ranges = this.paragraphRanges();
    return findAll(this.doc, query, { regex, matchCase, wholeWord }).map((m) => {
      const i = ranges.findIndex((r) => m.from >= r.from && m.to <= r.to);
      if (i < 0) return null;
      const r = ranges[i];
      return { id: r.id, text: this.doc.textBetween(r.from, r.to, '', '￼'), match: this.doc.textBetween(m.from, m.to, '', '￼'), start: m.from - r.from, end: m.to - r.from };
    }).filter(Boolean);
  }

  /** The comments: [{ id, author, date, text, quote (the commented text), replyTo, resolved }]. */
  comments() {
    const markers = commentMarkers(this.doc);
    return this.session.comments.map((c) => {
      const m = markers.get(String(c.id)) || {};
      const quote = m.start != null && m.end != null ? this.doc.textBetween(m.start, m.end, ' ', '') : '';
      return { id: String(c.id), author: c.author, date: c.date, text: c.text, quote, replyTo: c.parentId ?? null, resolved: !!c.done };
    });
  }

  /** The tracked changes: [{ kind: 'insertion' | 'deletion', author, date, text, id (the paragraph's) }]. */
  changes() {
    const out = [];
    const ranges = this.paragraphRanges();
    this.doc.descendants((n, pos) => {
      if (!n.isText) return true;
      for (const m of n.marks) {
        if (m.type !== schema.nodes.paragraph.schema.marks.ins && m.type !== schema.marks.del) continue;
        const i = ranges.findIndex((r) => pos >= r.from && pos <= r.to);
        const kind = m.type === schema.marks.ins ? 'insertion' : 'deletion';
        const last = out[out.length - 1];
        const id = ranges[i]?.id ?? null;
        if (last && last.kind === kind && last.author === m.attrs.author && last.id === id && last.end === pos) { last.text += n.text; last.end = pos + n.nodeSize; continue; }
        out.push({ kind, author: m.attrs.author || '', date: m.attrs.date || '', text: n.text, id, end: pos + n.nodeSize });
      }
      return false;
    });
    return out.map(({ end, ...c }) => c);
  }

  /** Form fields (content controls): { name: value }, and their details with `details: true`. */
  fields({ details = false } = {}) {
    return details ? listFields(this.doc).map(({ tag, alias, kind, value, checked, items }) => ({ tag, alias, kind, value, checked, items })) : fieldValues(this.doc);
  }

  // --- Editing (tracked by default) ----------------------------------------------------------------------------

  /**
   * Replaces text: the first match (in paragraph `id` if given), or every one with `all`. With `regex`, $1, $2... in
   * the replacement are what the pattern's groups caught. Returns how many were replaced.
   */
  replace(query, replacement, { id, all = false, regex = false, matchCase = true, wholeWord = false, tracked = true } = {}) {
    let matches = findAll(this.doc, query, { regex, matchCase, wholeWord });
    if (id) { const r = this.range(id); matches = matches.filter((m) => m.from >= r.from && m.to <= r.to); }
    // (Not text already deleted by a tracked change.)
    matches = matches.filter((m) => !this.deletedAt(m.from, m.to));
    if (!all) matches = matches.slice(0, 1);
    if (!matches.length) return 0;
    const tr = this.state.tr;
    for (const m of [...matches].reverse()) tr.insertText(replacementFor(m, replacement), m.from, m.to);
    this.apply(tr, tracked);
    return matches.length;
  }

  /** Rewrites paragraph `id`'s text (keeping its style and its first run's formatting). */
  setParagraphText(id, text, { tracked = true } = {}) {
    const r = this.range(id);
    const old = this.doc.textBetween(r.from, r.to, '', '￼');
    if (old === text) return false;
    // Only what differs is changed: the common start and end are kept, so a tracked change shows just the edit.
    let s = 0;
    while (s < old.length && s < text.length && old[s] === text[s]) s++;
    let e = 0;
    while (e < old.length - s && e < text.length - s && old[old.length - 1 - e] === text[text.length - 1 - e]) e++;
    const tr = this.state.tr;
    tr.insertText(text.slice(s, text.length - e), this.textPos(r, s), this.textPos(r, old.length - e));
    this.apply(tr, tracked);
    return true;
  }

  /** Inserts a paragraph after paragraph `after` (or at the start with null), in style `style` (a name or id). Returns its id. */
  insertParagraph(after, text, { style = null, tracked = true } = {}) {
    let at = 0;
    let styleId = null;
    if (after != null) {
      const r = this.range(after);
      at = r.to + 1;
      styleId = this.doc.resolve(r.from).parent.attrs.styleId;
    }
    if (style) styleId = this.styleId(style);
    const node = schema.nodes.paragraph.create({ styleId }, text ? schema.text(text) : null);
    const tr = this.state.tr.insert(at, node);
    this.apply(tr, tracked);
    return this.paragraphRanges().find((r) => r.from === at + 1)?.id ?? null;
  }

  /** Deletes paragraph `id` (as a tracked deletion unless `tracked: false`). */
  deleteParagraph(id, { tracked = true } = {}) {
    const r = this.range(id);
    const $p = this.doc.resolve(r.from);
    const tr = this.state.tr.delete($p.before(), $p.after());
    this.apply(tr, tracked);
    return true;
  }

  /** Adds a comment on paragraph `id`: on `quote` within it if given (its first occurrence), else on the whole paragraph. */
  addComment(id, text, { quote = null, author = this.author } = {}) {
    const r = this.range(id);
    let from = r.from, to = r.to;
    if (quote) {
      const hit = this.search(quote, { matchCase: true }).find((h) => h.id === id);
      if (!hit) throw new Error(`"${quote}" isn't in ${id}.`);
      from = this.textPos(r, hit.start); to = this.textPos(r, hit.end);
    }
    const c = newComment(this.session, this.doc, author);
    c.text = text;
    const tr = this.state.tr;
    tr.insert(to, [marker('commentRangeEnd', c.id), marker('commentReference', c.id)]);
    tr.insert(from, marker('commentRangeStart', c.id));
    this.commit(tr.setMeta('odx-untracked', true));
    return String(c.id);
  }

  /** Replies to comment `commentId`. */
  reply(commentId, text, { author = this.author } = {}) {
    const parent = this.session.comments.find((c) => String(c.id) === String(commentId));
    if (!parent) throw new Error(`No comment ${commentId}.`);
    const m = commentMarkers(this.doc).get(String(commentId)) || {};
    const c = newComment(this.session, this.doc, author, { parentId: parent.id });
    c.text = text;
    if (m.start != null && m.ref != null) {
      const tr = this.state.tr;
      tr.insert(m.ref + 1, marker('commentReference', c.id));
      tr.insert(m.end + 1, marker('commentRangeEnd', c.id));
      tr.insert(m.start, marker('commentRangeStart', c.id));
      this.commit(tr.setMeta('odx-untracked', true));
    }
    return String(c.id);
  }

  /** Fills form fields by Tag or Title: { CustomerName: 'Acme', Agree: true }. Returns the names that matched none. */
  setFields(values) {
    const view = { get state() { return self.state; }, dispatch: (tr) => self.commit(tr) };
    const self = this;
    return setFieldValues(view, values);
  }

  /** Accepts (or with `accept: false` rejects) every tracked change, or those in paragraph `id`. */
  resolveChanges({ accept = true, id } = {}) {
    const r = id ? this.range(id) : { from: 0, to: this.doc.content.size };
    const tr = this.state.tr;
    resolveChanges(tr, r.from, r.to, accept);
    this.commit(tr);
  }

  /** The .docx with the edits (bytes). */
  save() {
    this.view?.odx.flushStories?.();
    return this.session.save(this.doc);
  }

  // --- Inner workings -------------------------------------------------------------------------------------------

  /** Applies an edit, as a tracked change under the agent's name unless `tracked` is false. */
  apply(tr, tracked) {
    this.commit(tracked ? trackTransaction(this.state, tr, this.author) : tr.setMeta('odx-untracked', true));
  }

  /** A transaction applied, the paragraphs' ids carried to their new places (a paragraph deleted outright loses its). */
  commit(tr) {
    this.paragraphRanges(); // (every paragraph has its id before the change)
    const moved = new Map();
    for (const [pos, id] of this.idAt) {
      const r = tr.mapping.mapResult(pos, 1);
      if (!r.deleted) moved.set(r.pos, id);
    }
    this.idAt = moved;
    // (In an editor: already tracked under the agent's name, so the editor doesn't track it again under its own.)
    if (this.view) { this.view.dispatch(tr.setMeta('odx-untracked', true)); this.view.odx.pages?.balloons?.(); }
    else this._state = this._state.apply(tr);
  }

  /** Each paragraph's id and content range: [{ id, from, to }], in the order of paragraphs(). */
  paragraphRanges() {
    const out = [];
    const ids = new Map();
    this.doc.descendants((n, pos) => {
      if (n.type !== schema.nodes.paragraph) return true;
      let id = this.idAt.get(pos);
      if (!id) id = `p${this.nextId++}`;
      ids.set(pos, id);
      out.push({ id, from: pos + 1, to: pos + 1 + n.content.size });
      return false;
    });
    this.idAt = ids;
    return out;
  }

  range(id) {
    const r = this.paragraphRanges().find((x) => x.id === id);
    if (!r) throw new Error(`No paragraph ${id}.`);
    return r;
  }

  /** The document position of character `k` of a paragraph's text (each inline object counts as one). */
  textPos(r, k) {
    let left = k, at = r.from;
    this.doc.nodesBetween(r.from, r.to, (n, pos) => {
      if (left <= 0 || n.type === schema.nodes.paragraph) return left > 0;
      if (n.isText) { const take = Math.min(left, n.text.length); if (take === left) { at = pos + take; left = 0; return false; } left -= n.text.length; at = pos + n.nodeSize; }
      else if (n.isLeaf) { left -= 1; at = pos + n.nodeSize; }
      return false;
    });
    return at;
  }

  deletedAt(from, to) {
    let del = false;
    this.doc.nodesBetween(from, to, (n) => { if (n.isText && n.marks.some((m) => m.type === schema.marks.del)) del = true; });
    return del;
  }

  styleName(id) {
    return (id && this.session.styles.styles.get(id)?.name) || (id ? id : 'Normal');
  }

  styleId(name) {
    const want = name.toLowerCase();
    for (const [id, st] of this.session.styles.styles) if (id.toLowerCase() === want || (st.name || '').toLowerCase() === want) return id;
    throw new Error(`No style "${name}".`);
  }

  headingLevel(node) {
    const lvl = node.attrs.view?.outlineLvl ?? this.session.styleProp(node.attrs.styleId, 'outlineLvl');
    if (lvl != null) return lvl > 8 ? null : lvl + 1;
    for (let id = node.attrs.styleId, n = 0; id && n < 10; n++) {
      const st = this.session.styles.styles.get(id);
      const m = st && /^heading\s*(\d)$/i.exec(st.name || '');
      if (m) return Number(m[1]);
      id = st && st.basedOn;
    }
    return null;
  }
}

/** A paragraph's text as it reads with its tracked changes accepted (deleted text left out). */
function finalText(node) {
  let t = '';
  node.descendants((n) => {
    if (n.isText) { if (!n.marks.some((m) => m.type === schema.marks.del)) t += n.text; }
    else if (n.type.name === 'tab') t += '\t';
    else if (n.type.name === 'hard_break') t += '\n';
    return true;
  });
  return paragraphMarkRev(node.attrs) === 'del' && !t ? '' : t;
}
