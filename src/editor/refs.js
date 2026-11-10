// Reference fields: PAGEREF (the page a bookmark is on), REF (a bookmark's text, or its paragraph's number) and
// STYLEREF (the text of the nearest paragraph in a style: a running head), worked out from the laid-out pages as Word
// works them out. In headers and footers Word does it for every page as it draws it (pagefields.js asks here); in the
// body it keeps the saved results until the fields are updated (F9: updateFields below).
import { schema } from './schema.js';
import { formatNumber } from '../ooxml/styles.js';
import { announce } from './live.js';

const NOT_FOUND = 'Error! Reference source not found.';
const NO_STYLE = 'Error! No text of specified style in document.';
const MISSING = [];
const FORMATS = { roman: 'lowerRoman', Roman: 'upperRoman', ROMAN: 'upperRoman', alphabetic: 'lowerLetter', ALPHABETIC: 'upperLetter', Arabic: null, arabic: null };

/** A field instruction's name and first argument: `PAGEREF _Toc123 \h` -> { name: 'PAGEREF', arg: '_Toc123' }. */
export function parseField(instr) {
  const m = /^\s*([A-Za-z]+)(?:\s+(?:"([^"]*)"|([^\s\\]+)))?/.exec(instr || '');
  if (!m) return null;
  return { name: m[1].toUpperCase(), word: m[1], arg: m[2] ?? m[3] ?? null, sw: (k) => new RegExp(`\\\\${k}\\b`, 'i').test(instr), fmt: /\\\*\s*(\w+)/.exec(instr)?.[1] };
}

/** The document's bookmarks: name (lower case) -> { from, to } (document positions of their start and end). */
export function bookmarks(doc) {
  const starts = new Map(); // id -> { name, from }
  const out = new Map();
  doc.descendants((n, pos) => {
    if (n.type !== schema.nodes.opaque_inline && n.type !== schema.nodes.opaque_block) return true;
    const xml = n.attrs.xml || '';
    if (n.attrs.kind === 'bookmarkStart') {
      const id = /w:id="([^"]*)"/.exec(xml)?.[1], name = /w:name="([^"]*)"/.exec(xml)?.[1];
      if (id != null && name) starts.set(id, { name, from: pos + n.nodeSize });
    } else if (n.attrs.kind === 'bookmarkEnd') {
      const s = starts.get(/w:id="([^"]*)"/.exec(xml)?.[1]);
      if (s && !out.has(s.name.toLowerCase())) out.set(s.name.toLowerCase(), { from: s.from, to: pos });
    }
    return false;
  });
  for (const s of starts.values()) if (!out.has(s.name.toLowerCase())) out.set(s.name.toLowerCase(), { from: s.from, to: s.from });
  return out;
}

/** The page (0-based index) document position `pos` is laid out on, or null. */
function pageIndexAt(view, pos) {
  try {
    const c = view.coordsAtPos(Math.min(pos, view.state.doc.content.size));
    return view.odx.pages.pageAt(c.top - view.dom.parentElement.getBoundingClientRect().top).p;
  } catch { return null; }
}

/** A page's number as its header shows it: the section's numbering and format, or the field's own (\* roman). */
function pageNumberText(view, p, fmt) {
  const { n, fmt: own } = view.odx.pages.numberAt(p);
  return formatNumber(n, fmt && fmt in FORMATS ? FORMATS[fmt] : own);
}

/** The list label of the paragraph holding `pos` ("3.2." -> "3.2"), as drawn; '' if it has none. */
function paragraphNumber(view, pos) {
  const $p = view.state.doc.resolve(pos);
  for (let d = $p.depth; d > 0; d--) {
    if ($p.node(d).type !== schema.nodes.paragraph) continue;
    const el = view.nodeDOM($p.before(d));
    return (el?.dataset?.label || '').replace(/[\s.)]+$/, '').trim();
  }
  return '';
}

/** The style's paragraphs (by its name or id, any case): [{ pos, text, p }], with the page each is on. */
function styledParagraphs(view, style) {
  const { session } = view.odx;
  const want = style.toLowerCase();
  const ids = new Set();
  for (const [id, st] of session.styles.styles) if (id.toLowerCase() === want || (st.name || '').toLowerCase() === want) ids.add(id);
  // (Not defined in styles.xml, but one of Word's built-in styles (w:latentStyles names them): it exists, unused.)
  session._latent ??= new Set([...(session.pkg.text('word/styles.xml') || '').matchAll(/<w:lsdException\b[^>]*w:name="([^"]*)"/g)].map((m) => m[1].toLowerCase()));
  if (!ids.size && !session._latent.has(want) && !/^heading [1-9]$/.test(want)) return MISSING;
  const out = [];
  view.state.doc.descendants((n, pos) => {
    if (n.type !== schema.nodes.paragraph) return true;
    if (ids.has(n.attrs.styleId)) {
      const text = n.textContent.replace(/\s+/g, ' ').trim();
      if (text) out.push({ pos, text });
      return false;
    }
    // A character style: each stretch of text in it (Word's STYLEREF finds those too).
    let run = null;
    n.forEach((c, off) => {
      const on = c.isText && c.marks.some((m) => m.type.name === 'cstyle' && ids.has(m.attrs.val));
      if (on && run) run.text += c.text;
      else if (on) out.push((run = { pos: pos + off, text: c.text }));
      else if (c.isText || c.isBlock) run = null;
    });
    return false;
  });
  for (const o of out) o.text = o.text.replace(/\s+/g, ' ').trim();
  for (const o of out) o.p = pageIndexAt(view, o.pos + 1);
  return out;
}

/**
 * A reference field's value, or undefined for other fields. `at`: where it's worked out: { page } for a header or footer
 * (that page), { pos } for the body (the field's own place).
 */
export function referenceValue(view, instr, at) {
  const f = parseField(instr);
  if (!f) return undefined;
  const doc = view.state.doc;
  // A field that's just a bookmark's name is a REF to it.
  const name = f.name === 'REF' || f.name === 'PAGEREF' || f.name === 'STYLEREF' ? f.name : bookmarks(doc).has(f.word.toLowerCase()) ? 'REF' : null;
  if (!name) return undefined;
  const arg = name === 'REF' && f.name !== 'REF' ? f.word : f.arg;
  if (name === 'STYLEREF') return styleRef(view, arg || '', at, f);
  const bm = arg && bookmarks(doc).get(arg.toLowerCase());
  if (!bm) return NOT_FOUND;
  // \p: where the bookmark is from the field: "above" or "below" on the same page, else "on page N" (Word's words).
  if (f.sw('p') && at.pos != null) {
    const there = pageIndexAt(view, bm.from), here = pageIndexAt(view, at.pos);
    if (there != null && there !== here) return `on page ${pageNumberText(view, there, f.fmt)}`;
    return bm.from < at.pos ? 'above' : 'below';
  }
  if (name === 'PAGEREF') {
    const p = pageIndexAt(view, bm.from);
    return p == null ? NOT_FOUND : pageNumberText(view, p, f.fmt);
  }
  // REF: the paragraph's number (\n, \r, \w), else the bookmarked text.
  if (f.sw('n') || f.sw('r') || f.sw('w')) return paragraphNumber(view, bm.from) || '0';
  return doc.textBetween(bm.from, Math.max(bm.from, bm.to), ' ', '');
}

/**
 * STYLEREF on a page, as Word finds it for a header or footer: the first paragraph in the style on that page (\l: the
 * last), else the nearest one before the page, else the first after it. In the body: the nearest before the field.
 */
function styleRef(view, style, at, f) {
  const list = style ? styledParagraphs(view, style) : null;
  // (A style the document doesn't have: Word says to apply it.)
  if (!list) return NO_STYLE;
  if (list === MISSING) return `Error! Use the Home tab to apply ${style} to the text that you want to appear here.`;
  if (!list.length) return NO_STYLE;
  let hit;
  if (at.page != null) {
    const on = list.filter((x) => x.p === at.page);
    hit = on.length ? (f.sw('l') ? on[on.length - 1] : on[0]) : [...list].reverse().find((x) => x.p != null && x.p < at.page) || list.find((x) => x.p != null && x.p > at.page);
  } else {
    hit = [...list].reverse().find((x) => x.pos < at.pos) || list[0];
  }
  if (!hit) return NO_STYLE;
  return f.sw('n') || f.sw('r') || f.sw('w') ? paragraphNumber(view, hit.pos + 1) || hit.text : hit.text;
}

const fldType = (n) => (n.type === schema.nodes.opaque_inline && n.attrs.kind === 'fldChar' ? /fldCharType="(\w+)"/.exec(n.attrs.xml)?.[1] : null);
const instrOf = (n) => (n.type === schema.nodes.opaque_inline && n.attrs.kind === 'instrText' ? (/<w:instrText[^>]*>([^<]*)</.exec(n.attrs.xml) || [])[1] || '' : null);
const unescape = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/**
 * The body's reference and page fields in [from, to]: [{ instr, resultFrom, resultTo }] (the result's document range,
 * between the field's separator and its end).
 */
function bodyFields(doc, from, to) {
  const out = [];
  const stack = [];
  doc.nodesBetween(from, to, (n, pos) => {
    const t = fldType(n);
    if (t === 'begin') stack.push({ instr: '', resultFrom: null });
    else if (t === 'separate' && stack.length) stack[stack.length - 1].resultFrom = pos + n.nodeSize;
    else if (t === 'end' && stack.length) {
      const f = stack.pop();
      // (Fields inside others too: a table of contents' page numbers are PAGEREF fields within its TOC field.)
      if (f.resultFrom != null) out.push({ instr: unescape(f.instr), resultFrom: f.resultFrom, resultTo: pos });
    } else if (stack.length && stack[stack.length - 1].resultFrom == null) {
      const i = instrOf(n);
      if (i != null) stack[stack.length - 1].instr += i;
    }
    return true;
  });
  return out;
}

/**
 * Word's Update Fields (F9, and what Word does as it prints): the reference fields (PAGEREF, REF, STYLEREF) and page
 * fields (PAGE, NUMPAGES) in the selection, or everywhere with `all` or nothing selected, get their values from the
 * pages as laid out now. A table of contents gets its page numbers (its PAGEREF fields), as Word's "Update page
 * numbers only"; References > Update table rebuilds it. One undoable step. Returns how many fields changed.
 */
export function updateFields(view, { all = false } = {}) {
  const { state } = view;
  const sel = state.selection;
  const [from, to] = all || sel.empty ? [0, state.doc.content.size] : [sel.from, sel.to];
  const doc = state.doc;
  const tr = view.state.tr;
  let count = 0;
  for (const fld of bodyFields(doc, from, to).sort((x, y) => y.resultFrom - x.resultFrom)) {
    const f = parseField(fld.instr);
    if (!f) continue;
    let value = referenceValue(view, fld.instr, { pos: fld.resultFrom });
    if (value === undefined && (f.name === 'PAGE' || f.name === 'NUMPAGES')) {
      const p = pageIndexAt(view, fld.resultFrom);
      value = f.name === 'PAGE' ? (p == null ? undefined : pageNumberText(view, p, f.fmt)) : String(view.odx.pages.count);
    }
    if (value === undefined) continue;
    const old = doc.textBetween(fld.resultFrom, fld.resultTo, '', '');
    if (old === value) continue;
    // The new text takes the formatting of the old result's first character (\* MERGEFORMAT), as Word does.
    let marks = null;
    doc.nodesBetween(fld.resultFrom, fld.resultTo, (n) => { if (marks == null && n.isText) marks = n.marks; return marks == null; });
    tr.replaceWith(tr.mapping.map(fld.resultFrom), tr.mapping.map(fld.resultTo), value ? schema.text(value, marks || []) : []);
    count++;
  }
  if (count) view.dispatch(tr.setMeta('odx-untracked', true).setMeta('odx-fields', true));
  return count;
}

/** Update Fields as a command (F9, References > Update fields): says how many were updated (to screen readers too). */
export function updateFieldsCommand(view, opts) {
  const n = updateFields(view.odx?.main || view, opts);
  announce(n ? `${n} field${n === 1 ? '' : 's'} updated` : 'No fields to update');
  return true;
}
