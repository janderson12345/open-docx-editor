// Table of contents: a TOC field listing the document's headings (outline
// levels 1-3) with their page numbers, as Word's "Automatic Table" inserts it.
// Word refreshes it the same way on "Update Table".
import { Node } from 'prosemirror-model';
import { TextSelection } from 'prosemirror-state';
import { schema } from './schema.js';
import { blocksFromXml } from '../ooxml/reader.js';
import { esc, unescapeXml } from '../ooxml/xml.js';
import { outlineLevel } from './commands.js';
import { formatNumber } from '../ooxml/styles.js';

const INSTR = ' TOC \\o "1-3" \\h \\z \\u ';
const fld = (type) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const isInstr = (n, re) => n.type === schema.nodes.opaque_inline && n.attrs.kind === 'instrText' && re.test(n.attrs.xml);
const fldType = (n) => (n.type === schema.nodes.opaque_inline && n.attrs.kind === 'fldChar' ? (/fldCharType="(\w+)"/.exec(n.attrs.xml) || [])[1] : null);

/**
 * What a TOC field's switches ask for: `o` (paragraphs by their style's level), `u` (by the level set on the paragraph
 * itself) and the levels taken, 0-based. As Word builds them (odx-test toc-u-only.docx, made with Word COM and updated in
 * Word): a bare TOC or \o with no range is \o "1-9"; \u alone takes every paragraph with an outline level 1-9, direct or
 * from its style (a custom style's too), and no heading by style name alone; with \o, its range limits \u's entries too.
 * \t "Style,level,..." (toc-t-switch.docx, the same way): the named styles at those levels (1 when a name has none),
 * matched by name in any case, over the level \o would give them; \t without \o takes those styles only, \u adding none.
 */
export function tocSwitches(instr) {
  const o = /\\o(?:\s+"?\s*(\d)\s*-\s*(\d)\s*"?)?/i.exec(instr);
  const u = /\\u\b/i.test(instr);
  const t = /\\t\s+"([^"]*)"/i.exec(instr);
  const styles = new Map();
  if (t) {
    // (Word's list separator: a comma in English Word, a semicolon in some others.)
    const parts = t[1].split(/[,;]/).map((x) => x.trim());
    for (let i = 0; i < parts.length; i++) {
      const name = parts[i];
      if (!name || /^\d+$/.test(name)) continue;
      let lvl = 1;
      if (/^[1-9]$/.test(parts[i + 1] || '')) lvl = Number(parts[++i]);
      styles.set(name.toLowerCase(), lvl - 1);
    }
  }
  const range = o && o[1] ? [Number(o[1]) - 1, Number(o[2]) - 1] : [0, 8];
  return { o: !!o || (!u && !t), u: u && (!!o || !t), from: range[0], to: range[1], styles };
}

/** A paragraph's level for a table with these switches (0-based), or null when it isn't an entry. */
export function tocLevel(session, node, sw) {
  if (sw.styles?.size) {
    const name = session.styles.styles.get(node.attrs.styleId || session.styles.defaultParagraph)?.name;
    const named = name == null ? undefined : sw.styles.get(name.toLowerCase());
    if (named != null) return named;
  }
  if (!sw.o && !sw.u) return null;
  const direct = node.attrs.view?.outlineLvl;
  let lvl;
  if (sw.o && sw.u) lvl = outlineLevel(session, node);
  else if (sw.u) lvl = direct ?? session.styleProp(node.attrs.styleId, 'outlineLvl');
  else lvl = outlineLevel(session, { attrs: { ...node.attrs, view: { ...node.attrs.view, outlineLvl: undefined } } });
  return lvl == null || lvl < sw.from || lvl > sw.to ? null : lvl;
}

/** The instruction text of the TOC field in blocks [first, last] (it may be split over several runs), or null. */
function tocInstr(doc, range) {
  let text = null, depth = 0;
  for (let i = range[0]; i <= range[1]; i++) {
    let done = false;
    doc.child(i).forEach((c) => {
      if (done) return;
      if (text == null) { if (isInstr(c, /\bTOC\b/)) { text = ''; depth = 1; } else return; }
      const t = fldType(c);
      if (t === 'begin') depth++;
      else if (t === 'end') depth--;
      else if (t === 'separate' && depth === 1) done = true;
      else if (depth === 1 && isInstr(c, /./)) text += unescapeXml((/<w:instrText\b[^>]*>([^<]*)</.exec(c.attrs.xml) || [])[1] || '');
    });
    if (done) break;
  }
  return text;
}

/** Headings for the table: [{ level (0-based), text, pos }]. */
function headings(view, skip, sw) {
  const { session } = view.odx;
  const out = [];
  view.state.doc.forEach((node, pos, i) => {
    if (skip && i >= skip[0] && i <= skip[1]) return;
    if (node.type !== schema.nodes.paragraph) return;
    const lvl = tocLevel(session, node, sw);
    const text = node.textContent.replace(/\s+/g, ' ').trim();
    if (lvl != null && text) out.push({ level: lvl, text, pos });
  });
  return out;
}

/** The page number a heading is on, as its page shows it (a section may restart the numbering or use roman numerals). */
function pageOf(view, pos) {
  const { pages } = view.odx;
  try {
    const c = view.coordsAtPos(pos + 1);
    const { n, fmt } = pages.numberAt(pages.pageAt(c.top - view.dom.parentElement.getBoundingClientRect().top).p);
    return formatNumber(n, fmt);
  } catch { return '1'; }
}

function tocXml(view, skip, title = false, instr = INSTR) {
  const { session } = view.odx;
  const page = session.page();
  const width = Math.round((page.width - page.left - page.right) * 15);
  // (Built-in styles by the document's own ids: a localised Word's TOC 1 is "TDC1" in Spanish.)
  const id = (x) => session.builtInId(x);
  const list = headings(view, skip, tocSwitches(instr));
  // Word adds the TOC N styles a table's entries need (TOC 4-9 for a \u table's deeper levels).
  const need = [...new Set(list.map((h) => `TOC${h.level + 1}`))];
  if (skip && session.addBuiltInStyles(need)) view.odx.restyle?.();
  const has = (x) => session.styles.styles.has(id(x));
  const heading = has('TOCHeading') ? `<w:pPr><w:pStyle w:val="${id('TOCHeading')}"/></w:pPr>` : '<w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:r>';
  const paras = !title ? [] : [`<w:p>${heading}<w:r>${has('TOCHeading') ? '' : '<w:rPr><w:b/><w:sz w:val="28"/></w:rPr>'}<w:t>Contents</w:t></w:r></w:p>`];
  const begin = `${fld('begin')}<w:r><w:instrText xml:space="preserve">${esc(instr)}</w:instrText></w:r>${fld('separate')}`;
  if (!list.length) list.push({ level: 0, text: 'No table of contents entries found.', none: true });
  list.forEach((h, i) => {
    const style = id(`TOC${h.level + 1}`);
    const ppr = `<w:pPr>${has(style) ? `<w:pStyle w:val="${style}"/>` : ''}<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${width}"/></w:tabs>${has(style) ? '' : `<w:spacing w:after="60"/><w:ind w:left="${h.level * 220}"/>`}</w:pPr>`;
    const num = h.none ? '' : `<w:r><w:tab/></w:r><w:r><w:t>${pageOf(view, h.pos)}</w:t></w:r>`;
    paras.push(`<w:p>${ppr}${i === 0 ? begin : ''}<w:r><w:t xml:space="preserve">${esc(h.text)}</w:t></w:r>${num}${i === list.length - 1 ? fld('end') : ''}</w:p>`);
  });
  return paras.join('');
}

/** Index range [first, last] of the top-level blocks holding the TOC field, or null. */
export function findToc(doc) {
  let start = -1, end = -1, depth = 0, inToc = false;
  doc.forEach((node, pos, i) => {
    if (end >= 0 || node.type !== schema.nodes.paragraph) return;
    node.forEach((c) => {
      if (end >= 0) return;
      if (!inToc && isInstr(c, /\bTOC\b/)) { inToc = true; start = i; depth = 1; return; }
      if (!inToc) return;
      const t = fldType(c);
      if (t === 'begin') depth++;
      else if (t === 'end' && --depth === 0) end = i;
    });
  });
  if (start < 0) return null;
  return [start, end < 0 ? start : end];
}

const nodes = (xml) => blocksFromXml(xml).map((j) => Node.fromJSON(schema, j));

/** Insert a table of contents at the cursor (as its own blocks). */
export function insertToc(view) {
  // Word's TOC Heading and TOC 1-3 styles, added to the document if it lacks them (their CSS with them).
  if (view.odx.session.addBuiltInStyles(['TOCHeading', 'TOC1', 'TOC2', 'TOC3'])) view.odx.restyle?.();
  const { state } = view;
  const $from = state.selection.$from;
  const at = $from.depth ? $from.before(1) : 0;
  const tr = state.tr.insert(at, nodes(tocXml(view, null, true)));
  tr.setSelection(TextSelection.near(tr.doc.resolve(at + 1))).setMeta('odx-untracked', true);
  view.dispatch(tr.scrollIntoView());
  // Page numbers shift once the table itself takes up room: update after layout.
  // That update isn't a separate step for Undo: one Ctrl+Z removes the whole table, as in Word.
  setTimeout(() => updateToc(view, { addToHistory: false }), 700);
  return true;
}

/** Rebuild the table of contents with the current headings and page numbers. */
export function updateToc(view, { addToHistory = true } = {}) {
  const range = findToc(view.state.doc);
  if (!range) return false;
  const doc = view.state.doc;
  let from = 0;
  for (let i = 0; i < range[0]; i++) from += doc.child(i).nodeSize;
  let to = from;
  for (let i = range[0]; i <= range[1]; i++) to += doc.child(i).nodeSize;
  const sdt = doc.child(range[0]).attrs.sdt;
  // (The field keeps its own switches: rebuilt from what they ask for, not the table Insert makes.)
  let fresh = nodes(tocXml(view, range, false, tocInstr(doc, range) || INSTR));
  if (sdt) fresh = fresh.map((n) => n.type.create({ ...n.attrs, sdt }, n.content, n.marks));
  view.dispatch(view.state.tr.replaceWith(from, to, fresh).setMeta('odx-untracked', true).setMeta('addToHistory', addToHistory));
  return true;
}
