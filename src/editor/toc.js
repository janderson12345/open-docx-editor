// Table of contents: a TOC field listing the document's headings (outline
// levels 1-3) with their page numbers, as Word's "Automatic Table" inserts it.
// Word refreshes it the same way on "Update Table".
import { Node } from 'prosemirror-model';
import { TextSelection } from 'prosemirror-state';
import { schema } from './schema.js';
import { blocksFromXml } from '../ooxml/reader.js';
import { esc } from '../ooxml/xml.js';
import { outlineLevel } from './commands.js';
import { formatNumber } from '../ooxml/styles.js';

const INSTR = ' TOC \\o "1-3" \\h \\z \\u ';
const fld = (type) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const isInstr = (n, re) => n.type === schema.nodes.opaque_inline && n.attrs.kind === 'instrText' && re.test(n.attrs.xml);
const fldType = (n) => (n.type === schema.nodes.opaque_inline && n.attrs.kind === 'fldChar' ? (/fldCharType="(\w+)"/.exec(n.attrs.xml) || [])[1] : null);

/** Headings for the table: [{ level (0-based), text, pos }]. */
function headings(view, skip) {
  const { session } = view.odx;
  const out = [];
  view.state.doc.forEach((node, pos, i) => {
    if (skip && i >= skip[0] && i <= skip[1]) return;
    if (node.type !== schema.nodes.paragraph) return;
    const lvl = outlineLevel(session, node);
    const text = node.textContent.replace(/\s+/g, ' ').trim();
    if (lvl != null && lvl <= 2 && text) out.push({ level: lvl, text, pos });
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

function tocXml(view, skip, title = false) {
  const { session } = view.odx;
  const page = session.page();
  const width = Math.round((page.width - page.left - page.right) * 15);
  const has = (id) => session.styles.styles.has(id);
  const list = headings(view, skip);
  const heading = has('TOCHeading') ? '<w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>' : '<w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:r>';
  const paras = !title ? [] : [`<w:p>${heading}<w:r>${has('TOCHeading') ? '' : '<w:rPr><w:b/><w:sz w:val="28"/></w:rPr>'}<w:t>Contents</w:t></w:r></w:p>`];
  const begin = `${fld('begin')}<w:r><w:instrText xml:space="preserve">${esc(INSTR)}</w:instrText></w:r>${fld('separate')}`;
  if (!list.length) list.push({ level: 0, text: 'No table of contents entries found.', none: true });
  list.forEach((h, i) => {
    const style = `TOC${h.level + 1}`;
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
  let fresh = nodes(tocXml(view, range));
  if (sdt) fresh = fresh.map((n) => n.type.create({ ...n.attrs, sdt }, n.content, n.marks));
  view.dispatch(view.state.tr.replaceWith(from, to, fresh).setMeta('odx-untracked', true).setMeta('addToHistory', addToHistory));
  return true;
}
