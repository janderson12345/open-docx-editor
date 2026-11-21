// The editor: a ProseMirror view laid out as a Word page, with Word-style keys
// and list numbering computed from numbering.xml.
import { commentsPlugin, addComment, refresh as refreshComments } from './comments.js';
import { hfDoubleClick, closeStory, flushStory } from './hfedit.js';
import { formsPlugin } from './fields.js';
import { bindingsPlugin } from './bindings.js';
import { noteClicks, closeNote, flushNote, insertNote } from './notes.js';
import { EditorState, Plugin, Selection, TextSelection } from 'prosemirror-state';
import { EditorView, Decoration, DecorationSet } from 'prosemirror-view';
import { Fragment, Slice } from 'prosemirror-model';
import { keymap } from 'prosemirror-keymap';
import { baseKeymap, chainCommands } from 'prosemirror-commands';
import { history, undo, redo } from 'prosemirror-history';
import { autocorrectPlugin } from './autocorrect.js';
import { updateFieldsCommand } from './refs.js';
import { tableEditing, columnResizing, goToNextCell, TableMap, TableView } from 'prosemirror-tables';
import { listCounter, resolveStyle, tableCellCss, tableStylesCss, tablePlacementCss } from '../ooxml/styles.js';
import { parseXml } from '../ooxml/xml.js';
import { ZERO_WIDTH } from '../ooxml/reader.js';
import { borderCss } from '../ooxml/props.js';
import { schema } from './schema.js';
import { uniformRun } from './runs.js';
import { columnsIn, sectionJoinPlugin } from './sections.js';
import { frameOf, frameStyle } from './frames.js';
import { listLabel, lineSize, pictureLine, labelLine, labelFormat, changeClasses, headingRole } from './pararules.js';
import * as cmd from './commands.js';
import { announce } from './live.js';

// A formatting shortcut says what it did, as Word does to a screen reader ("Bold on"): NVDA said nothing for Ctrl+B.
const spoken = (name, label) => (state, dispatch, view) => {
  const done = cmd.toggle(name)(state, dispatch, view);
  if (done && dispatch && view) announce(`${label} ${cmd.markActive(view.state, name) ? 'on' : 'off'}`);
  return done;
};
import { createPages, pagePushPlugin } from './pages.js';
import { justifyPlugin } from './justify.js';
import { wrapBothPlugin } from './wrapboth.js';
import { tableClearPlugin } from './tableclear.js';
import { hyphenatePlugin } from './hyphenate.js';
import { mailMergePlugin } from './mailmerge.js';
import { imageResizePlugin, shapeKeysPlugin } from './imgresize.js';
import { textBoxPlugin, closeTextBox } from './txbxedit.js';
import { chartPlugin } from './charts.js';
import { lineFixPlugin } from './lhfix.js';
import { textSpansPlugin } from './textspans.js';
import { widthsPlugin } from './widths.js';
import { fitTextPlugin } from './fittext.js';
import { createZoom } from './zoom.js';
import { joinsPlugin } from './joins.js';
import { marksPlugin, toggleMarks } from './marks.js';
import { csFormatPlugin } from './csformat.js';
import { pastePictures, pictureLinks } from './pastepics.js';
import { trackPlugin, trackTransaction, tracking, trackedDelete } from './track.js';
import { twipsToPx } from '../ooxml/units.js';
import { protectPlugin, tabLeaves } from './protect.js';
import { tableArrow, tableVert } from './tablecaret.js';
import { keepFitted } from './tablefit.js';

let scopeSeq = 0;

/** The resizable table view, plus the table style name the stylesheet keys on. */
class StyledTableView extends TableView {
  constructor(node, cellMinWidth, view) {
    super(node, cellMinWidth, view);
    this.tagStyle(node);
  }
  update(node) {
    const ok = super.update(node);
    if (ok) this.tagStyle(node);
    return ok;
  }
  tagStyle(node) {
    const m = node.attrs.tblPr && /<w:tblStyle w:val="([^"]+)"/.exec(node.attrs.tblPr);
    if (m) this.table.setAttribute('data-tstyle', m[1]);
    else this.table.removeAttribute('data-tstyle');
    // Fixed layout keeps the grid's widths; otherwise Word widens a column to fit its longest word (tablefit.js).
    if (/<w:tblLayout\b[^>]*w:type="fixed"/.test(node.attrs.tblPr || '')) this.table.setAttribute('data-fixed', '');
    else this.table.removeAttribute('data-fixed');
    // Autofit to contents (table and cell widths all "auto", as some generators write): Word sizes the
    // columns by their text, up to the page width, not by the grid.
    if (autofitContents(node)) {
      this.table.classList.add('odx-autofit');
      this.table.style.width = '';
      this.table.style.minWidth = '';
      for (const col of this.colgroup.children) col.style.width = '';
    } else {
      this.table.classList.remove('odx-autofit');
      // (The file's own grid, px: the widths Word last laid the table out at, which the fitting takes when it comes out
      // within a pixel and a half of them, tablefit.js.)
      const grid = [...(node.attrs.tblGrid || '').matchAll(/<w:gridCol\b[^>]*\bw:w="(\d+)/g)].map((m) => Math.round((Number(m[1]) / 15) * 100) / 100);
      if (grid.length) this.table.setAttribute('data-grid', grid.join(','));
      else this.table.removeAttribute('data-grid');
      keepFitted(this.table);
    }
    // The table's placement, set property by property: the wrapper's style also carries the view's decorations (the
    // page push's margin-top), which a whole new style attribute would wipe out on each update (a click moving the caret
    // updates the view), leaving the table pulled up onto the page before.
    const place = (node.attrs.deco && node.attrs.deco.style) || '';
    if (place !== (this.place || '')) {
      const parsed = document.createElement('div').style;
      parsed.cssText = place;
      for (const name of this.placed || []) this.dom.style.removeProperty(name);
      this.placed = [...parsed];
      for (const name of this.placed) this.dom.style.setProperty(name, parsed.getPropertyValue(name), parsed.getPropertyPriority(name));
      this.place = place;
    }
    // A table ending in an exact row has its bottom border below that row in Word, not in it: room for it (editor.css).
    const below = lastExactBorder(node);
    if (below) this.table.style.setProperty('--odx-lastadd', `${below}px`);
    else this.table.style.removeProperty('--odx-lastadd');
  }
}

/** px of the bottom border under a table's last row when that row is exact-height, else 0. */
function lastExactBorder(node) {
  const last = node.lastChild;
  if (!last || !/<w:trHeight\b[^>]*w:hRule="exact"/.test(last.attrs.trPr || '')) return 0;
  const side = (xml) => { const m = /<w:bottom\b[^>]*\/>/.exec(xml || ''); if (!m) return null; const v = /w:val="([^"]+)"/.exec(m[0])?.[1]; return v === 'nil' || v === 'none' ? 0 : Number(/w:sz="(\d+)"/.exec(m[0])?.[1] || 4) / 8 * (4 / 3) * (v === 'double' ? 3 : 1); };
  const tbl = side((/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/.exec(node.attrs.tblPr || '') || [])[0]);
  let most = 0;
  last.forEach((cell) => {
    const own = (cell.attrs.tcPr || []).find((x) => x[0] === 'tcBorders');
    const w = own ? side(own[1]) : null;
    most = Math.max(most, w ?? tbl ?? 0);
  });
  return Math.round(most * 100) / 100;
}

const autoW = (xml) => !xml || /w:type="auto"/.test(xml) || /w:w="0"/.test(xml) && !/w:type="pct"/.test(xml);
function autofitContents(node) {
  const pr = node.attrs.tblPr || '';
  if (/<w:tblLayout\b[^>]*w:type="fixed"/.test(pr) || !autoW((/<w:tblW\b[^>]*>/.exec(pr) || [])[0])) return false;
  let all = true;
  node.forEach((row) => row.forEach((cell) => {
    if (!all || cell.attrs.phantom) return;
    const w = (cell.attrs.tcPr || []).find((x) => x[0] === 'tcW');
    if (!autoW(w && w[1])) all = false;
  }));
  return all;
}

/** Cell borders/shading for one table, from its style, tblLook and direct formatting (cached per table shape). */
const byTable = new WeakMap(); // unchanged tables keep their node identity between edits

function tableCells(table, pos, cache, session) {
  const hit = byTable.get(table);
  if (hit) return hit.map((c) => ({ ...c, pos: pos + 1 + c.rel }));
  const map = TableMap.get(table);
  const cells = [];
  const seen = new Set();
  for (let i = 0; i < map.map.length; i++) {
    const rel = map.map[i];
    if (seen.has(rel)) continue;
    seen.add(rel);
    const node = table.nodeAt(rel);
    if (!node || !node.type.spec.tableRole?.includes('cell')) continue; // holes left by ragged rows
    const rect = map.findCell(rel);
    cells.push({ rel, node, row: rect.top, col: rect.left, rowspan: rect.bottom - rect.top, colspan: rect.right - rect.left, tcPr: node.attrs.tcPr, vmTops: node.attrs.vmTops });
  }
  const key = table.attrs.tblPr + '|' + map.width + 'x' + map.height + '|' + cells.map((c) => `${c.row},${c.col},${c.rowspan},${c.colspan},${c.tcPr.map((x) => x[1]).join('')},${(c.vmTops || []).join('.')}`).join(';');
  let res = cache.get(key);
  if (!res) {
    res = tableCellCss(session.styles, table.attrs.tblPr, map.height, map.width, cells, parseXml, session.theme);
    if (cache.size > 500) cache.clear();
    cache.set(key, res);
  }
  const out = cells.map((c, i) => ({ cell: true, rel: c.rel, node: c.node, deco: res.css[i] ? { style: res.css[i] } : null }));
  byTable.set(table, out);
  return out.map((c) => ({ ...c, pos: pos + 1 + c.rel }));
}

/** The paragraph mark's own font size (half-points), if it sets one. */


/**
 * Paragraph decorations that depend on neighbours: list labels and indents from
 * numbering.xml, and Word's "don't add space between paragraphs of the same
 * style" (contextualSpacing). Recomputed in one pass after every change.
 */
function paragraphPlugin(session) {
  // A paragraph's space before or after in twips (none under HTML-style automatic spacing).
  const spacing = (node, side) => ((node.attrs.view[`${side}Auto`] ?? session.styleProp(node.attrs.styleId, `${side}Auto`)) ? 0 : node.attrs.view[side] ?? session.styleProp(node.attrs.styleId, side) ?? session.styles.defaults.pPr[side] ?? 0);
  const tableCache = new Map();
  const styleCs = new Map();
  const contextual = (attrs) => {
    if (attrs.view.contextualSpacing != null) return attrs.view.contextualSpacing;
    const id = attrs.styleId || session.styles.defaultParagraph;
    if (!styleCs.has(id)) styleCs.set(id, !!(id && resolveStyle(session.styles.styles, id).pPr.contextualSpacing));
    return styleCs.get(id);
  };
  const borderKey = (attrs) => {
    const prop = (k) => attrs.view[k] ?? session.styleProp(attrs.styleId, k);
    const b = prop('borders');
    return b && (b.top || b.bottom || b.left || b.right) ? JSON.stringify([b, prop('indLeft') ?? 0, prop('indRight') ?? 0]) : null;
  };
  const compute = (doc) => {
    const next = listCounter(session.numbering);
    const decos = [];
    // Each section's own side margins, against the page's (the last section's): px to add to a top-level paragraph's
    // left and right (sample S106: its body sections have 0.75in sides under a 1in last section).
    const page = session.page();
    const sides = new Map(); // top-level index -> [dLeft, dRight]
    const collapse = new Set(); // top-level indexes of paragraphs that take no room
    {
      const ends = [];
      doc.forEach((node, _, i) => { const x = (node.attrs.extra || []).find((e) => e[0] === 'sectPr'); if (x) ends.push([i, x[1]]); });
      const side = (xml, a) => { const m = new RegExp(`<w:pgMar[^>]*\\sw:${a}="(\\d+)"`).exec(xml); return m ? twipsToPx(m[1]) : null; };
      let e = 0;
      doc.forEach((node, _, i) => {
        while (e < ends.length && ends[e][0] < i) e++;
        const xml = e < ends.length ? ends[e][1] : null;
        if (!xml) return;
        const l = side(xml, 'left'), r = side(xml, 'right');
        // (A wider page, a landscape section's, moves the right edge out by the difference.)
        const pw = /<w:pgSz[^>]*\sw:w="(\d+)"/.exec(xml);
        const dw = pw ? twipsToPx(pw[1]) - page.width : 0;
        const dl = l == null ? 0 : l + session.gutter(xml).left - page.left, dr = (r == null ? 0 : r - page.right) - dw;
        if (Math.abs(dl) > 0.1 || Math.abs(dr) > 0.1) sides.set(i, [dl, dr]);
      });
      // Sections in newspaper columns, in a document of several (pages.js lays them out column by column): their
      // paragraphs are a column wide. (A one-section document's columns narrow the whole content instead.)
      if (ends.length) {
        const all = ends.concat([[doc.childCount - 1, doc.attrs.sectPr || '']]);
        const inner = page.width - page.left - page.right;
        all.forEach(([hi, xml], n) => {
          const c = columnsIn(xml, page);
          if (!c) return;
          const narrow = inner - c.width;
          const lo = n ? all[n - 1][0] + 1 : 0;
          for (let i = lo; i <= hi; i++) sides.set(i, [0, narrow, !!c.unequal]);
          // The empty paragraph carrying the break into the columns takes no room in Word (corpus 2: the columns start
          // right under the text above it).
          if (n && !doc.child(all[n - 1][0]).childCount) collapse.add(all[n - 1][0]);
        });
      } else if (columnsIn(doc.attrs.sectPr, page)?.unequal) doc.forEach((_, __, i) => sides.set(i, [0, 0, true]));
    }
    // An empty paragraph that carries a continuous section break takes no room in Word, nor its own spacing (Word's prints
    // of generated documents: the line under it 18.7px below the text above, whatever its size, spacing or line; a
    // bookmark or an empty run in it changes nothing, a space or a tab makes it a line). The space after the paragraph
    // above and the space before the one below are both kept, one under the other; without
    // doNotUseHTMLParagraphAutoSpacing the space before the one below is less the break paragraph's own space after.
    // First in the document or right under a table it keeps its line and its space after, and loses its space before.
    // (Other break types leave the paragraph its line: the next section starts on another page.)
    const sectNone = new Map(); // top-level index -> px pulled off the space before the next paragraph
    const sectKept = new Set();
    {
      const ends = [];
      doc.forEach((node, _, i) => { const x = (node.attrs.extra || []).find((e) => e[0] === 'sectPr'); if (x) ends.push([i, x[1]]); });
      ends.forEach(([i], n) => {
        const next = n + 1 < ends.length ? ends[n + 1][1] : doc.attrs.sectPr || '';
        if (!/<w:type\b[^>]*\bw:val="continuous"/.test(next) || collapse.has(i)) return;
        const para = doc.child(i);
        if (para.type !== schema.nodes.paragraph || para.textContent) return;
        let ink = false;
        para.descendants((c) => { if (!ZERO_WIDTH.has(c.attrs?.kind) && c.attrs?.kind !== 'proofErr') ink = true; return !ink; });
        if (ink) return;
        let k = i - 1;
        while (k >= 0 && doc.child(k).type === schema.nodes.opaque_block) k--;
        if (k < 0 || doc.child(k).type !== schema.nodes.paragraph) { sectKept.add(i); return; }
        // (All there is of its section, the paragraph above ending the one before: it keeps its line, as a section of
        // columns' does. Word's prints: three such in a row, the first no room, the next two a line each.)
        if (n && ends[n - 1][0] === k) return;
        const after = i + 1 < doc.childCount && doc.child(i + 1).type === schema.nodes.paragraph ? doc.child(i + 1) : null;
        const plain = session.compat?.has('doNotUseHTMLParagraphAutoSpacing');
        sectNone.set(i, plain || !after ? 0 : Math.min(twipsToPx(Number(spacing(para, 'after')) || 0), twipsToPx(Number(spacing(after, 'before')) || 0)));
      });
    }
    // Contextual spacing across a table's edges, in Word's reading order: the paragraph before the table, then each row's
    // cells left to right, each row closed by its end mark, a paragraph of the default style. A paragraph's neighbour out
    // of its cell is the paragraph before the table, the next or previous cell's, or that mark (Word: the space after a
    // paragraph above a table goes when the first cell's paragraph has its style; a cell's last paragraph of the default
    // style loses its space after under any row; the paragraph under a table loses its space before when of that style).
    const sid = (n) => n.attrs.styleId || session.styles.defaultParagraph || '';
    const mark = sid({ attrs: {} });
    const edges = new Map(); // a cell's first or last paragraph -> { prev, next }: its neighbours' styles out of the cell
    const skip = (n) => n.type === schema.nodes.opaque_block && (ZERO_WIDTH.has(n.attrs.kind) || n.attrs.kind === 'proofErr');
    const ends = (cell) => {
      const kids = [];
      cell.forEach((c) => { if (!skip(c)) kids.push(c); });
      const para = (n) => (n && n.type === schema.nodes.paragraph ? n : null);
      return { first: para(kids[0]), last: para(kids[kids.length - 1]) };
    };
    const lastIn = (cell) => { let last = null; cell.forEach((c) => { if (!skip(c)) last = c; }); return last; };
    const edge = (n, k, v) => { if (!n) return; const e = edges.get(n) || {}; e[k] = v; edges.set(n, e); };
    const tableEdges = (table, before) => {
      let r = 0;
      table.forEach((row) => {
        if (row.type !== schema.nodes.table_row) return;
        const cells = [];
        row.forEach((c) => cells.push(ends(c)));
        cells.forEach((c, i) => {
          const left = i ? (cells[i - 1].last ? sid(cells[i - 1].last) : null) : r ? mark : before ? sid(before.node) : null;
          edge(c.first, 'prev', left);
          edge(c.last, 'next', i < cells.length - 1 ? (cells[i + 1].first ? sid(cells[i + 1].first) : null) : mark);
        });
        if (!r && before?.cs && cells[0]?.first && sid(cells[0].first) === sid(before.node)) before.style.push('margin-bottom:0', '--odx-sa:0px');
        r++;
      });
    };
    const visit = (parent, offset) => {
      let prev = null;
      let edgePrev = null; // the style of the row end mark above this block, when it follows a table
      let noRoom = 0; // blocks after a table placed down from the text with no room beside it (styles.js --odx-fnoroom)
      let floatGap = null; // px: the space after the paragraph before a floating table, when this block follows that table
      let fUnder = null; // px: how far down from this block such a table with no room beside it is placed
      parent.forEach((node, rel, index) => { // (index: a top-level block's, for sides and collapse; doc.childBefore made this quadratic)
        const pos = offset + rel;
        if (node.type !== schema.nodes.paragraph) {
          // A bookmark or other marker between two paragraphs shows nothing, so they still count as neighbours.
          if (node.type === schema.nodes.opaque_block && (ZERO_WIDTH.has(node.attrs.kind) || node.attrs.kind === 'proofErr')) return;
          const before = prev;
          prev = null;
          edgePrev = null;
          floatGap = null;
          fUnder = null;
          if (node.type === schema.nodes.table) {
            // (Its column: its own section's, which may be narrower than the last's, a portrait page before landscape ones.)
            const own = parent === doc ? sides.get(index) || [0, 0] : null;
            const place = tablePlacementCss(session.styles, node.attrs.tblPr, session.compatMode, parseXml, node.attrs.tblGrid, own ? page.width - page.left - page.right - own[0] - own[1] : 0);
            if (!/(?:^|;)float:/.test(place || '')) { tableEdges(node, before); edgePrev = mark; }
            noRoom = /--odx-fnoroom/.test(place || '') ? 40 : 0;
            fUnder = noRoom ? Number(/(?:^|;)margin-top:([\d.]+)px/.exec(place)?.[1] || 0) : null;
            // A floating table in the text between two paragraphs: Word keeps both their spaces, the one after the first
            // and the one before the second (they don't overlap as they would side by side), and the table starts below
            // the first's (editor.css: while the table floats; one the layout puts in the flow, longer than a page, has its
            // margins as they are).
            let style = place;
            if (/(?:^|;)float:/.test(place || '') && before) {
              floatGap = twipsToPx(spacing(before.node, 'after'));
              if (floatGap) style += `;--odx-fgap:${+floatGap.toFixed(2)}px`;
            }
            // In a section of newspaper columns among others (or with its own side margins), as narrow as its column, as
            // the paragraphs are: an autofit table fits itself to the column, as in Word.
            const sd = parent === doc && sides.get(index);
            // (With the table's own place: its indent, and in a document of Word 2010's mode or before its cell margin's
            // width out into the page margin, were lost here, a table 7.2px right of Word's in a section narrower than
            // the document's widest.)
            if (sd && !/(?:^|;)float:/.test(place || '')) {
              const length = (m) => (m && /^-?[\d.]+(?:px|pt)$/.test(m[1].trim()) ? m[1].trim() : null);
              const ml = length(/(?:^|;)margin-left:([^;]+)/.exec(place || '')), mr = length(/(?:^|;)margin-right:([^;]+)/.exec(place || ''));
              style = [style, `margin-inline-start:${ml ? `calc(${sd[0]}px + ${ml})` : `${sd[0]}px`}`, `margin-inline-end:calc(${sd[1]}px${mr ? ` + ${mr}` : ''}${sd[2] ? ' + var(--odx-coln, 0px)' : ''})`].filter(Boolean).join(';');
            }
            decos.push({ cell: true, pos, node, deco: style ? { style } : null });
            tableCells(node, pos, tableCache, session).forEach((c) => decos.push(c));
          }
          if (!node.isLeaf) visit(node, pos + 1);
          return;
        }
        const attrs = {};
        const style = [];
        // An empty paragraph that only carries a section break isn't numbered (Word: sample S106's empty Heading 1 holding
        // the break doesn't make the next heading 2.0).
        let sectOnly = !node.textContent.trim() && (node.attrs.extra || []).some((e) => e[0] === 'sectPr');
        if (sectOnly) node.descendants((c) => { if (c.type.name === 'image') sectOnly = false; });
        const l = sectOnly ? null : cmd.paragraphList(session, node.attrs);
        const label = l ? next(l.numId, l.ilvl) : null;
        listLabel(node, session, l, label, attrs, style);
        // Word sizes each line from the runs on it; CSS also counts the paragraph's
        // own font (the strut). When every run shares a font or size, give it to the
        // paragraph too, so small-text paragraphs (tables, code) aren't too tall.
        const run = uniformRun(node, (id) => session.charStyle?.(id), session.theme);
        lineSize(node, session, run, label, attrs, style);
        pictureLine(node, page, attrs);
        labelLine(node, session, l, label, run, style);
        labelFormat(node, session, l, label, style);
        // Long documents skip drawing off-screen paragraphs; a size guess from the
        // text length (and pictures) keeps the page count close until they're drawn.
        if (parent === doc && doc.childCount > 1500) {
          let px = 0;
          node.descendants((c) => { if (c.attrs?.height) px += parseFloat(c.attrs.height) || 0; });
          const lines = Math.max(1, Math.ceil(node.textContent.length / 90));
          style.push(`contain-intrinsic-size:auto calc(${lines * 1.25}em + ${Math.round(px)}px)`);
        }
        // A drop cap (w:framePr w:dropCap): Word draws the paragraph as a frame at the left of the next one, which wraps
        // beside it; the frame takes no room of its own in the flow.
        const frame = (node.attrs.extra || []).find((e) => e[0] === 'framePr');
        if (frame && /w:dropCap="(?:drop|margin)"/.test(frame[1])) {
          style.push('float:left', 'margin:0', 'padding:0', 'min-height:0');
          attrs.class = [attrs.class, 'odx-dropcap'].filter(Boolean).join(' ');
        }
        if (floatGap) style.push(`--odx-fgap:${+floatGap.toFixed(2)}px`, `--odx-fbefore:${+twipsToPx(spacing(node, 'before')).toFixed(2)}px`);
        floatGap = null;
        const cs = contextual(node.attrs);
        // HTML-style automatic spacing: 14pt, shared with the paragraph beside it (the margins collapse), none above the
        // first paragraph of a cell, and none between two items of one list (Word's PDFs). Under Word 2003's rule
        // (w:doNotUseHTMLParagraphAutoSpacing) it's a plain 5pt each side, added to the next paragraph's.
        const auto = (n, side) => n.type.name === 'paragraph' && (n.attrs.view[side] ?? session.styleProp(n.attrs.styleId, side));
        const plainAuto = session.compat?.has('doNotUseHTMLParagraphAutoSpacing');
        if (plainAuto) {
          if (auto(node, 'beforeAuto')) style.push(`margin-top:${prev && auto(prev.node, 'afterAuto') ? 10 : 5}pt`);
          if (auto(node, 'afterAuto')) style.push('margin-bottom:5pt', '--odx-sa:5pt');
          // Under that rule the space after a paragraph and the space before the next are both kept, one under the other
          // (Word's prints: 8pt after and 4pt before, the lines 37.3px apart; without the setting the larger, 32px). The
          // margins collapse to the larger, so the one above is added to this paragraph's own.
          if (prev && prev.node.type.name === 'paragraph' && !auto(node, 'beforeAuto') && !auto(prev.node, 'afterAuto') && !(parent === doc && (sectNone.has(index) || sectNone.has(index - 1)))) {
            const b = spacing(node, 'before'), a = prev.cs && prev.node.attrs.styleId === node.attrs.styleId ? 0 : spacing(prev.node, 'after');
            if (a > 0 && b > 0) style.push(`margin-top:${+twipsToPx(a + b).toFixed(2)}px`);
          }
        } else {
          // (A paragraph's own space, w:after="150", over a style with auto spacing that it doesn't turn off: still the auto
          // 14pt, as in Word's PDFs, the auto flag being kept apart from the number.)
          if (node.attrs.view.before != null && node.attrs.view.beforeAuto == null && auto(node, 'beforeAuto')) style.push('margin-top:14pt');
          if (node.attrs.view.after != null && node.attrs.view.afterAuto == null && auto(node, 'afterAuto')) style.push('margin-bottom:14pt', '--odx-sa:14pt');
          if (rel === 0 && auto(node, 'beforeAuto')) style.push('margin-top:0');
          // (One list: one numId, or two over one abstract list, as Word's PDFs have it. Each side on its own: an item's auto space after goes under the next item of its list even when that one has
          // none of its own, and the next's auto space before goes too, under an item with a plain 6pt after, which stays.)
          if (prev && l && prev.list && sameList(session, prev.list.numId, l.numId)) {
            // (In table cells too, as in Word: samples 1477, 1576, 1652 and 2311. Samples 1434 and 2665, whose cells
            // hold about 65 lines a page where Word's hold 50 either way, then break their pages elsewhere.)
            if (auto(node, 'beforeAuto')) style.push('margin-top:0');
            if (auto(prev.node, 'afterAuto')) prev.style.push('margin-bottom:0', '--odx-sa:0px');
          }
        }
        const bdr = borderKey(node.attrs);
        // Paragraphs with the same borders form one group, as in Word: its top border above the first, its bottom border
        // under the last (and a "between" line, if any, between them).
        if (prev && bdr && prev.bdr === bdr && prev.node.attrs.styleId === node.attrs.styleId) {
          // The group's "between" line, if it has one, under each paragraph but the last, its space above and below it
          // (sample S134's ruled lines: 22.56px apart in Word's PDF).
          const between = (node.attrs.view.borders ?? session.styleProp(node.attrs.styleId, 'borders'))?.between;
          const line = between && between.val && between.val !== 'nil' && between.val !== 'none';
          style.push('border-top:none', `padding-top:${line ? between.space || 0 : 0}pt`);
          if (line) prev.style.push(`border-bottom:${borderCss(between)}`, `padding-bottom:${between.space || 0}pt`);
          else prev.style.push('border-bottom:none', 'padding-bottom:0');
        }
        if (prev && prev.node.attrs.styleId === node.attrs.styleId) {
          // Each paragraph's contextual spacing removes its own space; an empty one's above also removes the space before
          // this one (sample S072: an item turning it off for itself still loses its space before under an empty one that has
          // it), but one with text doesn't (doc 14: "(1) Follow all..." keeps its 6pt under "...I recognize that I must:").
          if (cs || (prev.cs && !prev.node.textContent)) style.push('margin-top:0');
          if (prev.cs) prev.style.push('margin-bottom:0', '--odx-sa:0px');
        }
        if (cs) {
          const e = edges.get(node);
          const left = prev ? null : e?.prev ?? edgePrev;
          if (left != null && left === sid(node)) style.push('margin-top:0');
          if (e?.next != null && e.next === sid(node)) style.push('margin-bottom:0', '--odx-sa:0px');
        }
        edgePrev = null;
        // Holding a drawing placed from its paragraph's top (reader.js fromParagraphTop): how far above the text that top is,
        // the space this paragraph adds to the space after the one above (the two collapse, the larger kept).
        let fromTop = false;
        node.descendants((c) => { if (fromTop) return false; if (c.type.name === 'image' && /--odx-sb/.test(c.attrs.place || '')) fromTop = true; return true; });
        if (fromTop) {
          const last = (list, k) => { const v = [...list].reverse().find((x) => x.startsWith(`${k}:`)); return v == null ? null : parseFloat(v.slice(k.length + 1)) || 0; };
          const pt = 4 / 3;
          const before = last(style, 'margin-top') != null ? last(style, 'margin-top') * pt
            : auto(node, 'beforeAuto') ? (plainAuto ? 5 : 14) * pt : twipsToPx(spacing(node, 'before'));
          const after = !prev ? 0 : last(prev.style, 'margin-bottom') != null ? last(prev.style, 'margin-bottom') * pt
            : auto(prev.node, 'afterAuto') ? (plainAuto ? 5 : 14) * pt : twipsToPx(spacing(prev.node, 'after'));
          const sb = +Math.max(0, before - after).toFixed(2);
          // (--odx-sbtb: that space again under a top-and-bottom shape, over the paragraph's own text: reader.js.)
          style.push(`--odx-sb:${sb}px`, `--odx-sbtb:${node.textContent.trim() ? sb : 0}px`);
        }
        // (An empty one after such a table: something to fit on its line, so it goes under the table, as in Word.)
        const fempty = noRoom > 0 && !node.childCount;
        if (noRoom > 0) noRoom--;
        // The paragraph right after it, when its space before and first line don't fit in the room above the table (an
        // estimate from its size): Word sets it under the table with its space before kept (a Word test: 8pt before, under
        // a table 10.45pt down; 0pt before, a 13.5pt line, stayed above one 15.45pt down). The browser leaves the margin
        // above, beside the float (editor.css: odx-funder, while the table floats).
        let funder = false;
        if (fUnder != null) {
          const b = Math.min(1056, twipsToPx(Number(spacing(node, 'before')) || 0));
          const sz = Math.min(3276, Math.max(2, Number(run.size || session.styleRunProp(node.attrs.styleId, 'size') || session.styles.defaults.rPr.size || 20) || 20));
          if (b + (sz / 2) * (4 / 3) * 1.15 > fUnder) { funder = true; style.push(`--odx-funder:${+b.toFixed(2)}px`); }
          fUnder = null;
        }
        if (parent === doc && sectNone.has(index)) style.push(`--odx-sectpull:${+sectNone.get(index).toFixed(2)}px`);
        const cls = [attrs.class, ...changeClasses(node), parent === doc && collapse.has(index) && 'odx-collapse', parent === doc && sectNone.has(index) && 'odx-sectnone', parent === doc && sectKept.has(index) && 'odx-sectkept', fempty && 'odx-fempty', funder && 'odx-funder'].filter(Boolean);
        if (cls.length) attrs.class = cls.join(' ');
        headingRole(node, session, attrs);
        // A paragraph in a frame (frames.js): at its place, the text beside it; one after another in the same frame below it.
        const fr = parent === doc && frameOf(node);
        if (fr) {
          const f = frameStyle(fr, index > 0 && frameOf(doc.child(index - 1)) === fr, page);
          style.push(...f.style);
          Object.assign(attrs, f.attrs);
          // Its indents are inside the frame's width, as in Word (the frame's own margins place it): as padding, the
          // width less them (a Quote's 0.6in each side in a 3in frame).
          const ind = (k) => Math.max(0, twipsToPx(node.attrs.view[k] ?? session.styleProp(node.attrs.styleId, k) ?? 0));
          const il = ind('indLeft'), ir = ind('indRight');
          const wi = f.style.findIndex((x) => /^width:/.test(x));
          if ((il || ir) && wi >= 0 && /float:/.test(f.style.join(';'))) style.push(`width:calc(${f.style[wi].slice(6)} - ${+(il + ir).toFixed(2)}px)`, `padding-left:${+il.toFixed(2)}px`, `padding-right:${+ir.toFixed(2)}px`);
        }
        const sd = !fr && parent === doc && sides.get(index);
        if (sd) {
          // (Added to the paragraph's own indents, the last margin it sets.)
          // (Else the paragraph's own or its style's indent, which this margin would otherwise replace: a Quote's 0.6in
          // each side in a column.)
          const ind = (k) => { const tw = node.attrs.view[k] ?? session.styleProp(node.attrs.styleId, k); return tw ? `calc(${tw / 20}pt - var(--odx-${k === 'indLeft' ? 'bdl' : 'bdr'}, 0pt))` : '0pt'; };
          const own = (k) => { const v = [...style].reverse().find((x) => x.startsWith(`${k}:`)); return v ? v.slice(k.length + 1).trim() : ind(k === 'margin-inline-start' ? 'indLeft' : 'indRight'); };
          // (The section's sides are the page's left and right: a right-to-left paragraph's start is its right.)
          const rtl = (node.attrs.view.bidi ?? session.styleProp(node.attrs.styleId, 'bidi')) === true;
          const [ds, de] = rtl ? [sd[1], sd[0]] : sd;
          // (Columns of their own widths: laid out the widest wide, less --odx-coln on the right, which pages.js sets
          // for the column each block lands in.)
          const coln = sd[2] ? ' + var(--odx-coln, 0px)' : '';
          style.push(`margin-inline-start:calc(${own('margin-inline-start')} + ${ds}px${rtl ? coln : ''})`, `margin-inline-end:calc(${own('margin-inline-end')} + ${de}px${rtl ? '' : coln})`);
        }
        // A cell whose end mark is left out of its row's height (w:hideMark): its last paragraph, empty, takes no room, its
        // spacing and size with it (Word: a row of such cells is as tall as its cells' margins; one with text, or an
        // unmarked empty cell beside it, keeps its line). Still there, at no height, for the caret. (Not a numbered one:
        // Word keeps its line, as it would its number's.)
        if (parent.type === schema.nodes.table_cell && !node.childCount && !l && (parent.attrs.tcPr || []).some((x) => x[0] === 'hideMark') && lastIn(parent) === node) {
          style.push('height:0', 'min-height:0', 'margin-top:0', 'margin-bottom:0', 'padding-top:0', 'padding-bottom:0', 'overflow:hidden');
        }
        prev = { node, pos, attrs, style, cs, bdr, list: l };
        decos.push(prev);
      });
    };
    visit(doc, 0);
    return decos.map((d) => {
      if (d.cell) return d;
      if (d.style.length) d.attrs.style = d.style.join(';');
      return { pos: d.pos, node: d.node, deco: d.attrs.style || d.attrs.class || d.attrs.role ? d.attrs : null };
    });
  };
  /** Write computed values into paragraphs' `deco` attr where they changed. */
  // Many paragraphs changed (Accept All on thousands of changes takes their change classes off): one step from the first
  // to the last, not a step each (every plugin's decorations were mapped through each of thousands: a second).
  const MANY = 200;
  const sync = (tr) => {
    const changes = changed(tr.doc);
    if (changes.size > MANY) {
      const doc = tr.doc;
      const next = rebuild(doc, changes);
      const a = doc.content.findDiffStart(next.content);
      if (a == null) return tr;
      const { a: e } = doc.content.findDiffEnd(next.content);
      // (The positions are the same in both documents: only attributes changed. The selection is put back where it was.)
      const sel = tr.selection.toJSON();
      tr.replace(a, e, next.slice(a, e));
      try { tr.setSelection(Selection.fromJSON(tr.doc, sel)); } catch { /* (left as mapped) */ }
      return tr;
    }
    for (const [pos, deco] of changes) { const node = tr.doc.nodeAt(pos); tr.setNodeMarkup(pos, null, { ...node.attrs, deco }); }
    return tr;
  };
  // Stored as node attrs rather than view decorations: ProseMirror checks every
  // top-level decoration against every top-level node, which is quadratic on long documents.
  /** The paragraphs whose `deco` changed: position to new value. */
  const changed = (doc) => {
    const changes = new Map();
    for (const { pos, node, deco } of compute(doc)) if (JSON.stringify(deco) !== JSON.stringify(node.attrs.deco)) changes.set(pos, deco);
    return changes;
  };
  /** Same as sync, for a whole document in one pass (used when opening). */
  const initial = (doc) => {
    const changes = changed(doc);
    if (!changes.size) return doc;
    return rebuild(doc, changes);
  };
  const rebuild = (doc, changes) => {
    const walk = (node, offset) => {
      const kids = [];
      node.forEach((child, rel) => {
        const pos = offset + rel;
        const inner = child.isLeaf || child.isTextblock ? child : walk(child, pos + 1);
        kids.push(changes.has(pos) ? child.type.create({ ...child.attrs, deco: changes.get(pos) }, inner === child ? child.content : inner.content, child.marks) : inner);
      });
      return node.copy(Fragment.from(kids));
    };
    return walk(doc, 0);
  };
  return {
    sync,
    initial,
    /** The styles changed (another co-editor's arrived): what was worked out from the old ones is dropped. */
    reset() { tableCache.clear(); styleCs.clear(); },
    plugin: new Plugin({
      appendTransaction(trs, _old, state) {
        if (!trs.some((t) => t.docChanged)) return null;
        const tr = sync(state.tr);
        // (The formatting the caret carries (bold after Enter at the end of a bold item) stays: a step would drop it.)
        if (tr.docChanged && state.storedMarks) tr.setStoredMarks(state.storedMarks);
        return tr.docChanged ? tr : null;
      },
    }),
  };
}

/** Mark the top-level block holding the caret; CSS keeps it out of content-visibility skipping, which breaks caret placement. */
const activeBlock = new Plugin({
  props: {
    decorations(state) {
      const { $from } = state.selection;
      if ($from.depth < 1) return null;
      const start = $from.before(1);
      return DecorationSet.create(state.doc, [Decoration.node(start, start + $from.node(1).nodeSize, { class: 'odx-active' })]);
    },
  },
});

/**
 * Home and End (with Shift too) on a paragraph's first or last line when the paragraph begins or ends in a marker inside
 * a link or other mark, such as the field end after a Word TOC entry's page number (toc-u-only.docx: the hyperlink holds
 * the entry, its tab, and the PAGEREF field to its end). The browser's own End stopped before that uneditable marker, or
 * went to the end of the document; its Home from there went to after the entry's tab, or out of the paragraph; and what
 * was typed next went there. They go to the paragraph's start or end, as Word's do. (Other lines, and other paragraphs,
 * keep the browser's.)
 */
export const lineEdge = (dir, extend) => (state0, dispatch, view) => {
  if (!view) return false;
  // (A click just before: the browser's caret read first, or the key acted on where the caret had been.)
  view.domObserver?.flush?.();
  const { state } = view;
  const { $head, empty } = state.selection;
  if ((!empty && !extend) || !$head.parent.isTextblock) return false;
  const marker = (n) => n && !n.isText && n.marks.length > 0;
  if (!marker($head.parent.firstChild) && !marker($head.parent.lastChild)) return false;
  const to = dir < 0 ? $head.start() : $head.end();
  if (to === $head.pos) return true;
  try {
    const here = view.coordsAtPos($head.pos, dir < 0 ? -1 : 1), there = view.coordsAtPos(to, dir < 0 ? 1 : -1);
    if (Math.abs(here.top - there.top) > Math.max(2, (here.bottom - here.top) / 2)) return false; // (not that line)
  } catch { return false; }
  if (dispatch) dispatch(state.tr.setSelection(TextSelection.create(state.doc, extend ? state.selection.anchor : to, to)).scrollIntoView());
  return true;
};

/**
 * Pasted paragraphs' built-in styles by the document's own ids (Word's English ids, as a pasted <h1> gets, are another
 * language's in a document from a localised Word: Heading1 is "Ttulo1" from Spanish Word).
 *
 * Paragraphs pasted into an empty one: the first keeps its own paragraph formatting (a pasted heading stays a heading),
 * as in Word when what was copied holds that paragraph's mark, ending where the next begins. ProseMirror joins it into
 * the paragraph at the caret, which kept that paragraph's style. (Part of one paragraph takes the formatting of where it
 * goes, as in Word; and not while tracking changes, where Word would record the change of style.)
 */
function localStylesPlugin(session) {
  const map = (frag) => Fragment.fromArray(frag.content.map((n) => {
    if (n.isLeaf) return n;
    const id = n.type === schema.nodes.paragraph && n.attrs.styleId;
    const local = id && session.builtInId(id);
    return n.type.create(local && local !== id ? { ...n.attrs, styleId: local } : n.attrs, map(n.content), n.marks);
  }));
  let into = null; // { pos, attrs }: the empty paragraph pasted into and the first pasted paragraph's attrs
  return new Plugin({
    props: {
      transformPasted: (slice, view) => {
        const out = new Slice(map(slice.content), slice.openStart, slice.openEnd);
        const { selection } = view.state;
        const first = out.content.firstChild;
        into = null;
        // (Several paragraphs, or a lone heading: Word gives it its heading style. Words from one paragraph of body text
        // take the paragraph's they go into.)
        const heading = (n) => { const l = cmd.outlineLevel(session, n); return l != null && l >= 0 && l < 9; };
        if (selection.empty && selection.$from.parent.type === schema.nodes.paragraph && !selection.$from.parent.content.size
          && out.openStart > 0 && first?.type === schema.nodes.paragraph && (out.content.childCount > 1 || heading(first)) && !tracking(view.state).on) {
          into = { pos: selection.$from.before(), attrs: first.attrs };
        }
        return out;
      },
    },
    appendTransaction(trs, _old, state) {
      const tr = trs.find((t) => t.getMeta('uiEvent') === 'paste');
      if (!tr || !into) return null;
      const { attrs } = into;
      let pos = into.pos;
      into = null;
      for (const t of trs.slice(trs.indexOf(tr))) pos = t.mapping.map(pos, -1);
      const node = state.doc.nodeAt(pos);
      if (!node || node.type !== schema.nodes.paragraph) return null;
      // (The paragraph's own identity and content control stay its own: a copy's paraIds would be duplicates.)
      return state.tr.setNodeMarkup(pos, null, { ...attrs, pattrs: node.attrs.pattrs, sdt: node.attrs.sdt, deco: null });
    },
  });
}

/** Are two numIds one list for Word's spacing between items: the same, or two instances of one abstract list? */
function sameList(session, a, b) {
  if (a === b) return true;
  const nums = session.numbering?.nums;
  const x = nums?.get(String(a));
  const y = nums?.get(String(b));
  return !!(x && y && x.abstractId != null && x.abstractId === y.abstractId);
}

export function createEditor(mount, session, { onChange, author = 'Author', collab = null } = {}) {
  const scope = `odx-doc-${++scopeSeq}`;
  const styleEl = document.createElement('style');
  styleEl.textContent = session.css(`.${scope}`) + '\n' + tableStylesCss(`.${scope}`, session.styles, session.theme, { oldSizes: !!session.oldTableSizes }) + '\n' + session.picBulletCss(`.${scope}`);
  document.head.appendChild(styleEl);

  const page = session.page();
  const paras = paragraphPlugin(session);

  // Co-editing (src/collab.js): the shared body's undo, of this editor's own edits only, in place of the history's.
  const doUndo = collab ? (s, d) => (d ? collab.undo(s) : true) : undo;
  const doRedo = collab ? (s, d) => (d ? collab.redo(s) : true) : redo;
  const keys = keymap({
    'Mod-z': doUndo,
    'Mod-y': doRedo,
    'Mod-Shift-z': doRedo,
    Home: lineEdge(-1, false),
    'Shift-Home': lineEdge(-1, true),
    End: lineEdge(1, false),
    'Shift-End': lineEdge(1, true),
    // (At a table's edges: into, out of and across its cells, in right-to-left text too, keeping the caret's column.)
    ArrowLeft: tableArrow(-1),
    ArrowRight: tableArrow(1),
    ArrowUp: tableVert(-1),
    ArrowDown: tableVert(1),
    'Mod-b': spoken('bold', 'Bold'),
    'Mod-i': spoken('italic', 'Italic'),
    'Mod-u': spoken('underline', 'Underline'),
    'Mod-e': cmd.setAlign('center'),
    'Mod-l': cmd.setAlign('left'),
    'Mod-r': cmd.setAlign('right'),
    'Mod-j': cmd.setAlign('both'),
    'Mod-Shift-j': cmd.setAlign('distribute'), // Word's Distributed: the last line spread too
    'Shift-Enter': cmd.insertBreak(),
    'Mod-Enter': cmd.insertBreak('page'),
    'Mod-Alt-f': (s, d, v) => (v ? insertNote(v, 'footnote') : true),
    'Mod-Alt-d': (s, d, v) => (v ? insertNote(v, 'endnote') : true),
    'Mod-Alt-m': (s, d, v) => { if (v) addComment(v); return true; },
    'Mod-Shift-8': (s, d, v) => (v ? toggleMarks(v) : true), // Show/Hide ¶, as in Word
    // Update Fields, as in Word: those in the selection, or all of them with nothing selected.
    F9: (s, d, v) => (v ? updateFieldsCommand(v) : true),
    Enter: cmd.splitParagraph(session),
    Backspace: chainCommands(trackedDelete(-1), cmd.outdentOrUnlist(session)),
    Delete: trackedDelete(1),
    // (Not where the document can only be read: Tab goes on to the next control, protect.js.)
    Tab: (s, d, v) => !tabLeaves(s) && chainCommands(goToNextCell(1), cmd.changeListLevel(session, 1), cmd.insertTab)(s, d, v),
    'Shift-Tab': (s, d, v) => !tabLeaves(s) && chainCommands(goToNextCell(-1), cmd.changeListLevel(session, -1))(s, d, v),
  });

  const doc = paras.initial(session.doc);
  if (collab) collab.seed(doc); // (the first editor in a room gives it its body; the others show the room's)
  const state = EditorState.create({
    doc,
    plugins: [...formsPlugin(), bindingsPlugin(session), chartPlugin(), shapeKeysPlugin() /* (before the keymap: its Enter would replace a selected chart, its arrows move the caret off a shape) */, ...(collab ? collab.plugins() : [history()]), autocorrectPlugin, keys, keymap(baseKeymap), columnResizing({ View: StyledTableView }), tableEditing(), paras.plugin, sectionJoinPlugin, activeBlock, pagePushPlugin, textSpansPlugin, justifyPlugin, wrapBothPlugin, tableClearPlugin, lineFixPlugin, widthsPlugin, fitTextPlugin, joinsPlugin, marksPlugin, csFormatPlugin, pastePictures(session), pictureLinks, imageResizePlugin(), textBoxPlugin(), trackPlugin({ on: session.trackRevisions, author }), commentsPlugin(session), protectPlugin(), localStylesPlugin(session), hyphenatePlugin(session), mailMergePlugin(session)],
  });

  mount.classList.add('odx-canvas');
  // Keyboard users can scroll the pages too (WCAG 2.1.1): the scrolling area is a named, focusable region.
  // No role here: a named region made NVDA announce "Document pages, region" instead of the document itself.
  // It's focusable so keyboard users can scroll it (2.1.1). Tabbing in moves straight on into the document;
  // Shift+Tab out of the document stops here once, then leaves, so there's no trap.
  mount.tabIndex = 0;
  mount.addEventListener('focus', (e) => { if (!view.dom.contains(e.relatedTarget)) view.focus(); });
  const sheet = document.createElement('div');
  sheet.className = `odx-page ${scope}`;
  if (session.doc.childCount > 1500) sheet.classList.add('odx-large');
  sheet.style.width = `${page.width}px`;
  styleEl.textContent += `\n@page{size:${page.width}px ${page.height}px;margin:0}`; // print one sheet per page, no browser header/footer
  styleEl.textContent += `\n.${scope} > .odx-content{padding:${page.firstTop ?? page.top}px ${page.right}px 0 ${page.left}px;--odx-page-left:${page.left}px;--odx-page-right:${page.right}px}`;
  if (session.pageColor) sheet.style.setProperty('--odx-page-color', `#${session.pageColor}`);
  mount.replaceChildren(sheet);

  let pages = null;
  let ime = null; // the text an input method is composing with Track Changes on, { from, to }
  const view = new EditorView(sheet, {
    state,
    attributes: { spellcheck: 'true', class: 'odx-content', role: 'textbox', 'aria-multiline': 'true', 'aria-label': 'Document', lang: session.lang || 'en-US' },
    dispatchTransaction(tr) {
      // (`this`, the view: a shared document's binding dispatches while the view is still being made, and, its updates
      // deferred, after the view was destroyed for another document: nothing to do then.)
      const v = this;
      if (v.isDestroyed) return;
      // With Track Changes on, edits are recorded as revisions (undo/redo replay them as they were).
      const track = tracking(v.state);
      // (Not others' edits arriving in a shared document: they were tracked, if at all, where they were made.)
      const tracked = track.on && tr.docChanged && !tr.getMeta('odx-untracked') && !tr.getMeta('history$') && !(collab && (collab.remote(tr) || tr.getMeta('odx-remote')));
      // (While an input method composes (Chinese, Japanese...), its text is left as typed: marking it then redraws the
      // text under the composition, which leaves its first letter behind. It is marked inserted when the composition ends.)
      if (ime) ime = { from: tr.mapping.map(ime.from, -1), to: tr.mapping.map(ime.to, 1) };
      if (tracked && (v.composing || ime)) {
        tr.mapping.maps.forEach((m, i) => m.forEach((_a, _b, s, e) => {
          const rest = tr.mapping.slice(i + 1);
          const a = rest.map(s, -1), z = rest.map(e, 1);
          ime = ime ? { from: Math.min(ime.from, a), to: Math.max(ime.to, z) } : { from: a, to: z };
        }));
      } else if (tracked) tr = trackTransaction(v.state, tr, track.author);
      const sectPr = v.state.doc.attrs.sectPr;
      v.updateState(v.state.apply(tr));
      // (A shared document's own attributes, its last section's page setup and header references, go to the room.)
      if (v.odx?.docAttrs && !tr.getMeta('odx-remote') && v.state.doc.attrs.sectPr !== sectPr) v.odx.docAttrs.push(v.state.doc.attrs);
      session.trackRevisions = tracking(v.state).on;
      if (tr.docChanged && pages) pages.update();
      onChange && v.odx && onChange(v, tr);
    },
  });
  // The composed text, once the input method is done: taken out untracked and typed again, so it is a tracked insertion.
  const imeDone = () => {
    if (!ime || view.isDestroyed) return;
    if (view.composing) { setTimeout(imeDone, 80); return; }
    const { from, to } = ime;
    ime = null;
    const text = view.state.doc.textBetween(Math.max(0, from), Math.min(view.state.doc.content.size, to), '');
    if (!text || !tracking(view.state).on) return;
    view.dispatch(view.state.tr.delete(from, to).setMeta('odx-untracked', true).setMeta('addToHistory', false));
    view.dispatch(view.state.tr.insertText(text, from));
  };
  view.dom.addEventListener('compositionend', () => setTimeout(imeDone, 80));
  pages = createPages(view, session, { sheet, scope });
  view.odx = {
    session,
    author,
    main: view,
    // A header, footer or note being edited: its text into the document so far (for a save), or all of them closed.
    flushStories() { flushStory(); flushNote(); },
    closeStories() { closeStory(); closeNote(); closeTextBox(); },
    pages,
    /** Styles added to the document since it opened (session.addBuiltInStyles): their CSS too. */
    restyle() { styleEl.textContent += `\n${session.css(`.${scope}`)}\n${tableStylesCss(`.${scope}`, session.styles, session.theme, { oldSizes: !!session.oldTableSizes })}`; },
    // The pages' zoom (zoom.js); onZoom, set by the host, hears each change.
    zoom: createZoom(mount, sheet, page, (z) => view.odx?.onZoom?.(z)),
    collab,
    undo: doUndo,
    redo: doRedo,
    destroy: () => {
      view.odx.zoom.destroy();
      view.odx.comments?.stop();
      view.odx.stories?.stop();
      view.odx.docAttrs?.stop();
      view.odx.definitions?.stop();
      closeStory();
      closeNote();
      pages.destroy();
      view.destroy();
      styleEl.remove();
    },
  };
  // Shared comments (src/collab.js): the room's, ids nobody else will pick, others' changes redrawn here.
  if (collab) {
    session.uniqueIds = true;
    view.odx.comments = collab.bindComments(session, () => refreshComments(view)); // (its push then finds nothing new)
    view.odx.stories = collab.bindStories(session, { fromJSON: (json) => schema.nodeFromJSON(json), onRemote: () => pages.restory() });
    view.odx.docAttrs = collab.bindDoc(view);
    // Styles and lists another editor added: their CSS, then the paragraphs' list labels and spacing worked out again
    // (after the shared body has taken the same update, as bindDoc's), and the pages measured again.
    view.odx.definitions = collab.bindDefinitions(session, () => {
      paras.reset();
      view.odx.restyle();
      setTimeout(() => {
        if (view.isDestroyed) return;
        const tr = paras.sync(view.state.tr);
        if (tr.docChanged) view.dispatch(tr.setMeta('odx-remote', true).setMeta('addToHistory', false));
        pages.restory();
      });
    });
  }
  hfDoubleClick(view);
  noteClicks(view);
  return view;
}

/** Save the current editor contents to .docx bytes. */
export function saveEditor(view) {
  return view.odx.session.save(view.state.doc);
}
