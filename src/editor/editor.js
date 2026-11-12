// The editor: a ProseMirror view laid out as a Word page, with Word-style keys
// and list numbering computed from numbering.xml.
import { commentsPlugin, addComment, refresh as refreshComments } from './comments.js';
import { hfDoubleClick, closeStory, flushStory } from './hfedit.js';
import { formsPlugin } from './fields.js';
import { noteClicks, closeNote, flushNote, insertNote } from './notes.js';
import { EditorState, Plugin } from 'prosemirror-state';
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
import { columnsIn } from './sections.js';
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
    } else this.table.classList.remove('odx-autofit');
    const place = (node.attrs.deco && node.attrs.deco.style) || '';
    if (this.dom.getAttribute('style') !== place) this.dom.setAttribute('style', place);
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
    cells.push({ rel, node, row: rect.top, col: rect.left, rowspan: rect.bottom - rect.top, colspan: rect.right - rect.left, tcPr: node.attrs.tcPr });
  }
  const key = table.attrs.tblPr + '|' + map.width + 'x' + map.height + '|' + cells.map((c) => `${c.row},${c.col},${c.rowspan},${c.colspan},${c.tcPr.map((x) => x[1]).join('')}`).join(';');
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
    const visit = (parent, offset) => {
      let prev = null;
      let floatGap = null; // px: the space after the paragraph before a floating table, when this block follows that table
      parent.forEach((node, rel, index) => { // (index: a top-level block's, for sides and collapse; doc.childBefore made this quadratic)
        const pos = offset + rel;
        if (node.type !== schema.nodes.paragraph) {
          // A bookmark or other marker between two paragraphs shows nothing, so they still count as neighbours.
          if (node.type === schema.nodes.opaque_block && (ZERO_WIDTH.has(node.attrs.kind) || node.attrs.kind === 'proofErr')) return;
          const before = prev;
          prev = null;
          floatGap = null;
          if (node.type === schema.nodes.table) {
            const place = tablePlacementCss(session.styles, node.attrs.tblPr, session.compatMode, parseXml);
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
            if (sd && !/(?:^|;)float:/.test(place || '')) style = [style, `margin-inline-start:${sd[0]}px`, `margin-inline-end:calc(${sd[1]}px${sd[2] ? ' + var(--odx-coln, 0px)' : ''})`].filter(Boolean).join(';');
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
          if (auto(node, 'afterAuto')) style.push('margin-bottom:5pt');
        } else {
          if (rel === 0 && auto(node, 'beforeAuto')) style.push('margin-top:0');
          if (prev && node.attrs.numId && node.attrs.numId !== '0' && prev.node.attrs.numId === node.attrs.numId && auto(node, 'beforeAuto') && auto(prev.node, 'afterAuto')) {
            style.push('margin-top:0');
            prev.style.push('margin-bottom:0');
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
          if (prev.cs) prev.style.push('margin-bottom:0');
        }
        const cls = [attrs.class, ...changeClasses(node), parent === doc && collapse.has(index) && 'odx-collapse'].filter(Boolean);
        if (cls.length) attrs.class = cls.join(' ');
        headingRole(node, session, attrs);
        // A paragraph in a frame (frames.js): at its place, the text beside it; one after another in the same frame below it.
        const fr = parent === doc && frameOf(node);
        if (fr) {
          const f = frameStyle(fr, index > 0 && frameOf(doc.child(index - 1)) === fr, page);
          style.push(...f.style);
          Object.assign(attrs, f.attrs);
        }
        const sd = !fr && parent === doc && sides.get(index);
        if (sd) {
          // (Added to the paragraph's own indents, the last margin it sets.)
          const own = (k) => { const v = [...style].reverse().find((x) => x.startsWith(`${k}:`)); return v ? v.slice(k.length + 1).trim() : '0pt'; };
          // (The section's sides are the page's left and right: a right-to-left paragraph's start is its right.)
          const rtl = (node.attrs.view.bidi ?? session.styleProp(node.attrs.styleId, 'bidi')) === true;
          const [ds, de] = rtl ? [sd[1], sd[0]] : sd;
          // (Columns of their own widths: laid out the widest wide, less --odx-coln on the right, which pages.js sets
          // for the column each block lands in.)
          const coln = sd[2] ? ' + var(--odx-coln, 0px)' : '';
          style.push(`margin-inline-start:calc(${own('margin-inline-start')} + ${ds}px${rtl ? coln : ''})`, `margin-inline-end:calc(${own('margin-inline-end')} + ${de}px${rtl ? '' : coln})`);
        }
        prev = { node, pos, attrs, style, cs, bdr };
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
  const sync = (tr) => {
    for (const { pos, node, deco } of compute(tr.doc)) {
      if (JSON.stringify(deco) !== JSON.stringify(node.attrs.deco)) tr.setNodeMarkup(pos, null, { ...node.attrs, deco });
    }
    return tr;
  };
  // Stored as node attrs rather than view decorations: ProseMirror checks every
  // top-level decoration against every top-level node, which is quadratic on long documents.
  /** Same as sync, for a whole document in one pass (used when opening). */
  const initial = (doc) => {
    const changes = new Map();
    for (const { pos, node, deco } of compute(doc)) if (JSON.stringify(deco) !== JSON.stringify(node.attrs.deco)) changes.set(pos, deco);
    if (!changes.size) return doc;
    const rebuild = (node, offset) => {
      const kids = [];
      node.forEach((child, rel) => {
        const pos = offset + rel;
        const inner = child.isLeaf || child.isTextblock ? child : rebuild(child, pos + 1);
        kids.push(changes.has(pos) ? child.type.create({ ...child.attrs, deco: changes.get(pos) }, inner === child ? child.content : inner.content, child.marks) : inner);
      });
      return node.copy(Fragment.from(kids));
    };
    return rebuild(doc, 0);
  };
  return {
    sync,
    initial,
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
 * Pasted paragraphs' built-in styles by the document's own ids (Word's English ids, as a pasted <h1> gets, are another
 * language's in a document from a localised Word: Heading1 is "Ttulo1" from Spanish Word).
 */
function localStylesPlugin(session) {
  const map = (frag) => Fragment.fromArray(frag.content.map((n) => {
    if (n.isLeaf) return n;
    const id = n.type === schema.nodes.paragraph && n.attrs.styleId;
    const local = id && session.builtInId(id);
    return n.type.create(local && local !== id ? { ...n.attrs, styleId: local } : n.attrs, map(n.content), n.marks);
  }));
  return new Plugin({ props: { transformPasted: (slice) => new Slice(map(slice.content), slice.openStart, slice.openEnd) } });
}

export function createEditor(mount, session, { onChange, author = 'Author', collab = null } = {}) {
  const scope = `odx-doc-${++scopeSeq}`;
  const styleEl = document.createElement('style');
  styleEl.textContent = session.css(`.${scope}`) + '\n' + tableStylesCss(`.${scope}`, session.styles, session.theme) + '\n' + session.picBulletCss(`.${scope}`);
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
    plugins: [...formsPlugin(), chartPlugin(), shapeKeysPlugin() /* (before the keymap: its Enter would replace a selected chart, its arrows move the caret off a shape) */, ...(collab ? collab.plugins() : [history()]), autocorrectPlugin, keys, keymap(baseKeymap), columnResizing({ View: StyledTableView }), tableEditing(), paras.plugin, activeBlock, pagePushPlugin, textSpansPlugin, justifyPlugin, lineFixPlugin, widthsPlugin, fitTextPlugin, joinsPlugin, marksPlugin, csFormatPlugin, pastePictures(session), pictureLinks, imageResizePlugin(), textBoxPlugin(), trackPlugin({ on: session.trackRevisions, author }), commentsPlugin(session), protectPlugin(), localStylesPlugin(session), hyphenatePlugin(session), mailMergePlugin(session)],
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
    restyle() { styleEl.textContent += `\n${session.css(`.${scope}`)}\n${tableStylesCss(`.${scope}`, session.styles, session.theme)}`; },
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
  }
  hfDoubleClick(view);
  noteClicks(view);
  return view;
}

/** Save the current editor contents to .docx bytes. */
export function saveEditor(view) {
  return view.odx.session.save(view.state.doc);
}
