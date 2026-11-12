// Print-layout view: splits the continuous editor column into Word-style pages.
//
// The editable content stays one ProseMirror column. After each change we
// measure the top-level blocks, decide which block starts each page, and push
// those blocks down with a generated stylesheet (so ProseMirror's DOM is never
// touched). A layer behind the text draws the page sheets, headers and footers.
//
// Paragraphs split line by line and tables row by row (rows inside their cells),
// with spacers inserted as view decorations. Each section's top and bottom
// margins apply from the page it starts on; page size and side margins are the
// last section's.
import { DOMSerializer } from 'prosemirror-model';
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { schema } from './schema.js';
import { layoutTabs, layoutTabsIn } from './tabs.js';
import { fitRowspans, clipExactRows } from './rowfit.js';
import { storyDrawer, hfNodes, styleStoryTables } from './hfstory.js';
import { printCopies } from './printcopies.js';
import { sectionBreaks, columnsIn, sectionXmls } from './sections.js';
import { drawLineNumbers } from './linenumbers.js';
import { fontsSettled } from './fontcheck.js';
import { referenceValue, updateFields } from './refs.js';
import { lineHeightOf } from './util.js';
import { wrapText } from './textspans.js';
import { drawBalloons } from './balloons.js';
import { noteRefs, noteLabel } from './notes.js';
import { applySqueeze } from './justify.js';
import { applyLineFix } from './lhfix.js';
import { applyWidths } from './widths.js';
import { applyFitText } from './fittext.js';
import { fitAutofitColumns } from './tablefit.js';
import { baselineCss, cellShifts, hfShifts, shiftFor } from './baseline.js';
import { buildPrintLines, removePrintLines, engineOn } from '../engine/printlines.js';
import { twipsToPx } from '../ooxml/units.js';

// Spacers that push a paragraph's remaining lines onto the next page. They
// are view decorations (not document content), set by the layout pass below.
const pushKey = new PluginKey('odx-page-push');
export const pagePushPlugin = new Plugin({
  key: pushKey,
  state: {
    init: () => DecorationSet.empty,
    apply(tr, set) {
      const next = tr.getMeta(pushKey);
      return next ? DecorationSet.create(tr.doc, next) : set.map(tr.mapping, tr.doc);
    },
  },
  props: { decorations: (state) => pushKey.getState(state) },
});
const spacer = (h) => () => {
  const el = document.createElement('span');
  el.className = 'odx-push';
  el.contentEditable = 'false';
  el.style.height = `${h}px`;
  return el;
};

const GAP = 24; // px between sheets
const TIE = 0.33; // px an exactly spaced line must end above the page's bottom margin (splitParagraph)

export function createPages(view, session, { sheet, scope }) {
  const page = session.page();
  let gap = GAP;
  // Pages sit a whole px apart: Edge prints each sheet a whole px tall, so a page of a fractional height (A4's 16839
  // twips, 1122.6px) pitched at that height printed each next page's content 0.4px higher (sample-docx's page 5,
  // 1.3px). The room on a page stays its exact height (pageH): rounded, doc 24's page 3 took a line Word moves on.
  const step = (h) => Math.round(h);
  let pitch = step(page.height) + gap;
  // Top and bottom margins per page, as [first page, margins] in page order
  // (filled in by the layout pass as sections start).
  let margins = [[0, page]];
  let sections = [[0, 0]]; // [first page, section index] for each section start, for editing its headers
  // `fresh`: the section starts on that page (not continued from the page before), so it gets its first-page header.
  const setMargins = (from, v, fresh) => {
    while (margins.length > 1 && margins[margins.length - 1][0] >= from) margins.pop();
    if (margins[margins.length - 1][1] !== v) margins.push([from, v, fresh]);
  };
  // A section's first page can have its own (first-page) header and footer, so its own margins.
  const marginsAt = (p) => {
    let e = margins[0];
    for (const m of margins) if (m[0] <= p) e = m; else break;
    return e[2] && e[0] === p && e[1].first ? e[1].first : e[1];
  };
  // Sections can have their own page height, so pages don't all sit one pitch apart.
  const pageH = (p) => {
    let e = margins[0];
    for (const m of margins) if (m[0] <= p) e = m; else break;
    return e[1].height || page.height;
  };
  // A page's number and its format: one more than the page before, or the section's start where it restarts.
  const numberOf = (p) => {
    let num = 0, fmt;
    for (let k = 0; k <= p; k++) {
      let st = margins[0];
      for (const m of margins) if (m[0] <= k) st = m; else break;
      num = st[0] === k && (st[2] || k === 0) && st[1].numStart != null ? st[1].numStart : num + 1;
      fmt = st[1].numFmt;
    }
    return { n: num, fmt };
  };
  // And their own page width (a landscape section among portrait ones).
  const pageW = (p) => {
    let e = margins[0];
    for (const m of margins) if (m[0] <= p) e = m; else break;
    return e[1].width || page.width;
  };
  const pageTop = (p) => {
    let y = 0;
    for (let k = 0; k < margins.length; k++) {
      const from = margins[k][0];
      const to = k + 1 < margins.length ? Math.min(margins[k + 1][0], p) : p;
      if (to > from) y += (to - from) * (step(margins[k][1].height || page.height) + gap);
    }
    return y;
  };
  const pageOf = (y) => {
    let p = 0;
    for (let k = 0; k < margins.length; k++) {
      const h = step(margins[k][1].height || page.height) + gap;
      const top = pageTop(margins[k][0]);
      const end = k + 1 < margins.length ? margins[k + 1][0] : Infinity;
      p = margins[k][0] + Math.max(0, Math.floor((y - top) / h));
      if (p < end) return p;
    }
    return p;
  };
  const contentTop = (p) => colTops.get(p) ?? pageTop(p) + marginsAt(p).top;
  const contentBottom = (p) => Math.min(pageTop(p) + pageH(p) - marginsAt(p).bottom - (notesH.get(p) || 0), cuts.get(p) ?? Infinity);
  // Footnotes: which go at the bottom of each page, and the room they take there.
  let notesOn = new Map(); // page -> [{ kind, id, n }]
  let notesH = new Map(); // page -> px
  let cuts = new Map(); // page -> y: a table set down near its page's foot ends the text there (pagination's band)
  const NOTE_SEP = 14; // separator line and the space around it
  const ENDNOTE_LEAD = 4; // Word's separator under the text is a line of the Normal style, a little taller than NOTE_SEP
  const noteHeights = new Map(); // note doc -> px
  const noteEl = (note, kind, n) => {
    const el = document.createElement('div');
    el.className = 'odx-note';
    el.dataset.kind = kind;
    el.dataset.id = note.id;
    el.append(serializer.serializeFragment(note.doc.content));
    styleStoryTables(session, note.doc, el, { note: true }); // (a table in a note: its columns, borders and shading, as in the body)
    wrapText(el);
    el.querySelectorAll(`[data-kind="${kind}Ref"]`).forEach((m) => { m.textContent = noteLabel(kind, n); });
    return el;
  };
  const noteBounds = new Map(); // note doc -> where each of its paragraphs ends, px from its top
  const noteHeight = (note, kind, n) => {
    if (!noteHeights.has(note.doc)) {
      const box = document.createElement('div');
      box.className = 'odx-notes';
      box.style.cssText = `visibility:hidden;top:0;left:${page.left}px;right:${page.right}px`;
      const el = noteEl(note, kind, n);
      box.append(el);
      sheet.append(box);
      const top = box.getBoundingClientRect().top;
      noteHeights.set(note.doc, box.getBoundingClientRect().height);
      // (A table's rows each end a place it can be cut too: a note holding a long table goes on on the next page from
      // the row that doesn't fit, as in Word.)
      noteBounds.set(note.doc, [...el.children].flatMap((c) => {
        const t = c.tagName === 'TABLE' ? c : c.querySelector(':scope > table');
        const rows = t && t.tBodies[0] ? [...t.tBodies[0].rows] : [];
        return (rows.length ? rows : [c]).map((e) => e.getBoundingClientRect().bottom - top);
      }).filter((b) => b > 0));
      box.remove();
    }
    return noteHeights.get(note.doc);
  };
  // The notes drawn at each page's foot: those carried on from the pages before (a note longer than the room a page has
  // for it continues at the foot of the next, as in Word), then the page's own; `skip` px of the first already shown.
  let noteFlow = new Map(); // page -> { items, skip, h }
  const NOTE_MIN_BODY = 48; // the room a page keeps for its body text (its reference's line, at least)
  const sumNotes = () => {
    notesH = new Map();
    noteFlow = new Map();
    const withNotes = [...notesOn.keys()].filter((p) => notesOn.get(p).length).sort((a, b) => a - b);
    if (!withNotes.length) return;
    const last = withNotes[withNotes.length - 1];
    let carry = [], skip = 0;
    for (let p = withNotes[0]; carry.length || p <= last; p++) {
      const items = [...carry, ...(notesOn.get(p) || [])];
      if (!items.length) continue;
      const total = items.reduce((t, r) => t + r.h, 0) - skip;
      let room = Math.max(40, pageH(p) - marginsAt(p).top - marginsAt(p).bottom - NOTE_MIN_BODY - NOTE_SEP);
      if (total <= room) { noteFlow.set(p, { items, skip, h: total }); notesH.set(p, total + NOTE_SEP); carry = []; skip = 0; continue; }
      // Too long for the page: in Word it takes all the room under the text with its reference (that text's page ends
      // there), not only what the page keeps for its body.
      const ats = (notesOn.get(p) || []).map((r) => r.at).filter((v) => v != null);
      if (ats.length) room = Math.max(40, Math.min(room, pageTop(p) + pageH(p) - marginsAt(p).bottom - Math.max(...ats) - NOTE_SEP));
      // Cut at the last paragraph end that fits (or, a paragraph taller than the room, at the room).
      let at = 0, cut = 0;
      for (const r of items) {
        for (const b of noteBounds.get(r.note?.doc) || [r.h]) if (at + b - skip <= room && at + b - skip > 0) cut = at + b - skip;
        at += r.h;
      }
      if (!cut) cut = room;
      noteFlow.set(p, { items, skip, h: cut });
      notesH.set(p, cut + NOTE_SEP);
      const rest = [];
      let acc = 0, next = 0;
      for (const r of items) {
        if (acc + r.h > skip + cut + 0.5) { if (!rest.length) next = skip + cut - acc; rest.push(r); }
        acc += r.h;
      }
      carry = rest;
      skip = next;
      if (p > last + 500) break; // (a guard: never runs away)
    }
  };
  let blockNotes = new Map(); // block index -> [{ kind, id, n, h }]
  const placed = new Map(); // note key -> page
  // (`at`: where the block holding the references ends, when known: a note too long for its page has the room below it.)
  const reserveNotes = (i, p, at = null) => {
    for (const r of blockNotes.get(i) || []) {
      const key = r.kind + r.id;
      if (placed.has(key)) continue;
      placed.set(key, p);
      r.at = at;
      if (!notesOn.has(p)) notesOn.set(p, []);
      notesOn.get(p).push(r);
      sumNotes(p);
    }
  };
  // A footnote goes on the page its reference's line does: in a paragraph split over pages, not always the paragraph's
  // first (corpus 3's c3-01: a reference in a paragraph's lines carried to page 2 has its note there, as in Word).
  const notesByLine = (b, shift0, split) => {
    const list = blockNotes.get(b.i);
    if (!list || !split.pushes.length) return;
    const base = view.dom.getBoundingClientRect().top + b.undo;
    const pushes = [...b.el.querySelectorAll('.odx-push')].map((e) => e.getBoundingClientRect());
    for (const r of list) {
      if (r.kind !== 'footnote') continue;
      const el = view.nodeDOM(r.pos);
      if (!el || el.nodeType !== 1) continue;
      const y0 = el.getBoundingClientRect().top;
      const nat = y0 - base - pushes.reduce((sum, q) => sum + (q.top + q.height / 2 <= y0 ? q.height : 0), 0);
      const to = pageOf(nat + shift0 + split.pushes.reduce((sum, x) => sum + (x.pos <= r.pos ? x.delta : 0), 0));
      const key = r.kind + r.id;
      const from = placed.get(key);
      if (from == null || from === to) continue;
      notesOn.set(from, (notesOn.get(from) || []).filter((x) => x !== r));
      sumNotes(from);
      placed.set(key, to);
      if (!notesOn.has(to)) notesOn.set(to, []);
      notesOn.get(to).push(r);
      sumNotes(to);
    }
  };
  const moveNotes = (i, to) => {
    for (const r of blockNotes.get(i) || []) {
      const key = r.kind + r.id;
      const from = placed.get(key);
      if (from == null || from === to) continue;
      notesOn.set(from, (notesOn.get(from) || []).filter((x) => x !== r));
      sumNotes(from);
      placed.delete(key);
    }
    reserveNotes(i, to);
  };
  const pageHeight = (p) => contentBottom(p) - contentTop(p);
  // Newspaper columns (w:cols) for a document that is one section: each column is laid out as its own
  // page of the flow, then moved beside the one before it. Blocks move to the next column whole.
  const cols = (() => {
    let multi = false;
    view.state.doc.forEach((n) => { if ((n.attrs.extra || []).some((e) => e[0] === 'sectPr')) multi = true; });
    return multi ? null : columnsIn(view.state.doc.attrs.sectPr, page);
  })();
  if (cols) view.dom.style.paddingRight = `${page.width - page.left - cols.width}px`;
  // Sections in newspaper columns in a document of several (corpus 2's c2-07, c3-06 to c3-08): each is laid out column
  // by column from where it starts, its blocks a column wide (editor.js sides). Blocks go to the next column whole or,
  // paragraphs, line by line as in single-section columns. A section followed by a continuous one is balanced, as in
  // Word: its last sheet's columns are cut to an even height (worked out a pass behind, `balanced`) and what follows
  // starts under the longest. Top-level indexes lo..hi; hi is Infinity for the last section.
  const colRegions = (doc) => {
    if (cols) return [];
    const ends = [];
    doc.forEach((n, _, i) => { const x = n.type === schema.nodes.paragraph && (n.attrs.extra || []).find((e) => e[0] === 'sectPr'); if (x) ends.push([i, x[1]]); });
    if (!ends.length) return [];
    const all = ends.concat([[Infinity, doc.attrs.sectPr || '']]);
    const out = [];
    all.forEach(([hi, xml], n) => {
      const c = columnsIn(xml, page);
      if (!c) return;
      const next = all[n + 1];
      out.push({ lo: n ? all[n - 1][0] + 1 : 0, hi, ...c, balance: !!next && /<w:type\b[^>]*w:val="continuous"/.test(next[1]) });
    });
    return out;
  };
  let colTops = new Map(); // flow page -> its content top, for the columns beside the first one on a columns section's first sheet
  const balanced = new Map(); // section's lo -> { sheet (of its own, from 0), h }: the column height its last sheet is cut to
  let balancedDoc = null;
  let balancePasses = 0;
  let narrowed = new Map(); // block index -> px its column (of their own widths) is narrower than the widest, as laid out
  let colsAgain = false;
  let colsPasses = 0;
  let colsRepass = false; // the coming pass is one of those
  // Splitting paragraphs line by line needs each page-crossing paragraph laid
  // out; in very long documents (content-visibility skips off-screen blocks)
  // that costs too much, so there paragraphs move to the next page whole.
  const lineSplits = !sheet.classList.contains('odx-large') && !cols;

  const layer = document.createElement('div');
  layer.className = 'odx-sheets';
  sheet.prepend(layer);
  const rules = document.createElement('style');
  document.head.appendChild(rules);

  const serializer = DOMSerializer.fromSchema(schema);
  const story = storyDrawer(session);

  // Repeated table header rows are drawn here, behind the text, in the gap the split leaves.
  const repeats = document.createElement('div');
  repeats.className = 'odx-repeats';
  view.dom.after(repeats); // (after the document: its copies of lines come after the paragraphs they copy, as in a search)
  // The gaps between sheets, drawn over the text layer, so paragraph or cell shading
  // that runs across a page break doesn't paint the gap.
  const gaps = document.createElement('div');
  gaps.className = 'odx-gaps';
  repeats.after(gaps);
  // Line numbers in the left margin (linenumbers.js), for sections that have them.
  const lineNums = document.createElement('div');
  lineNums.className = 'odx-lnnums';
  lineNums.setAttribute('aria-hidden', 'true'); // (a reference beside the lines, not part of the text: not read out)
  gaps.after(lineNums);

  /**
   * Rows that have to start a new page, Word-style: a row that doesn't fit moves
   * to the next page, below a copy of the header rows. Returns null when the
   * first row doesn't fit either (the whole table moves instead).
   */
  function splitTable(b, shift, tableFirstOnPage) {
    const { rows, headerRows } = b;
    const headH = headerRows ? rows[headerRows - 1].bottom - rows[0].top : 0;
    const pushes = new Map();
    const offs = new Map();
    const cellPushes = [];
    const bands = [];
    const closes = []; // where a row broken across pages gets its closing border
    const inRow = new Set();
    const heads = [];
    let extra = 0;
    const extraAt = [];
    for (let j = 0; j < rows.length; j++) {
      extraAt[j] = extra;
      const top = rows[j].top + shift + extra;
      const bottom = rows[j].bottom + shift + extra;
      const pg = pageOf(top);
      const brk = j > headerRows && j > 0 && b.node.child(j) && rowBreakAt(b.node, j) && top > contentTop(pg) + headH + 1;
      if (!brk && bottom + (rows[j].line || 0) <= contentBottom(pg)) continue;
      // Keep with next: the rows above that keep with this one (a paragraph in each keeps with the next) go to the next
      // page with it, as in Word; when they reach the table's first row, the table moves (sample: Heading 1 in its cells).
      if (!brk) {
        let s = j;
        while (s > headerRows && rowKeeps(b.node.child(s - 1)) && !pushes.has(s - 1) && !inRow.has(s - 1) && pageOf(rows[s - 1].top + shift + extraAt[s - 1]) === pg) s--;
        if (s < j) {
          const topS = rows[s].top + shift + extraAt[s];
          const firstS = s === 0 ? tableFirstOnPage : topS <= contentTop(pg) + (pushes.size || inRow.size ? headH : 0) + 1;
          const tall = rows[j].bottom - rows[s].top + headH > pageHeight(pg + 1);
          if (!firstS && !tall && s <= headerRows && !tableFirstOnPage) return null;
          if (!firstS && !tall && s > headerRows) {
            const d = contentTop(pg + 1) + headH - topS;
            pushes.set(s, d);
            offs.set(s, extraAt[s]);
            if (!headerRows) closes.push(contentTop(pg + 1));
            heads.push(contentTop(pg + 1));
            extra = extraAt[s] + d;
            j = s;
            continue;
          }
        }
      }
      // Word breaks a row inside its cells, line by line, unless the row says it can't split, or (sample S072) the room left
      // on the page is less than the row's minimum height (w:trHeight "at least"): then the row moves whole.
      const tr = b.node.child(j)?.attrs.trPr || '';
      const minM = /<w:trHeight\b[^>]*w:val="(\d+(?:\.\d+)?)"/.exec(tr);
      const minH = minM && !/<w:trHeight\b[^>]*w:hRule="(?:exact|auto)"/.test(tr) ? twipsToPx(minM[1]) : 0;
      const short = minH && contentBottom(pg) - top < minH && minH <= pageHeight(pg + 1);
      const rs = !brk && j >= headerRows && lineSplits && !short && !(j > headerRows && j < rows.length - 1 && rowKeeps(b.node.child(j))) && splitRow(b, j, shift + extra, headH);
      if (rs) {
        cellPushes.push(...rs.pushes);
        for (const [q, cut] of rs.crossings) {
          bands.push([cut - (b.tableTop + shift), contentTop(q + 1) + headH - (b.tableTop + shift)]);
          heads.push(contentTop(q + 1));
          closes.push(cut);
          if (!headerRows) closes.push(contentTop(q + 1) + 1); // and its top border where it carries on
        }
        inRow.add(j);
        extra += rs.growth;
        continue;
      }
      const first = j === 0 ? tableFirstOnPage : top <= contentTop(pg) + (pushes.size || inRow.size ? headH : 0) + 1;
      if (first || rows[j].bottom - rows[j].top > pageHeight(pg + 1)) continue; // taller than a page: let it run over
      if (j <= headerRows) return null;
      const d = contentTop(pg + 1) + headH - top;
      pushes.set(j, d);
      offs.set(j, extra); // everything above this row has moved down by this much (pushed and split rows)
      if (!headerRows) closes.push(contentTop(pg + 1)); // the moved row's top border, where it now starts
      heads.push(contentTop(pg + 1));
      extra += d;
    }
    return pushes.size || inRow.size ? { rows: pushes, offs, heads, total: extra, cellPushes, bands, inRow, closes } : null;
  }

  const cellWidow = (p) => {
    try { return widowControl(view.state.doc.resolve(view.posAtDOM(p, 0)).parent); } catch { return false; }
  };
  // A row can't split when it says so, or when every paragraph of text in it keeps with the next (Word moves it whole).
  const keepsWhole = (row) => {
    let any = false;
    let all = true;
    row.descendants((n) => {
      if (n.type !== schema.nodes.paragraph) return true;
      if (n.textContent.trim()) { any = true; if (!keep(n)) all = false; }
      return false;
    });
    return any && all;
  };
  const cantSplit = (row) => /<w:cantSplit(?![^>]*w:val="(?:0|false|off)")/.test(row.attrs.trPr || '') || keepsWhole(row);

  // Our spacers in a cell (not in the tables in it), and how much they have made a row of a table in a cell taller:
  // what's below them is measured less that, at its natural place.
  const spacersIn = (td) => [...td.querySelectorAll('.odx-push')].filter((e) => e.closest('td') === td).map((e) => e.getBoundingClientRect());
  const grown = (ntr) => {
    if (!ntr.querySelector('.odx-push')) return 0;
    let nat = -Infinity;
    for (const ntd of ntr.cells) {
      const last = ntd.lastElementChild;
      if (!last) continue;
      const lb = last.getBoundingClientRect().bottom;
      nat = Math.max(nat, lb - undoIn(ntd, lb) + (parseFloat(getComputedStyle(ntd).paddingBottom) || 0));
    }
    return Number.isFinite(nat) ? Math.max(0, ntr.getBoundingClientRect().bottom - nat) : 0;
  };
  const tablesIn = (td) => [...td.children].map((e) => (e.tagName === 'TABLE' ? e : e.classList.contains('tableWrapper') ? e.querySelector(':scope > table') : null)).filter(Boolean);
  function undoIn(td, y) {
    let u = spacersIn(td).reduce((sum, r) => sum + (r.bottom <= y + 0.5 ? r.height : 0), 0);
    for (const t of tablesIn(td)) for (const ntr of t.rows) {
      if (ntr.getBoundingClientRect().bottom > y + 0.5) break;
      u += grown(ntr);
    }
    return u;
  }
  /**
   * Lines of each cell of row j that move to the next page(s), measured like
   * paragraphs. `rowShift` is how far the row has moved from its natural place.
   * Returns null when the row should move whole (nothing of it fits, or it can't split).
   */
  function splitRow(b, j, rowShift, headH) {
    if (cantSplit(b.node.child(j))) return null;
    const tr = b.table.tBodies[0].rows[j];
    const base = view.dom.getBoundingClientRect().top + b.rows[j].undo;
    const pushes = [];
    const crossings = new Map(); // page -> bottom of the row's part that stays on it
    let keeps = false;
    let newBottom = -Infinity;
    // The lines of a run of paragraphs in one cell from `extra0` on: which move to the next page, as a paragraph's.
    function runLines(lines, extra0, padTop, padBottom, line, seq0, out) {
      let extra = extra0;
      for (let n = 0; n < lines.length; n++) {
        let top = lines[n].top + rowShift + extra;
        const bottom = lines[n].bottom + rowShift + extra;
        const pg = pageOf(top);
        if (bottom + padBottom + line <= contentBottom(pg)) continue; // Word keeps the cell's bottom margin above the page break
        // Widow/orphan control inside the cell: never leave one line of a paragraph behind, or carry one over alone.
        const own = lines.filter((l) => l.p === lines[n].p);
        const k = own.indexOf(lines[n]);
        if (k > 0 && !lines[n].empty && cellWidow(lines[n].p)) {
          let keep = k; // lines of this paragraph left on the page
          if (own.length - keep === 1 && own.length > 2) keep -= 1;
          if (keep === 1) keep = 0;
          n -= k - keep;
          top = lines[n].top + rowShift + extra;
        }
        const whole = lines[n].empty || lines[n] === lines.find((l) => l.p === lines[n].p);
        const lead = whole || Math.abs(lines[n].lead) > 10 ? 0 : lines[n].lead;
        const delta = Math.round((contentTop(pg + 1) + headH + padTop + lead - top) * 10) / 10;
        if (delta <= 0) continue;
        const pos = whole ? view.posAtDOM(lines[n].p, 0) : lineStart(lines[n].p, lines[n].screenTop);
        if (pos == null) return false;
        pushes.push({ pos, delta });
        const kept = n > 0 && pageOf(lines[n - 1].top + rowShift + extra) === pg ? lines[n - 1].bottom + rowShift + extra + padBottom : -Infinity;
        crossings.set(pg, Math.max(crossings.get(pg) ?? -Infinity, kept));
        if (seq0 + n === 0) out.atStart = true;
        else out.kept = true;
        extra += delta;
      }
      if (lines.length) out.end = Math.max(out.end, lines[lines.length - 1].bottom + rowShift + extra + padBottom);
      return extra;
    }
    // One cell from `extra0` on (`undoOut`: the natural offset the cells round it give a place in it). Its paragraphs line
    // by line; a table in it row by row, each of its cells as one of these, the row as tall as its tallest (Word breaks a
    // table in a cell across pages too). false: the row has to move whole.
    function plan(td, undoOut, extra0, padTop0, padBottom0, line, seq0, out) {
      // Cells centred or bottom-aligned in the row would shift as lines move, so such a row moves whole.
      if (!/^(top|baseline)$/.test(getComputedStyle(td).verticalAlign)) return false;
      const cs = getComputedStyle(td);
      const padTop = padTop0 + (parseFloat(td.style.getPropertyValue('--odx-pt')) || parseFloat(cs.paddingTop) || 0);
      const padBottom = padBottom0 + (parseFloat(cs.paddingBottom) || 0);
      const undoAt = (y) => undoOut(y) + undoIn(td, y);
      let extra = extra0;
      let seq = seq0;
      let lines = [];
      const flush = () => {
        const e = runLines(lines, extra, padTop, padBottom, line, seq, out);
        if (e === false) return false;
        extra = e;
        seq += lines.length;
        lines = [];
        return true;
      };
      for (const el of td.children) {
        if (el.tagName === 'P') {
          const r = el.getBoundingClientRect();
          const undo = b.rows[j].undo + undoAt(r.top);
          const top = r.top - base - undoAt(r.top);
          const bottom = r.bottom - base - undoAt(r.bottom);
          const pl = paragraphLines({ el, top, bottom, undo });
          // A paragraph's last line also carries its bottom border and padding (a signature rule, say), which has to fit too.
          // Where a first line's text sits against the top of its paragraph's content (the text box can start above the
          // line when the line is shorter than the font): a line carried to the next page sits that way too.
          const ps = getComputedStyle(el);
          const lead = pl ? pl[0].screenTop - (r.top + parseFloat(ps.paddingTop) + parseFloat(ps.borderTopWidth)) : 0;
          if (pl) pl.forEach((l, k) => lines.push(k === pl.length - 1 ? { ...l, p: el, lead, bottom: Math.max(l.bottom, bottom) } : { ...l, p: el, lead }));
          else lines.push({ top, bottom, p: el, empty: true });
          continue;
        }
        const t = el.tagName === 'TABLE' ? el : el.classList.contains('tableWrapper') ? el.querySelector(':scope > table') : null;
        if (!t || !t.tBodies[0] || !flush()) return false; // anything else in the cell: move the row whole
        for (const ntr of t.tBodies[0].rows) {
          if (/<w:cantSplit/.test(ntr.dataset?.trpr || '')) return false;
          const start = extra;
          let most = start;
          for (const ntd of ntr.cells) {
            if (ntd.hasAttribute('data-odx-phantom')) continue;
            const e = plan(ntd, undoAt, start, padTop, padBottom, 0, seq, out);
            if (e === false) return false;
            most = Math.max(most, e);
          }
          extra = most;
          seq += 1;
        }
      }
      if (!flush()) return false;
      return extra;
    }
    for (const td of tr.cells) {
      if (td.hasAttribute('data-odx-phantom')) continue; // an empty grid slot, nothing to move
      const line = parseFloat(td.style.getPropertyValue('--odx-bdb')) || 0; // the line under the row, the next row's: it has to fit too
      const out = { atStart: false, kept: false, end: -Infinity };
      if (plan(td, () => 0, 0, 0, 0, line, 0, out) === false) return null;
      // The row splits only when a cell that breaks keeps some of its lines here (as in Word);
      // if one would have to move whole, so does the row.
      if (out.atStart) return null;
      if (out.kept) keeps = true;
      if (!Number.isFinite(out.end)) continue;
      for (const [q, v] of crossings) if (pageOf(out.end - 1) === q) crossings.set(q, Math.max(v, out.end)); // cells that end on that page
      newBottom = Math.max(newBottom, out.end);
    }
    if (!keeps || !pushes.length) return null;
    return { pushes, crossings: [...crossings].map(([q, v]) => [q, Number.isFinite(v) ? Math.min(v, contentBottom(q)) : b.rows[j].top + rowShift]), growth: Math.max(0, newBottom - (b.rows[j].bottom + rowShift)) };
  }

  const widowControl = (node) => node.attrs.view.widowControl ?? session.styleProp(node.attrs.styleId, 'widowControl') ?? session.styles.defaults.pPr.widowControl ?? true; // on when nothing says (Word's own default, though the OOXML spec says off)
  const keepLines = (node) => node.attrs.view.keepLines ?? session.styleProp(node.attrs.styleId, 'keepLines');

  /**
   * The paragraph's lines in natural (unpaginated) coordinates: line boxes
   * estimated from the text fragments, with our own spacers taken out.
   */
  function paragraphLines(b) {
    const base = view.dom.getBoundingClientRect().top + b.undo;
    const pushes = [...b.el.querySelectorAll('.odx-push')].map((e) => e.getBoundingClientRect());
    const natural = (y) => y - base - pushes.reduce((sum, r) => sum + (r.top + r.height / 2 <= y ? r.height : 0), 0);
    const frags = [];
    const walker = document.createTreeWalker(b.el, NodeFilter.SHOW_TEXT);
    const range = document.createRange();
    for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      if (t.parentElement.closest('.odx-push')) continue;
      range.selectNodeContents(t);
      for (const r of range.getClientRects()) if (r.height > 0) frags.push(r);
    }
    b.el.querySelectorAll('img:not(.ProseMirror-separator), br:not(.ProseMirror-trailingBreak)').forEach((e) => { const r = e.getBoundingClientRect(); if (r.height > 0) frags.push(r); });
    // A paragraph ending in a line break has an empty last line (drawn by the editor's own break): Word counts it
    // for widow control (three lines and the break split two and two).
    const trailing = frags.length ? b.el.querySelector('br.ProseMirror-trailingBreak') : null;
    let trailRect = null;
    if (trailing) { const r = trailing.getBoundingClientRect(); if (r.height > 0) { frags.push(r); trailRect = r; } }
    // (Drawn by the paragraph's ::after in place of that break, editor.css odx-endbreak: the line at its content's foot.)
    if (frags.length && !trailRect && b.el.classList.contains('odx-endbreak')) {
      const cs = getComputedStyle(b.el);
      const er = b.el.getBoundingClientRect();
      const bottom = er.bottom - (parseFloat(cs.paddingBottom) || 0) - (parseFloat(cs.borderBottomWidth) || 0);
      const h = parseFloat(getComputedStyle(b.el, '::after').lineHeight) || 0;
      if (h > 0) { trailRect = { top: bottom - h, bottom, left: er.left, height: h }; frags.push(trailRect); }
    }
    frags.sort((x, y) => x.top - y.top);
    const lines = [];
    for (const r of frags) {
      const last = lines[lines.length - 1];
      if (last && r.top < last.fragBottom - 2) {
        last.fragTop = Math.min(last.fragTop, r.top);
        last.fragBottom = Math.max(last.fragBottom, r.bottom);
        last.left = Math.min(last.left, r.left);
      } else lines.push({ fragTop: r.top, fragBottom: r.bottom, left: r.left });
    }
    if (!lines.length) return null;
    const breaks = [...b.el.querySelectorAll('.odx-page-break')].map((e) => e.getBoundingClientRect().bottom);
    const blockTop = b.top;
    const blockBottom = b.bottom;
    return lines.map((l, n) => ({
      // Word puts a line's extra spacing below its text, so a line runs from its text top to the next line's
      top: n === 0 ? blockTop : natural(l.fragTop),
      bottom: n === lines.length - 1 ? blockBottom : natural(lines[n + 1].fragTop),
      screenTop: l.fragTop,
      left: l.left,
      forced: n > 0 && breaks.some((y) => y <= l.fragTop + 1 && y > lines[n - 1].fragTop),
      // (The empty line after a closing break: it starts at the paragraph's end.)
      blank: n === lines.length - 1 && trailRect != null && Math.abs(l.fragTop - trailRect.top) < 1,
    }));
  }

  /**
   * Document position where the line whose text starts at screen y `top`
   * begins. Found from the DOM (not posAtCoords, which only sees what's on screen).
   */
  function lineStart(el, top) {
    const onLine = (r) => r && r.bottom > top + 1 && r.top < top + 2;
    const range = document.createRange();
    const skipPush = { acceptNode: (n) => (n.nodeType === 1 && n.classList.contains('odx-push') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) };
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, skipPush);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      if (n.nodeType === 1) {
        if ((n.tagName === 'IMG' || n.tagName === 'BR') && onLine(n.getBoundingClientRect())) return view.posAtDOM(n.parentNode, [...n.parentNode.childNodes].indexOf(n));
        continue;
      }
      range.selectNodeContents(n);
      const rects = [...range.getClientRects()];
      if (!rects.some((r) => r.top >= top - 2 && r.top < top + 2) && !rects.some(onLine)) continue;
      // first character on that line (character tops only grow through a text node)
      let lo = 0;
      let hi = n.length - 1;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        range.setStart(n, mid);
        range.setEnd(n, mid + 1);
        const r = range.getClientRects();
        const t = r.length ? r[r.length - 1].top : -Infinity;
        if (t >= top - 2) hi = mid; else lo = mid + 1;
      }
      return view.posAtDOM(n, lo);
    }
    // (The empty line after a closing break, drawn by ::after: it starts where the paragraph's content ends.)
    const tb = el.classList.contains('odx-endbreak') && el.querySelector(':scope > br.ProseMirror-trailingBreak');
    if (tb) return view.posAtDOM(el, [...el.childNodes].indexOf(tb));
    return null;
  }

  /**
   * Lines that must start a new page: after a page break, or where the page
   * runs out (with Word's widow/orphan control and keep-lines-together).
   * Returns null when the paragraph should move to the next page whole.
   */
  function leadingBreak(el) {
    const br = el.querySelector('.odx-page-break');
    if (!br) return false;
    const range = document.createRange();
    range.setStart(el, 0);
    range.setEndBefore(br);
    return !range.toString().trim();
  }

  /**
   * How far a paragraph's last line may hang below the page's text area: Word places the extra space of
   * "multiple" line spacing below the text and doesn't count it when fitting a line (1px for rounding).
   */
  /** A table set down at a place on its page (styles.js): { y, rel, above, below } in px, or null. */
  function floatAt(b) {
    if (!b.rows) return null;
    const t = b.el.matches('table') ? b.el : b.el.querySelector('table');
    const cs = t && getComputedStyle(t);
    const y = cs && cs.getPropertyValue('--odx-fy').trim();
    if (!y) return null;
    return { y: parseFloat(y), rel: cs.getPropertyValue('--odx-fyrel').trim(), above: parseFloat(cs.getPropertyValue('--odx-fyt')) || 0, below: parseFloat(cs.getPropertyValue('--odx-fyb')) || 0 };
  }

  /** The room a line must leave above the bottom margin: TIE for exact and single spacing (doc 14's single 9pt line
   * under a table, on the margin in Word's own layout, moved), none under multiple spacing's extra. */
  function tie(el) {
    const cs = getComputedStyle(el);
    const run = parseFloat(cs.getPropertyValue('--odx-runmult'));
    const mult = parseFloat(cs.getPropertyValue('--odx-mult')) || 1;
    return run === 0 || mult === 1 ? TIE : 0;
  }
  function lineSlack(el) {
    const cs = getComputedStyle(el);
    const mult = parseFloat(cs.getPropertyValue('--odx-mult')) || 1;
    const lh = parseFloat(lineHeightOf(cs));
    return 1 + (mult > 1 && lh ? lh * (1 - 1 / mult) : 0);
  }

  /** Where a paragraph breaks at its column's bottom: { cut (px below its top), height, delta (for what follows) } or null. */
  function colSplit(b, shift, firstOnPage, p) {
    const lines = paragraphLines(b);
    if (!lines || lines.length < 2 || keepLines(b.node)) return null;
    const slack = lineSlack(b.el) - 1;
    let at = lines.findIndex((l) => l.bottom + shift > contentBottom(p) + slack);
    if (at <= 0) return null;
    if (widowControl(b.node)) {
      if (lines.length - at < 2 && at - 1 > 0) at = lines.length - 2; // keep two lines for the next column
      if (at < 2 && !firstOnPage) return null; // a single line can't stay behind
    }
    if (at <= 0) return null;
    const cut = lines[at].top - b.top;
    return { cut, height: b.bottom - b.top, delta: contentTop(p + 1) - (lines[at].top + shift) };
  }

  // Columns of their own widths: a paragraph going on in a column of another width is laid out again at that width, as
  // Word does, so its lines there aren't the first column's. The text from its cut is copied out (`cs.rewrap`: the copy,
  // its width and height), and what follows moves by how much taller or shorter that makes it.
  const colWidthAt = (g, v) => g.widths[(((v - (g.P ?? 0)) % g.num) + g.num) % g.num];
  function rewrapCarry(b, cs, g, p) {
    const grow = colWidthAt(g, p + 1) - colWidthAt(g, p);
    if (Math.abs(grow) < 0.5) return;
    // The first character at or below the cut, in the paragraph's own layout (as paragraphLines measures it).
    const base = view.dom.getBoundingClientRect().top + b.undo;
    const pushes = [...b.el.querySelectorAll('.odx-push')].map((e) => e.getBoundingClientRect());
    const natural = (y) => y - base - pushes.reduce((sum, r) => sum + (r.top + r.height / 2 <= y ? r.height : 0), 0);
    const y = b.top + cs.cut - 1;
    const range = document.createRange();
    const box = (t, i) => { range.setStart(t, i); range.setEnd(t, i + 1); return [...range.getClientRects()].find((r) => r.height > 0); };
    let at = null;
    const walker = document.createTreeWalker(b.el, NodeFilter.SHOW_TEXT);
    for (let t = walker.nextNode(); t && !at; t = walker.nextNode()) {
      if (!t.length || t.parentElement.closest('.odx-push')) continue;
      const last = box(t, t.length - 1);
      if (last && natural(last.top) < y) continue; // (all of it above the cut)
      for (let i = 0; i < t.length; i++) { const r = box(t, i); if (r && natural(r.top) >= y) { at = [t, i]; break; } }
    }
    if (!at) return;
    const path = [];
    for (let n = at[0]; n !== b.el; n = n.parentNode) path.unshift([...n.parentNode.childNodes].indexOf(n));
    const copy = b.el.cloneNode(true);
    let end = copy;
    for (const i of path) end = end.childNodes[i];
    range.setStart(copy, 0);
    range.setEnd(end, at[1]);
    range.deleteContents();
    copy.removeAttribute('contenteditable');
    copy.style.removeProperty('margin-top');
    copy.style.margin = '0';
    copy.style.textIndent = '0';
    copy.classList.add('odx-contlines'); // (no list number or bullet: editor.css)
    copy.querySelectorAll('.ProseMirror-trailingBreak, .odx-push').forEach((e) => e.remove());
    const width = widthIn(b, g, p + 1);
    const h = heightAt(copy, width);
    cs.rewrap = { el: copy, width, h };
    cs.delta += h - (cs.height - cs.cut);
  }
  /** A block's width in flow page v's column: as laid out, less the narrowing it has now, plus that column's. */
  function widthIn(b, g, v) {
    const now = parseFloat(getComputedStyle(b.el).getPropertyValue('--odx-coln')) || 0;
    return b.el.getBoundingClientRect().width + now - (g.width - colWidthAt(g, v));
  }
  /** The height of `el` (a paragraph's copy, its margins 0) laid out `width` wide; with `lines`, its lines' bottoms too. */
  function heightAt(el, width, lines = null) {
    const outer = document.createElement('div'); // (not the sheet's own child .odx-content: the layout's rules are for that)
    outer.style.cssText = `position:absolute;visibility:hidden;left:0;top:0;width:${width}px`;
    const host = document.createElement('div');
    host.className = 'odx-content';
    host.append(el);
    outer.append(host);
    sheet.append(outer);
    const box = el.getBoundingClientRect();
    if (lines) {
      const range = document.createRange();
      const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      for (let t = walker.nextNode(); t; t = walker.nextNode()) {
        range.selectNodeContents(t);
        for (const r of range.getClientRects()) {
          if (!r.height) continue;
          const l = lines.find((x) => Math.abs(x.top - (r.top - box.top)) < 2);
          if (l) l.bottom = Math.max(l.bottom, r.bottom - box.top); else lines.push({ top: r.top - box.top, bottom: r.bottom - box.top });
        }
      }
      lines.sort((x, y) => x.top - y.top);
    }
    outer.remove();
    el.remove();
    return box.height;
  }
  /**
   * A paragraph laid out at another column's width (the one it had last time), at flow page p's column: null when its
   * width there is the one it has, else where it would end there (`bottom`) and whether it could be split at that
   * column's foot (`splits`, the test colSplit makes; it is then split there on the next pass, at that width).
   */
  function atColumnWidth(b, g, p, shift, firstOnPage) {
    if (!g?.unequal || b.node.type !== schema.nodes.paragraph) return null;
    const now = parseFloat(getComputedStyle(b.el).getPropertyValue('--odx-coln')) || 0;
    if (Math.abs(now - (g.width - colWidthAt(g, p))) < 0.5) return null;
    const copy = b.el.cloneNode(true);
    copy.removeAttribute('contenteditable');
    copy.style.removeProperty('margin-top');
    copy.style.margin = '0';
    const lines = [];
    const bottom = b.top + heightAt(copy, widthIn(b, g, p), lines);
    let at = lines.findIndex((l) => b.top + l.bottom + shift > contentBottom(p) + lineSlack(b.el) - 1);
    if (lines.length < 2 || keepLines(b.node)) at = -1;
    if (at > 0 && widowControl(b.node)) {
      if (lines.length - at < 2 && at - 1 > 0) at = lines.length - 2;
      if (at < 2 && !firstOnPage) at = -1;
    }
    return { bottom, splits: at > 0 };
  }

  // A table in columns goes on in the next column row by row: cut above the first row that doesn't fit, its rows from
  // there drawn again at the next column's top (as a paragraph's lines are, colSplit), under its header rows repeated.
  function colSplitTable(b, shift, p) {
    const table = b.el.tagName === 'TABLE' ? b.el : b.el.querySelector('table');
    const trs = table && table.tBodies[0] ? [...table.tBodies[0].rows] : [];
    if (!trs.length) return null;
    const top0 = b.el.getBoundingClientRect().top;
    const room = contentBottom(p) + 1 - (b.top + shift); // how far down the block the column ends
    const at = trs.findIndex((tr) => tr.getBoundingClientRect().bottom - top0 > room);
    if (at < 0) return null;
    // Inside the row, between its cells' lines, as Word breaks a row line by line (when the row can split and its cells
    // sit at the top): the lowest place above the column's foot that no line of any cell straddles.
    let cut = null;
    const node = b.node.child(at);
    const rowTop = trs[at].getBoundingClientRect().top - top0;
    if (!sheet.classList.contains('odx-large') && node && !cantSplit(node) && ![...trs[at].cells].some((td) => !/^(top|baseline)$/.test(getComputedStyle(td).verticalAlign))) {
      const base = view.dom.getBoundingClientRect().top + b.undo;
      const spans = [];
      for (const el of trs[at].querySelectorAll('p')) {
        const r = el.getBoundingClientRect();
        const pl = paragraphLines({ el, top: r.top - base, bottom: r.bottom - base, undo: b.undo });
        if (pl) for (const l of pl) spans.push([l.top - b.top, l.bottom - b.top]);
        else spans.push([r.top - top0, r.bottom - top0]);
      }
      const clean = (y) => !spans.some(([t, z]) => t < y - 0.5 && z > y + 0.5);
      const tops = [...new Set(spans.map(([t]) => t))].filter((y) => y > rowTop + 1 && y <= room + 0.5 && clean(y)).sort((x, y) => y - x);
      if (tops.length) cut = tops[0];
    }
    if (cut == null) {
      if (at <= 0) return null;
      cut = rowTop;
    }
    let h = 0;
    while (h < b.node.childCount && isHeader(b.node.child(h))) h++;
    const hb = h && h < trs.length ? trs[h].getBoundingClientRect().top - top0 : 0;
    const head = hb && cut >= hb - 0.5 ? hb : 0; // (the header rows drawn again above what goes on, once the cut is past them)
    return { cut, head, height: b.bottom - b.top, delta: contentTop(p + 1) + head - (b.top + shift + cut) };
  }

  function splitParagraph(b, shift, blockFirstOnPage, moved = false) {
    const lines = paragraphLines(b);
    if (!lines) return null;
    const pushes = [];
    let extra = 0;
    let pageStart = 0; // first line on the current page
    let tail = 0;
    const widow = widowControl(b.node);
    // Only the extra below the text may pass the margin (no more). And an exactly spaced line must end a little
    // short of it: Word moves one ending 0.05pt above the margin (doc 13's page 3, its own layout). (Not lines of other
    // spacing: sample4's 1.08-spaced Calibri ends 0.2px above it and stays.)
    const slack = lineSlack(b.el) - 1 - tie(b.el);
    // A line carried to the next page under "multiple" spacing keeps the browser's half of the extra above its text, as
    // the paragraph's first line does (the print shift takes it back): Word's text sits at the line's top either way.
    const pcs = getComputedStyle(b.el);
    const mult = pcs.getPropertyValue('--odx-lhlen').trim() === '1' ? 1 : parseFloat(pcs.getPropertyValue('--odx-mult')) || 1;
    const er = b.el.getBoundingClientRect();
    // (Less a list label's padding above it, the bullet line's own: sample S071's bullet carried to page 8, 1.9px low.)
    const lblUp = b.el.classList.contains('odx-li') ? parseFloat(getComputedStyle(b.el, '::before').paddingTop) || 0 : 0;
    const lead0 = lines[0].screenTop - (er.top + parseFloat(pcs.paddingTop) + parseFloat(pcs.borderTopWidth)) - lblUp;
    // (Under "at least" spacing too: Word puts the extra above the text there as well, corpus 2's 18pt-at-least Calibri
    // carried to page 11 3px high without it.)
    const atLeast = pcs.getPropertyValue('--odx-lhlen').trim() === '1' && parseFloat(pcs.getPropertyValue('--odx-runmult')) !== 0;
    const lead = (mult > 1 || atLeast) && Math.abs(lead0) <= 20 ? lead0 : 0; // (double spacing's too: corpus 2's double-spaced Cambria carried to page 13 sat 8px high without it)
    // And under any spacing, a text box a little taller than its line starts above it: a line the page ran out on goes
    // over with its box, not its text, at the page's top, as a paragraph's first line does (sample S071's single-spaced
    // Calibri, page 7 1.3px low). (Not one after a page break: sample S062's headings, 1.3px high with it.)
    const above = !lead && lead0 < 0 && lead0 >= -2 ? lead0 : 0;
    // A line's bottom is the next line's text top, which sits below that line's box by the text's offset in it: taken
    // off when fitting (doc 04: the second line of 4.4 ends 959.3px down a 960px page, as Word fits it).
    // (Up to the half of "multiple" spacing's extra, as for lead: sample S058's double-spaced 13pt Times, its
    // 10px offset left on, moved page 1's last line to page 2.)
    const inset = lead0 > 0 && (lead0 <= 6 || (lead && lead0 <= 20)) ? lead0 : 0;
    for (let n = 0; n < lines.length; n++) {
      const top = lines[n].top + shift + extra;
      const bottom = lines[n].bottom + shift + extra - (n < lines.length - 1 ? inset : 0);
      const pg = pageOf(top);
      let at = -1;
      if (lines[n].forced) {
        // A paragraph that starts a page with a page break: the break already happened (Word doesn't add a blank page).
        if (n === 1 && moved && pushes.length === 0 && leadingBreak(b.el)) continue;
        at = n;
      }
      else if (bottom > contentBottom(pg) + slack && n > pageStart) {
        if (keepLines(b.node) && pushes.length === 0) return blockFirstOnPage ? { pushes, total: 0 } : null;
        at = n;
        if (widow) {
          if (lines.length - at < 2 && at - 1 > pageStart) at = lines.length - 2; // keep two lines for the next page
          if (at - pageStart < 2 && pageStart === 0 && !blockFirstOnPage) return null; // a single line can't stay behind
        }
      } else if (bottom > contentBottom(pg) + slack && n === 0 && !blockFirstOnPage && !lines[1]?.forced) return null;
      // (A first line that only holds a page break stays at the bottom of its page, as in Word: the text after it starts the next.)
      if (at < 0) continue;
      const lineTop = lines[at].top + shift + extra;
      const delta = contentTop(pageOf(lines[at].forced ? lineTop : top) + 1) + lead + (lines[at].forced ? 0 : above) - lineTop;
      // A closing break's empty line that doesn't fit: no spacer before it (after a block, Edge gives that line no box).
      // It stays, hidden in the margin, and the next block starts the page that much lower, as if the line went first, as
      // in Word (sample S041: page 3 opens with it).
      if (lines[at].blank && !lines[at].forced) { tail = lines[at].bottom - lines[at].top; continue; }
      const pos = lineStart(b.el, lines[at].screenTop);
      if (pos == null || delta <= 0) continue;
      pushes.push({ pos, delta: Math.round(delta * 10) / 10 });
      extra += Math.round(delta * 10) / 10;
      pageStart = at;
      n = at;
    }
    // A page break with nothing after it on its line: the next block starts the new page.
    const breakAfter = hasPageBreak(b.node) && !lines.some((l) => l.forced);
    return { pushes, total: extra, breakAfter, tail };
  }

  let headKey = '';
  const nodeIds = new WeakMap(); // block node -> a number (a block whose content changed is a new node)
  let idSeq = 0;
  function drawHeads(list) {
    // (Unchanged repeats are left as they are: redrawing a long document's every pass took a third of its layout time.)
    const key = list.map((h) => { if (!nodeIds.has(h.b.node)) nodeIds.set(h.b.node, ++idSeq); return `${nodeIds.get(h.b.node)}:${h.b.i}:${h.b.top.toFixed(1)}:${h.top}:${h.rule ? 1 : 0}:${h.dy ?? ''}:${h.dx ?? ''}:${h.cut ?? ''}:${h.end ?? ''}:${h.rewrap ? `${h.rewrap.width.toFixed(1)}/${h.rewrap.h.toFixed(1)}` : ''}`; }).join(',');
    if (!list.length && !headKey) return;
    if (key === headKey && repeats.childElementCount) return;
    headKey = key;
    const origin = layer.getBoundingClientRect();
    const frag = document.createDocumentFragment();
    for (const { b, top, rule, cont, cut, end, dx, dy, rewrap } of list) {
      if (rewrap) {
        // Its lines laid out again at the next column's width (rewrapCarry): from the top of that copy.
        const r = b.el.getBoundingClientRect();
        const wrap = document.createElement('div');
        wrap.className = 'odx-repeat';
        wrap.style.cssText = `top:${(r.top - origin.top + dy + cut).toFixed(2)}px;left:${(r.left - origin.left + dx).toFixed(2)}px;width:${rewrap.width}px`;
        const host = document.createElement('div');
        host.className = 'odx-content';
        host.append(rewrap.el.cloneNode(true));
        wrap.append(host);
        frag.append(wrap);
        continue;
      }
      if (cont) {
        // A paragraph's lines past its column's bottom, again at the next column's top (columns: drawn, not editable there).
        const r = b.el.getBoundingClientRect();
        const wrap = document.createElement('div');
        wrap.className = 'odx-repeat';
        wrap.style.cssText = `top:${(r.top - origin.top + dy).toFixed(2)}px;left:${(r.left - origin.left + dx).toFixed(2)}px;width:${r.width}px`;
        const host = document.createElement('div');
        host.className = 'odx-content';
        const copy = b.el.cloneNode(true);
        copy.removeAttribute('contenteditable');
        copy.style.removeProperty('margin-top'); // (its page push, set as !important: margin 0 alone wouldn't take)
        copy.style.margin = '0';
        copy.style.clipPath = `inset(${cut.toFixed(2)}px -200px ${end != null ? (r.height - end).toFixed(2) : -20}px -200px)`; // (end: only down to there, a table's header rows)
        copy.querySelectorAll('.ProseMirror-trailingBreak, .odx-push').forEach((e) => e.remove());
        host.append(copy);
        wrap.append(host);
        frag.append(wrap);
        continue;
      }
      const t = b.table;
      if (rule) {
        // The bottom border of a row's part left behind at a page break, in the table's line colour.
        const r = t.getBoundingClientRect();
        const shadow = getComputedStyle(t.querySelector('td') || t).boxShadow;
        if (!shadow || shadow === 'none') continue; // a borderless table gets no break rule
        const color = (/rgb\([^)]*\)/.exec(shadow) || ['#000'])[0];
        const line = document.createElement('div');
        line.className = 'odx-repeat';
        line.style.cssText = `top:${top - 1}px;left:${r.left - origin.left}px;width:${r.width}px;height:1px;background:${color}`;
        frag.append(line);
        continue;
      }
      const copy = t.cloneNode(true);
      copy.removeAttribute('contenteditable');
      const body = copy.tBodies[0];
      while (body.rows.length > b.headerRows) body.lastElementChild.remove();
      copy.querySelectorAll('.column-resize-handle, .selectedCell').forEach((n) => n.classList.remove('selectedCell') || (n.classList.contains('column-resize-handle') && n.remove()));
      const r = t.getBoundingClientRect();
      const wrap = document.createElement('div');
      wrap.className = 'odx-repeat';
      wrap.style.cssText = `top:${top}px;left:${r.left - origin.left}px;width:${r.width}px`;
      wrap.append(copy);
      frag.append(wrap);
    }
    repeats.replaceChildren(frag);
  }

  let pages = 1;
  let frame = 0;
  let timer = 0;
  let lastMs = 0;
  let passes = 0;
  // Floating tables laid out in the flow, by node (an edit above one changes its index, not its node; a changed table is
  // a new node, measured again).
  const unfloated = new WeakSet();
  let applied = new Map(); // block index -> { extra px we added above it }
  const keep = (node) => node.type === schema.nodes.paragraph && (node.attrs.view.keepNext ?? session.styleProp(node.attrs.styleId, 'keepNext'));
  const paraBreakBefore = (node) => node.type === schema.nodes.paragraph && (node.attrs.view.pageBreakBefore ?? session.styleProp(node.attrs.styleId, 'pageBreakBefore'));
  // A table row starts a page when the first paragraph of its first cell says to, as in Word (the row moves, under its
  // header rows; a paragraph in another cell, or after the first, doesn't), unless the row above keeps with it (a
  // paragraph in it keeps with the next); a table whose first row does, starts one.
  const rowBreaksBefore = (row) => !!row.firstChild?.firstChild && paraBreakBefore(row.firstChild.firstChild);
  const rowKeeps = (row) => { let any = false; row.firstChild?.descendants((n) => { if (n.type === schema.nodes.paragraph) { if (keep(n)) any = true; return false; } return true; }); return any; };
  const rowBreakAt = (table, j) => j > 0 && rowBreaksBefore(table.child(j)) && !rowKeeps(table.child(j - 1));
  const rowBreakIn = (table) => { for (let j = 1; j < table.childCount; j++) if (rowBreakAt(table, j)) return true; return false; };
  const breakBefore = (node) => paraBreakBefore(node) || (node.type === schema.nodes.table && !!node.firstChild && rowBreaksBefore(node.firstChild));
  const spaceBefore = (node) => {
    if (node.type !== schema.nodes.paragraph) return 0;
    // HTML-style automatic spacing (w:beforeAutospacing) comes to nothing at the top of a page (but Word 2003's plain 5pt,
    // w:doNotUseHTMLParagraphAutoSpacing, stays).
    if (node.attrs.view.beforeAuto ?? session.styleProp(node.attrs.styleId, 'beforeAuto')) return session.compat?.has('doNotUseHTMLParagraphAutoSpacing') ? 5 * 4 / 3 : 0;
    const tw = node.attrs.view.before ?? session.styleProp(node.attrs.styleId, 'before') ?? session.styles.defaults.pPr.before ?? 0;
    return (tw / 20) * (4 / 3);
  };
  const spaceAfter = (node) => {
    if (node.type !== schema.nodes.paragraph) return 0;
    if (node.attrs.view.afterAuto ?? session.styleProp(node.attrs.styleId, 'afterAuto')) return 0;
    const tw = node.attrs.view.after ?? session.styleProp(node.attrs.styleId, 'after') ?? session.styles.defaults.pPr.after ?? 0;
    return (tw / 20) * (4 / 3);
  };
  const isHeader = (row) => /<w:tblHeader(?![^>]*w:val="(?:0|false|off)")/.test(row.attrs.trPr || '');
  const hasPageBreak = (node) => {
    let found = false;
    if (node.type === schema.nodes.paragraph) node.forEach((c) => { if (c.type === schema.nodes.hard_break && (c.attrs.type === 'page' || c.attrs.type === 'column')) found = true; });
    return found;
  };

  const heights = new Map();
  const storyHeight = (doc) => {
    if (!doc) return 0;
    if (!heights.has(doc)) {
      const el = story(doc, 'odx-header', 1, 1);
      el.style.cssText = `visibility:hidden;top:0;left:${page.left}px;right:${page.right}px`;
      sheet.append(el);
      clipExactRows(el); // (its exact rows as drawn: a cell's overflow doesn't grow them)
      fitRowspans(el);
      heights.set(doc, el.getBoundingClientRect().height);
      el.remove();
    }
    return heights.get(doc);
  };
  let sectCache = { doc: null, ends: new Map(), first: page };

  /**
   * Floats placed against a page or its margins (reader.js anchorCss: --odx-lrel), not their paragraph: Word draws them
   * on the page their paragraph is on, at that height, and that page's lines wrap round them, those above the paragraph
   * too. Each is drawn again at the start of the page's first paragraph, below a hairline-wide float as tall as the drop
   * to its height, and the original is hidden. Positions are read from the current layout (one pass behind: the next
   * pass settles them).
   */
  const liftFloats = (base) => {
    const out = [];
    view.state.doc.descendants((node, pos) => {
      if (node.type !== schema.nodes.image || !/--odx-lrel/.test(node.attrs.place || '')) return true;
      const img = view.nodeDOM(pos);
      const host = img && img.nodeType === 1 ? img.closest('.odx-content > *') : null;
      if (!host || host.parentElement !== view.dom) return false;
      const pg = pageOf(host.getBoundingClientRect().top - base);
      const place = node.attrs.place;
      const rel = /--odx-lrel:\s*(\w+)/.exec(place)[1];
      const off = parseFloat((/--odx-ly:\s*(-?[\d.]+)px/.exec(place) || [])[1]) || 0;
      const align = (/--odx-lyalign:\s*(\w+)/.exec(place) || [])[1];
      const h = Number(node.attrs.height) || 0;
      const m = marginsAt(pg);
      const top0 = pageTop(pg) + (rel === 'page' ? 0 : m.top);
      const bottom0 = pageTop(pg) + pageH(pg) - (rel === 'page' ? 0 : m.bottom);
      const y = align === 'top' ? top0 : align === 'bottom' ? bottom0 - h : align === 'center' ? (top0 + bottom0 - h) / 2 : top0 + off;
      const first = [...view.dom.children].find((e) => e.tagName === 'P' && e.getBoundingClientRect().top - base >= contentTop(pg) - 1);
      if (!first || pageOf(first.getBoundingClientRect().top - base) !== pg) return false;
      const drop = Math.round((y - (first.getBoundingClientRect().top - base)) * 10) / 10;
      const at = view.posAtDOM(first, 0);
      const cs = getComputedStyle(img);
      const side = cs.float === 'right' ? 'right' : 'left';
      const src = img.getAttribute('src') || '';
      const style = `float:${side};clear:${side};width:${node.attrs.width}px;height:${h}px;margin:${Math.min(0, drop)}px ${cs.marginRight} ${cs.marginBottom} ${cs.marginLeft}`;
      out.push(Decoration.widget(at, () => {
        const w = document.createElement('span');
        w.className = 'odx-lift';
        w.contentEditable = 'false';
        const sp = document.createElement('span');
        sp.style.cssText = `float:${side};width:0.01px;height:${Math.max(0, drop)}px`;
        const copy = document.createElement('img');
        copy.src = src;
        copy.alt = node.attrs.alt || '';
        copy.style.cssText = style;
        w.append(sp, copy);
        return w;
      }, { side: -1, key: `lift:${at}:${drop}:${style}`, ignoreSelection: true }));
      out.push(Decoration.node(pos, pos + node.nodeSize, { class: 'odx-lifted' }, { key: 'lifted' }));
      return false;
    });
    return out;
  };

  /**
   * Drawings out of the flow placed against the page or its margins vertically (reader.js anchorCss: --odx-arel): Word
   * draws them on the page their paragraph is on, at that height (a cover page's title box, a full-page picture behind
   * the text), not at the paragraph. Given that top, from the layout before (one pass behind, as liftFloats).
   */
  const placeOnPage = (base) => {
    const out = [];
    view.state.doc.descendants((node, pos) => {
      if (node.type !== schema.nodes.image || !/--odx-arel/.test(node.attrs.place || '')) return true;
      const img = view.nodeDOM(pos);
      const host = img && img.nodeType === 1 ? img.closest('.odx-content > *') : null;
      const cb = img && img.offsetParent;
      if (!host || host.parentElement !== view.dom || !cb) return false;
      const pg = pageOf(host.getBoundingClientRect().top - base);
      const place = node.attrs.place;
      const rel = /--odx-arel:\s*(\w+)/.exec(place)[1];
      const off = parseFloat((/--odx-ay:\s*(-?[\d.]+)px/.exec(place) || [])[1]) || 0;
      const align = (/--odx-ayalign:\s*(\w+)/.exec(place) || [])[1];
      const h = Number(node.attrs.height) || 0;
      const m = marginsAt(pg);
      const p0 = pageTop(pg), p1 = p0 + pageH(pg);
      const [top0, bottom0] = rel === 'page' ? [p0, p1] : rel === 'topMargin' ? [p0, p0 + m.top] : rel === 'bottomMargin' ? [p1 - m.bottom, p1] : [p0 + m.top, p1 - m.bottom];
      const y = align === 'top' ? top0 : align === 'bottom' ? bottom0 - h : align === 'center' ? (top0 + bottom0 - h) / 2 : top0 + off;
      const top = Math.round((y - (cb.getBoundingClientRect().top - base)) * 10) / 10;
      out.push(Decoration.node(pos, pos + node.nodeSize, { style: `top:${top}px;margin-top:0` }, { key: `onpage:${top}` }));
      return false;
    });
    return out;
  };

  // The top-level blocks' elements in order, from the view's own tree in one walk (view.nodeDOM finds each by counting
  // from the start: quadratic, 70ms a pass over 5,000 blocks); null when that tree doesn't line up with the document.
  function topElements(doc) {
    const kids = view.docView?.children;
    if (!kids) return null;
    const out = [];
    for (const d of kids) if (d.node) out.push(d.dom);
    return out.length === doc.childCount ? out : null;
  }

  // The blocks where they are now, less the pushes the last pass applied: the natural (unpaginated) flow, without a
  // second layout. Tables with their rows, so a table can continue on the next page; not in columns (they move whole).
  function measureBlocks(dom, doc, base, inRegion) {
    const blocks = [];
    let undo = 0;
    const els = topElements(doc);
    doc.forEach((node, pos, i) => {
      const el = els ? els[i] : view.nodeDOM(pos);
      // Hidden blocks (deleted text with markup off) have no box: measured, they'd sit at the top of the document.
      if (!el || el.nodeType !== 1 || !el.getClientRects().length) return;
      const r = el.getBoundingClientRect();
      const a = applied.get(i);
      // (A floating block's margin moves only itself: what follows it in the flow is where it was.)
      const float = a && a.extra && outOfFlow(el);
      if (a && !float) undo += a.extra;
      const block = { i, node, el, top: r.top - base - undo - (float ? a.extra : 0), undo: undo + (float ? a.extra : 0) };
      if (node.type === schema.nodes.table) {
        // Rows too, so a table can continue on the next page.
        const table = el.tagName === 'TABLE' ? el : el.querySelector('table');
        const trs = table && table.tBodies[0] ? [...table.tBodies[0].rows] : [];
        if (table && trs.length === node.childCount && !cols && !inRegion(i)) { // (in columns, tables move whole)
          block.table = table;
          block.tableTop = table.getBoundingClientRect().top - base - undo;
          block.rows = trs.map((tr, j) => {
            // A pushed row's extra padding is inside it: it moves its bottom, not its top.
            const rr = tr.getBoundingClientRect();
            const rowTop = rr.top - base - undo;
            const d = a && a.rows && a.rows.get(j);
            if (d) undo += d;
            const contentUndo = undo;
            let bottom = rr.bottom;
            if (a && a.inRow && a.inRow.has(j)) {
              // Split inside its cells: the row's natural bottom is its tallest cell without our spacers.
              bottom = -Infinity;
              for (const td of tr.cells) {
                const last = td.lastElementChild;
                if (!last) continue;
                const spacers = undoIn(td, Infinity); // (in a table in the cell, a row's tallest cell's)
                bottom = Math.max(bottom, last.getBoundingClientRect().bottom - spacers + (parseFloat(getComputedStyle(td).paddingBottom) || 0));
              }
              if (!(bottom > rr.top)) bottom = rr.bottom;
            }
            // The line under the row is the next row's (styles.js), but a row ending a page has to have room for it.
            const line = Math.max(0, ...[...tr.cells].map((td) => parseFloat(td.style.getPropertyValue('--odx-bdb')) || 0));
            const row = { top: rowTop, bottom: bottom - base - undo, undo: contentUndo, line };
            undo += rr.bottom - bottom;
            return row;
          });
          if (block.rows.some((x) => x.bottom <= x.top)) block.rows = null; // not laid out (content-visibility)
          let h = 0;
          while (h < node.childCount && isHeader(node.child(h))) h++;
          block.headerRows = h < node.childCount ? h : 0;
        }
      }
      if (a && a.inline) undo += a.inline;
      block.bottom = r.bottom - base - undo - (float ? a.extra : 0);
      blocks.push(block);
    });
    return blocks;
  }

  // Each note reference's note, by the block it's in (endnotes by the last block), with its height: blockNotes.
  function assignNotes(doc) {
    blockNotes = new Map();
    const refs = noteRefs(doc);
    if (refs.length) {
      const ends = [];
      doc.forEach((node, pos, i) => ends.push([i, pos + node.nodeSize]));
      let bi = 0;
      for (const r of refs) {
        const note = session.notes?.[r.kind]?.get(r.id);
        if (!note) continue;
        while (bi < ends.length - 1 && ends[bi][1] <= r.pos) bi++;
        // Endnotes go after the last text, at the end of the document.
        const i = r.kind === 'endnote' ? ends[ends.length - 1][0] : ends[bi][0];
        if (!blockNotes.has(i)) blockNotes.set(i, []);
        blockNotes.get(i).push({ ...r, h: noteHeight(note, r.kind, r.n), note });
      }
    }

  }

  // The end of a columns section, at block k (top: where it would go): its last flow page and sheet. Balanced, what
  // follows goes under its lowest column there (on the flow page of its last column, so the sheet's other columns stay
  // where they are): `moved` is how far block k moves down (also added to `starts`); `again`, that the column height
  // its last sheet is cut to changed, so layout runs again.
  function endColumns(re, k, b, top, blocks, rec, sectEnds, starts) {
    let moved = null;
    let again = false;
    const bottoms = new Map();
    const ends = new Map(); // flow page -> how low its column ends with the space after its last paragraph
    let vEnd = re.P;
    for (let j = re.k0; j < k; j++) {
      const x = rec.get(j);
      if (!x) continue;
      const v = pageOf(x.top);
      vEnd = Math.max(vEnd, v);
      if (!blocks[j].node.childCount && sectEnds.get(blocks[j].i)) continue; // (an empty section-break paragraph: Word's columns end above it)
      bottoms.set(v, Math.max(bottoms.get(v) ?? -Infinity, x.bottom));
      ends.set(v, Math.max(ends.get(v) ?? -Infinity, x.bottom + (x.carry ? 0 : spaceAfter(blocks[j].node)))); // (a paragraph cut there has none)
      vEnd = Math.max(vEnd, v);
      if (x.carry) {
        const y = contentTop(v + 1) + x.carry;
        bottoms.set(v + 1, Math.max(bottoms.get(v + 1) ?? -Infinity, y));
        ends.set(v + 1, Math.max(ends.get(v + 1) ?? -Infinity, y + spaceAfter(blocks[j].node)));
        vEnd = Math.max(vEnd, v + 1);
      }
    }
    const sheet = Math.floor((vEnd - re.P) / re.num);
    re.vEnd = vEnd;
    re.offAfter = vEnd - (re.S0 + sheet);
    for (const v of [...colTops.keys()]) if (v > vEnd) colTops.delete(v);
    if (re.balance) {
      const first = re.P + sheet * re.num;
      let total = 0, low = 0;
      const base = (v) => (v === re.P ? re.y0 : contentTop(v)); // (the first column starts where the section does)
      for (let v = first; v <= vEnd; v++) {
        const h = Math.max(0, (bottoms.get(v) ?? base(v)) - base(v));
        total += h;
        low = Math.max(low, (ends.get(v) ?? base(v)) - base(v));
      }
      for (let v = first; v <= vEnd; v++) cuts.delete(v);
      // (The lowest a column ends with the space after its last paragraph, but not one cut at its foot; then the space
      // before less the section-break paragraph's space after, as at a new section's page: corpus 3's c3-06, c3-08.)
      const d = base(vEnd) + low + Math.max(0, spaceBefore(b.node) - spaceAfter(blocks[k - 1].node)) - top;
      starts.set(b.i, (starts.get(b.i) || 0) + d);
      moved = d;
      // The height to cut its columns to: an even share, and half a line more so the first column takes the odd one.
      const line = parseFloat(lineHeightOf(getComputedStyle(blocks[re.k0].el))) || 18;
      const want = { sheet, h: total / re.num + line / 2 };
      const had = balanced.get(re.lo);
      if (!had || had.sheet !== sheet || Math.abs(had.h - want.h) > 2) { balanced.set(re.lo, want); again = true; }
    }
    return { moved, again };
  }

  // A columns section starting at block k: where its first column starts (`P`, `y0`), the tops of the columns beside it,
  // and, balanced, the cut its last sheet's columns end at.
  function startColumns(rs, regs, sect, firstOnPage, p, top) {
    const newPage = sect && sect.page && !firstOnPage;
    const P = newPage ? p + 1 : p;
    const y0 = newPage ? contentTop(P) : Math.max(top, contentTop(p));
    const n = regs.indexOf(rs);
    rs.P = P;
    rs.y0 = y0;
    rs.off0 = n ? regs[n - 1].offAfter || 0 : 0;
    rs.S0 = P - rs.off0;
    for (let c = 1; c < rs.num; c++) colTops.set(P + c, pageTop(P + c) + (y0 - pageTop(P)));
    const bal = rs.balance && balanced.get(rs.lo);
    if (bal) for (let c = 0; c < rs.num - 1; c++) { const v = P + bal.sheet * rs.num + c; cuts.set(v, (v === P ? y0 : contentTop(v)) + bal.h); }
  }

  // Columns: each block moved up beside the column before it (CSS), the lines a block carries into the next column
  // (drawn again there), and the page count, sections and page setup counted in sheets.
  function placeColumns(blocks, regs, starts, colSplits, paraSplits, splits) {
    const colCss = [];
    const contHeads = [];
    const live = regs.filter((r) => r.P != null);
    if (cols || live.length) {
      // Each block's column: its page in the flow. Move it up beside the column before it. A columns section's columns
      // count from the page it starts on; what follows it moves up by the flow pages its columns took beyond their sheets.
      const geo = (k) => (cols ? { P: 0, S0: 0, off0: 0, ...cols } : live.find((r) => k >= r.k0 && k < r.k1));
      const offAt = (k) => { let o = 0; for (const r of live) if (k >= r.k1 && r.offAfter != null) o = r.offAfter; return o; };
      const inner = page.width - page.left - page.right;
      const place = (k, v) => {
        const g = geo(k);
        if (!g) return { s: v - offAt(k), c: 0, w: 0, x: 0 };
        // (narrow: how much narrower than the widest its column is, columns of their own widths.)
        if (v < g.P) return { s: v - g.off0, c: 0, w: 0, x: 0, narrow: g.unequal ? g.width - g.widths[0] : 0 };
        // (A right-to-left section (w:bidi) fills its columns from the right, as in Word.)
        const c = (v - g.P) % g.num;
        return { s: g.S0 + Math.floor((v - g.P) / g.num), c, w: g.width + g.space, x: g.rtl ? inner - g.xs[c] - g.widths[c] : g.xs[c], narrow: g.unequal ? g.width - g.widths[c] : 0 };
      };
      const narrowNow = new Map();
      let sh = 0;
      let lastS = blocks.length ? place(blocks.length - 1, pages - 1).s : 0;
      for (let k = 0; k < blocks.length; k++) {
        const b = blocks[k];
        sh += starts.get(b.i) || 0;
        const v = pageOf(b.top + sh);
        const { s, x, narrow } = place(k, v);
        const up = v - s; // pages of flow above this column's sheet
        // Its column's width (measured with it too, so a block that lands in another column is measured again there).
        if (narrow) { narrowNow.set(b.i, narrow); colCss.push(`.${scope} > .odx-content > :nth-child(${b.i + 1}){--odx-coln:${narrow.toFixed(2)}px}`); }
        if (up || x) colCss.push(`.${scope} > .odx-content:not(.odx-measure) > :nth-child(${b.i + 1}){transform:translate(${x.toFixed(1)}px,${-(up * pitch)}px)}`);
        const cs = colSplits.get(b.i);
        if (cs) {
          colCss.push(`.${scope} > .odx-content:not(.odx-measure) > :nth-child(${b.i + 1}){clip-path:inset(-20px -200px ${(cs.height - cs.cut).toFixed(2)}px -200px)}`);
          const v2 = v + 1, q = place(k, v2), up2 = v2 - q.s;
          lastS = Math.max(lastS, q.s); // (the lines it carries over count for the page count)
          const head = cs.head || 0;
          contHeads.push({ b, cont: true, cut: cs.cut, rewrap: cs.rewrap, dx: q.x - x, dy: (contentTop(v2) + head - up2 * pitch) - (b.top + sh + cs.cut - up * pitch) });
          if (head) contHeads.push({ b, cont: true, cut: 0, end: head, dx: q.x - x, dy: (contentTop(v2) - up2 * pitch) - (b.top + sh - up * pitch) });
        }
        if (!cols) sh += (paraSplits.get(b.i)?.total || 0) + (splits.get(b.i)?.total || 0); // (what a block split over pages pushes below it)
      }
      pages = lastS + 1;
      // Sections start on sheets, not flow pages.
      const toSheet = (v) => { let o = 0; for (const r of live) { if (v < r.P) break; if (r.vEnd == null || v <= r.vEnd) return r.S0 + Math.floor((v - r.P) / r.num); o = r.offAfter; } return v - o; };
      if (!cols) {
        sections = sections.map(([f, n]) => [toSheet(f), n]);
        // And their page setup (headers, footers, size, numbering) with them: a section after a columns section had its
        // header and numbers on the sheet its flow page would have been.
        const bySheet = [];
        for (const [f, v, fresh] of margins) {
          const at = toSheet(f);
          if (bySheet.length && bySheet[bySheet.length - 1][0] === at) bySheet.pop(); // (the later one wins a sheet)
          bySheet.push([at, v, fresh]);
        }
        margins = bySheet;
      }
      colCss.push(`.${scope} > .odx-content{height:${pageTop(pages) - gap}px;overflow:visible}`);
      // Blocks that changed column width: lay out again at the new widths.
      const moved = narrowNow.size !== narrowed.size || [...narrowNow].some(([i, n]) => Math.abs((narrowed.get(i) ?? 0) - n) > 0.01);
      narrowed = narrowNow;
      if (moved) colsAgain = true;
    }
    return { colCss, contHeads };
  }

  /**
   * A section's vertical alignment (w:vAlign): on each of its pages, its blocks centred between the margins, at the
   * bottom, or spread over the page (both: the room left goes evenly between them; one block alone stays at the top),
   * as in Word. Added to `starts`; the next page's first block takes it back off. Pages with a block split across them
   * keep their place. (The empty paragraph that only ends a section takes no room in Word, nor here.)
   */
  function verticalAlign(doc, blocks, starts, pulls, splits, paraSplits) {
    const how = sectionXmls(doc).map((x) => /<w:vAlign\b[^>]*w:val="(center|bottom|both)"/.exec(x)?.[1] || null);
    if (!how.some(Boolean)) return;
    const extra = new Map(starts);
    for (const list of pulls.values()) for (const [i, d] of list) extra.set(i, (extra.get(i) || 0) + d);
    // Where each block ends up, and its section.
    let acc = 0, sect = 0;
    const at = blocks.map((b) => {
      acc += extra.get(b.i) || 0;
      const split = (paraSplits.get(b.i)?.total || 0) + (splits.get(b.i)?.total || 0);
      const r = { b, top: b.top + acc, bottom: b.bottom + acc + split, split: split !== 0, sect, after: parseFloat(getComputedStyle(b.el).marginBottom) || 0 };
      acc += split;
      r.p = pageOf(r.top);
      if (sectCache.ends.get(b.i)) { r.sectEnd = true; sect++; }
      return r;
    });
    const groups = new Map(); // "section:page" -> blocks
    for (const r of at) {
      if (!how[r.sect] || (r.sectEnd && !r.b.node.childCount)) continue;
      const key = `${r.sect}:${r.p}`;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(r);
    }
    for (const g of groups.values()) {
      const p = g[0].p, v = how[g[0].sect], end = g[g.length - 1];
      if (g.some((r) => r.split || r.bottom > contentBottom(p) + 1)) continue;
      // (Nor a page that another section goes on to share, after a continuous break: the room below is that one's.)
      if (at.some((r) => r.p === p && r.sect !== g[0].sect && r.top > end.top)) continue;
      const slack = contentBottom(p) - (end.bottom + end.after);
      if (slack <= 0.5) continue; // (a full page: nothing to move)
      const moves = v === 'both' ? g.slice(1).map((r) => [r, slack / (g.length - 1)]) : [[g[0], v === 'center' ? (slack - (g[0].top - contentTop(p))) / 2 : slack]];
      let total = 0;
      for (const [r, d] of moves) { starts.set(r.b.i, (starts.get(r.b.i) || 0) + d); total += d; }
      const next = at.find((r) => r.p > p); // (the next page's first block; an empty section-ending paragraph goes down with the rest)
      if (next && total) starts.set(next.b.i, (starts.get(next.b.i) || 0) - total);
    }
  }

  // The walk's result applied: margin-tops (node decorations), split tables' padded rows and clip paths (CSS), spacer
  // widgets inside split paragraphs and cells, baseline shifts and lifted floats. Returns the CSS and the header rows to
  // draw again.
  const outOfFlow = (el) => { const cs = getComputedStyle(el); return cs.float !== 'none' || cs.position === 'absolute' || cs.position === 'fixed'; };
  function applyPushes(doc, dom, blocks, base, colRegs, starts, pulls, splits, paraSplits) {
    // Pushed blocks: margin-top = natural gap + delta. The natural gap is
    // measured between this block's top and the previous block's bottom.
    // Set on the blocks themselves (node decorations), not by rules on their place among their siblings
    // (:nth-child): with those, a paragraph added (Enter) made the browser restyle every block after it, 130-190 ms a
    // keystroke in a 443-page document.
    const posOf = new Map();
    doc.forEach((node, pos, i) => posOf.set(i, [pos, pos + node.nodeSize]));
    const blockDecos = [];
    const onBlock = (i, attrs, key) => { const r = posOf.get(i); if (r) blockDecos.push(Decoration.node(r[0], r[1], attrs, { key })); };
    doc.forEach((node, pos, i) => { if (unfloated.has(node)) onBlock(i, { style: 'float:none !important' }, 'nf'); });
    const css = [];
    applied = new Map();
    const index = new Map(blocks.map((b, k) => [b.i, k]));
    const extra = new Map(starts);
    for (const list of pulls.values()) for (const [i, d] of list) extra.set(i, (extra.get(i) || 0) + d);
    for (const [i, delta] of extra) {
      const k = index.get(i);
      // (From the block before it in the flow: a margin doesn't count from a floating table's bottom, the text beside it
      // starts at its top. A paragraph after one, starting a page, went to 342px above the page: everything after it lost.)
      let q = k - 1;
      while (q >= 0 && outOfFlow(blocks[q].el)) q--;
      const gap = q >= 0 ? blocks[k].top - blocks[q].bottom : blocks[k].top - (parseFloat(getComputedStyle(dom).paddingTop) || 0); // the page's top padding is the first section's plain top margin, not one moved down for a tall header
      const margin = Math.round((gap + delta) * 10) / 10;
      applied.set(i, { extra: margin - gap });
      // Under the bottom margin of the block before, a smaller margin collapses into it: give the difference as a negative
      // one. (Only after a balanced columns section does a block move up past it.)
      const mb = colRegs.length && k > 0 ? parseFloat(getComputedStyle(blocks[k - 1].el).marginBottom) || 0 : 0;
      const set = margin < mb ? Math.round((margin - mb) * 10) / 10 : margin;
      onBlock(i, { style: `margin-top:${set}px !important` }, `mt:${set}`);
    }
    // Split tables: rows that start a page get extra top padding (room for the
    // page gap and the repeated header row), and the table is clipped across
    // that band so its borders and shading don't run through the gap.
    const heads = [];
    for (const [i, split] of splits) {
      const b = blocks[index.get(i)];
      const rows = new Map();
      // (Found by a class on the table itself, for the same reason.)
      onBlock(i, { class: `odx-sp${i}` }, `sp:${i}`);
      const sel = `.${scope} > .odx-content > .odx-sp${i}`;
      const bands = [];
      let extra = 0;
      for (const [j, d] of split.rows) {
        // Split rows above grow too, so use the offset measured while splitting.
        const rel = b.rows[j].top - b.tableTop + (split.offs?.get(j) ?? extra);
        const dd = Math.round(d * 10) / 10;
        rows.set(j, dd);
        bands.push([rel + 1, rel + dd]);
        extra += dd;
        css.push(`${sel} tbody > tr:nth-child(${j + 1}) > td{padding-top:calc(var(--odx-pt, 0px) + ${dd}px) !important;--odx-trh-add:${dd}px}`);
      }
      for (const band of split.bands) bands.push(band);
      bands.sort((x, y) => x[0] - y[0]);
      const cut = bands.map(([a, z]) => `V${a.toFixed(1)} H-20 V${z.toFixed(1)} H99999`).join(' ');
      css.push(`${sel} table{clip-path:path('M-20 -20 H99999 ${cut} V9999999 H-20 Z')}`);
      const a = applied.get(i) || { extra: 0 };
      a.rows = rows;
      a.inRow = split.inRow;
      applied.set(i, a);
      if (b.headerRows) for (const t of split.heads) heads.push({ b, top: t });
      for (const t of split.closes) heads.push({ b, top: t, rule: true });
    }
    const decos = [...blockDecos];
    for (const split of splits.values()) {
      for (const { pos, delta } of split.cellPushes) decos.push(Decoration.widget(pos, spacer(delta), { side: -1, key: `push:${delta}`, ignoreSelection: true }));
    }
    for (const [i, split] of paraSplits) {
      const a = applied.get(i) || { extra: 0 };
      a.inline = split.total;
      applied.set(i, a);
      for (const { pos, delta } of split.pushes) decos.push(Decoration.widget(pos, spacer(delta), { side: -1, key: `push:${delta}`, ignoreSelection: true }));
    }
    // Table-cell paragraphs' print-only baseline shifts ride along as node decorations.
    for (const c of cellShifts(view, sheet)) decos.push(Decoration.node(c.pos, c.pos + c.size, { style: `--odx-bl:${c.d}px` }, { key: `bl:${c.d}` }));
    decos.push(...liftFloats(base));
    decos.push(...placeOnPage(base));
    // (Keys without positions: a decoration a keystroke only moves stays as it is, rather than its block drawn again;
    // the positions are in the signature.)
    const pushSig = decos.map((d) => `${d.from}:${d.to}:${d.spec.key}`).sort().join(',');
    // (The set the view has, moved through the edits since: the same as this one when the layout didn't change, as after
    // most keystrokes, and then it isn't set again: building it afresh took 70ms a keystroke in a 750-page document.)
    const now = pushKey.getState(view.state)?.find().map((d) => `${d.from}:${d.to}:${d.spec.key}`).sort().join(',');
    if (pushSig !== now) {
      view.dispatch(view.state.tr.setMeta(pushKey, decos).setMeta('addToHistory', false));
    }
    return { css, heads };
  }

  // The layout's own writes into the document's elements (tab widths, row heights, column widths) with the editor's
  // change observer paused: it would take them for edits and read back everything between the first and the last of
  // them, a third of a 20MB document's tables at each keystroke (2.4-4s).
  const quiet = (fn) => {
    const o = view.domObserver;
    if (!o) return fn();
    o.stop();
    try { return fn(); } finally { o.start(); }
  };
  let passed = false; // the last layout() got past its early returns
  // The layout measures the pages at 100% whatever the zoom (zoom.js): every length it reads and sets is the page's own.
  const unzoomed = (fn) => {
    const z = sheet.style.zoom;
    if (!z) return fn();
    sheet.style.zoom = '';
    try { return fn(); } finally { sheet.style.zoom = z; }
  };
  function layout() { return unzoomed(layoutPass); }
  function layoutPass() {
    frame = 0;
    passed = false;
    if (!colsRepass) colsPasses = 0; // (a pass asked for from outside: the column widths may settle afresh)
    colsRepass = false;
    const t0 = performance.now();
    quiet(() => layoutTabs(view, session));
    if (quiet(() => fitAutofitColumns(view))) return schedule(); // column widths changed: measure again once they're drawn
    if (applyLineFix(view)) return schedule(); // line heights changed: measure again once they're drawn
    if (applyWidths(view)) return schedule(); // hinted fonts' widths corrected: measure again once they're drawn
    if (applyFitText(view)) return schedule(); // fitted text spaced out to its width: measure again
    if (applySqueeze(view, session)) return; // lines changed: measure again once they're drawn
    if (quiet(() => fitRowspans(view.dom))) return schedule(); // merged cells' extra height moved to their last row: measure again
    passed = true;
    const dom = view.dom;
    const doc = view.state.doc;
    // Word lets a floating table longer than a page run on in the flow: nothing wraps beside it.
    // Kept as a rule too: redrawing the table's decorations would reset an inline style.
    [...dom.children].forEach((t, i) => {
      if (getComputedStyle(t).float !== 'none' && t.getBoundingClientRect().height > pageHeight(0) && i < doc.childCount) { quiet(() => { t.style.float = 'none'; }); unfloated.add(doc.child(i)); }
    });
    const colRegs = colRegions(doc);
    const inRegion = (i) => colRegs.some((r) => i >= r.lo && i <= r.hi);
    if (cols || colRegs.length) dom.classList.add('odx-measure'); // measure the flow without the columns' moves
    sheet.classList.add('odx-bl-off'); // and without the print-only baseline shifts
    const base = dom.getBoundingClientRect().top;
    const blocks = measureBlocks(dom, doc, base, inRegion);
    // Walk the blocks and decide where pages start. `shift` is how far the
    // blocks already pushed have moved everything below them.
    const starts = new Map(); // block index -> extra px pushed (applied as margin-top)
    const carried = new Set(); // blocks after a frame moved to the next page: their push is the frame's
    const tails = new Map(); // block index -> the height of the line before it that Word starts its page with
    const splits = new Map(); // table block index -> { rows: Map(row -> extra px), heads: [page top] }
    const paraSplits = new Map(); // paragraph block index -> { pushes: [{ pos, delta }], total }
    const colSplits = new Map(); // (columns) paragraph block index -> { cut, height }: its lines from cut go on in the next column
    // Tables set down at a place on their page (styles.js: --odx-fy) and the blocks moved below them: block index ->
    // [[block index, px]...] added to those blocks' margin-top; taken back off when the walk comes past the block again.
    const pulls = new Map();
    let band = null; // { p, y0, y1 }: where such a table sits, which text after it goes above or below
    let shift = 0;
    let breakPending = false;
    if (sectCache.doc !== doc) sectCache = { doc, ...sectionBreaks(doc, session, page, storyHeight) };
    const sectEnds = sectCache.ends;
    margins = [[0, sectCache.first, true]];
    sections = [[0, 0]];
    notesOn = new Map();
    notesH = new Map();
    noteFlow = new Map();
    cuts = new Map();
    colTops = new Map();
    for (const r of colRegs) {
      r.k0 = blocks.findIndex((x) => x.i >= r.lo);
      r.k1 = blocks.findIndex((x) => x.i > r.hi);
      if (r.k1 < 0) r.k1 = blocks.length;
    }
    const regs = colRegs.filter((r) => r.k0 >= 0 && r.k0 < r.k1);
    if (balancedDoc !== doc) { balancedDoc = doc; balancePasses = 0; }
    let balanceAgain = false;
    const rec = new Map(); // (columns sections) block k -> { top, bottom, carry } where it was last placed
    const frameDy = new Map(); // block index -> px a frame set down on its page is moved below where it comes
    placed.clear();
    assignNotes(doc);
    for (let k = 0; k < blocks.length; k++) {
      const b = blocks[k];
      if (pulls.has(b.i)) { for (const [, d] of pulls.get(b.i)) shift -= d; pulls.delete(b.i); }
      if (k === 0 || (band && band.k >= k)) { if (band) cuts.delete(band.p); band = null; }
      let top = b.top + shift;
      let p = pageOf(top);
      // A paragraph in a frame (frames.js) floats beside the text, out of the flow. One anchored to the page or the
      // margins is set down at its place on the page where it comes: how far below that its float is moved (--odx-fdy).
      const frame = b.el.dataset?.frame;
      if (frame && b.node.type === schema.nodes.paragraph && getComputedStyle(b.el).float !== 'none') {
        if (frame === 'page' || frame === 'margin') {
          const now = parseFloat(getComputedStyle(b.el).getPropertyValue('--odx-fdy')) || 0;
          const natural = top - now;
          const pg = pageOf(natural);
          const base = frame === 'page' ? pageTop(pg) : contentTop(pg);
          const room = frame === 'page' ? pageH(pg) : contentBottom(pg) - contentTop(pg);
          const h = b.bottom - b.top;
          const al = b.el.dataset.fyalign;
          const y = al === 'top' ? base : al === 'center' ? base + (room - h) / 2 : al === 'bottom' ? base + room - h : base + Number(b.el.dataset.fy || 0);
          frameDy.set(b.i, y - natural);
          if (Math.abs(y - natural - now) > 0.5) colsAgain = true; // (the text beside it moves: lay out again)
        } else if (frame !== 'cont') { // (the frame's first paragraph; the rest are 'cont')
          // A frame in the text that doesn't fit in what's left of its page goes to the next page whole, and the text
          // after it with it, as in Word (four headings in a frame below a title: the page holds only the title).
          let k2 = k;
          let bottom = b.bottom + shift;
          while (k2 + 1 < blocks.length && blocks[k2 + 1].el.dataset?.frame === 'cont') bottom = Math.max(bottom, blocks[++k2].bottom + shift);
          if (bottom > contentBottom(p) + 1 && top > contentTop(p) + 1) { // (even one taller than a page: it starts the next)
            const d = contentTop(p + 1) - top;
            starts.set(b.i, (starts.get(b.i) || 0) + d);
            const nb = blocks[k2 + 1];
            if (nb) { starts.set(nb.i, (starts.get(nb.i) || 0) + d); carried.add(nb.i); } // (the text after it: a float's margin doesn't move it)
            shift += d;
          }
        }
        continue;
      }
      // A table set down on its page: there, out of the flow. What follows it starts where it would have (the page above
      // it fills first), and a block reaching into its band goes below it, as in Word.
      const fy = floatAt(b);
      if (fy) {
        const y0 = (fy.rel === 'page' ? pageTop(p) : contentTop(p)) + fy.y;
        const nb = blocks[k + 1];
        const d1 = y0 - top;
        const d2 = nb ? top + (nb.top - b.bottom) - (nb.top + shift + d1) : 0;
        pulls.set(b.i, nb ? [[b.i, d1], [nb.i, d2]] : [[b.i, d1]]);
        shift += d1 + d2;
        band = { k, p, y0: y0 - fy.above, y1: y0 + (b.bottom - b.top) + fy.below };
        // With less than two lines' room under it, the page's text ends above it: a table reaching into it splits there
        // and goes on on the next page (the contingency plan's page 4), rather than all of it moving under the band.
        if (band.y1 > contentBottom(p) - 40) { cuts.set(p, band.y0); band.cut = true; }
        continue;
      }
      if (band && !band.cut && band.p === p && top < band.y1 && b.bottom + shift > band.y0) {
        const d = band.y1 - top;
        pulls.set(b.i, [[b.i, d]]);
        shift += d;
        top += d;
        p = pageOf(top);
      }
      // The end of a columns section (endColumns).
      const re = regs.find((r) => r.k1 === k && r.P != null);
      if (re) {
        const e = endColumns(re, k, b, top, blocks, rec, sectEnds, starts);
        if (e.moved != null) { shift += e.moved; top += e.moved; p = pageOf(top); }
        if (e.again) balanceAgain = true;
      }
      // (Below it, its space after and the note separator's line as Word draws it, when a note has to start under it.)
      reserveNotes(b.i, p, b.bottom + shift + spaceAfter(b.node) + 16);
      // A new section's margins: from this page if it starts one, else from the next.
      while (margins.length > 1 && margins[margins.length - 1][0] > p + 1) margins.pop(); // left by an earlier pass over these blocks
      const sect = k > 0 && sectEnds.get(blocks[k - 1].i);
      if (sect) {
        const from = sect.page ? (top <= contentTop(p) + 1 ? p : p + 1) : p + 1;
        setMargins(from, sect.next, sect.page);
        while (sections.length > 1 && sections[sections.length - 1][0] >= from) sections.pop();
        sections.push([from, sect.index]);
      }
      const firstOnPage = top <= contentTop(p) + 1;
      // A columns section: the first column starts where the section does (on a new page, at its top), the others beside
      // it at that height (flow pages P+1.. of the same sheet), then whole sheets of columns. Balanced, the columns of
      // its last sheet (as the last pass found it) end at the height worked out for them, all but the last.
      const rs = regs.find((r) => r.k0 === k);
      if (rs) startColumns(rs, regs, sect, firstOnPage, p, top);
      const inReg = regs.find((r) => k >= r.k0 && k < r.k1);
      const inCols = cols || !!inReg;
      if (inReg) rec.set(k, { top, bottom: b.bottom + shift });
      // An empty paragraph that only carries a section break may hang into the bottom margin: Word doesn't start a page for it.
      // Nor after a manual page break: the page break and the section break then make one new page, not two.
      if (!firstOnPage && !b.node.childCount && sectEnds.get(b.i) && top < contentBottom(p)) { if (sectEnds.get(b.i).page) breakPending = false; continue; }
      // (In columns of their own widths, a paragraph's bottom at the width of the column it's in: atColumnWidth.)
      const wide = inCols ? atColumnWidth(b, cols || inReg, p, shift, top <= contentTop(p) + 1) : null;
      const bottom = wide ? wide.bottom : b.bottom;
      let target = null;
      let overflow = false;
      let forced = false;
      let sectStart = false;
      if (k > 0 && (breakPending || breakBefore(b.node) || (sect && sect.page)) && !firstOnPage) {
        target = p + 1;
        forced = !breakPending; // after a manual page break Word drops the space before, as at a natural break
        sectStart = forced && !!(sect && sect.page) && !breakBefore(b.node);
      }
      else if (p === 0 && k === 0 && top < contentTop(p) + spaceBefore(b.node) - 0.5) {
        // Under a header that reaches past the top margin, the first paragraph keeps its space before: Word moves the body down whole.
        // (Checked first: under a header reaching further than the space before, the paragraph would otherwise lose it.)
        target = p;
        forced = true;
      }
      else if (top < contentTop(p)) target = p; // landed in a gap or a top margin
      else if (inCols && b.node.type === schema.nodes.paragraph && !hasPageBreak(b.node) && bottom + shift > contentBottom(p) + lineSlack(b.el) - 1) {
        // In columns a paragraph goes on in the next column line by line, as in Word: it stays put, cut at the column's
        // bottom, and its lines from there are drawn again at the next column's top (drawHeads); what follows it moves
        // down below them.
        // (Laid out at another column's width: split on the next pass, at this one's, when it can be; else it moves on.)
        if (wide?.splits) continue;
        const cs = wide ? null : colSplit(b, shift, firstOnPage, p);
        if (cs) {
          if ((cols || inReg).unequal) rewrapCarry(b, cs, cols || inReg, p);
          colSplits.set(b.i, cs);
          if (inReg) rec.set(k, { top, bottom: top + cs.cut, carry: (cs.rewrap ? cs.rewrap.h : cs.height - cs.cut) + (cs.head || 0) });
          const nb = blocks[k + 1];
          if (nb) { starts.set(nb.i, (starts.get(nb.i) || 0) + cs.delta); shift += cs.delta; }
          continue;
        }
        if (!firstOnPage && b.bottom - b.top <= pageHeight(p + 1)) {
          target = p + 1;
          overflow = true;
        }
      }
      else if (inCols && b.node.type === schema.nodes.table && b.bottom + shift > contentBottom(p) + 1) {
        const cs = colSplitTable(b, shift, p);
        if (cs) {
          colSplits.set(b.i, cs);
          if (inReg) rec.set(k, { top, bottom: top + cs.cut, carry: cs.height - cs.cut + (cs.head || 0) });
          const nb = blocks[k + 1];
          if (nb) { starts.set(nb.i, (starts.get(nb.i) || 0) + cs.delta); shift += cs.delta; }
          continue;
        }
        if (!firstOnPage) {
          target = p + 1;
          overflow = true;
        }
      }
      else if (lineSplits && !inCols && b.node.type === schema.nodes.paragraph && (hasPageBreak(b.node) || b.bottom + shift > contentBottom(p) + lineSlack(b.el) - 1 - tie(b.el))) {
        // Paragraphs continue on the next page line by line (and after a page break), as in Word.
        const split = splitParagraph(b, shift, firstOnPage);
        if (split && split.pushes.length) {
          paraSplits.set(b.i, split);
          notesByLine(b, shift, split);
          shift += split.total;
        }
        if (split) {
          breakPending = split.breakAfter;
          if (blocks[k + 1]) { if (split.tail) tails.set(blocks[k + 1].i, split.tail); else tails.delete(blocks[k + 1].i); }
          continue;
        }
        // (One taller than a page moves too, and splits on its new page.)
        if (split === null && !firstOnPage) {
          target = p + 1;
          overflow = true;
        }
      } else if (!(lineSplits && b.node.type === schema.nodes.paragraph && b.node.textContent.trim()) && (bottom + shift > contentBottom(p) + 1 || (b.rows && rowBreakIn(b.node)))) {
        // (A paragraph of text that passed the line test above stays: only the extra below its text passes the margin.
        // One holding only a picture has no such extra, and moves.)
        // Tables continue on the next page row by row, as in Word; other blocks move whole.
        const split = b.rows && splitTable(b, shift, firstOnPage);
        if (split) {
          splits.set(b.i, split);
          shift += split.total;
          breakPending = false;
          continue;
        }
        // A table whose first row can't start here moves to the next page even when it's longer than a page (it splits there).
        if (!firstOnPage && bottom + shift > contentBottom(p) + 1 && (b.rows || b.bottom - b.top <= pageHeight(p + 1))) {
          target = p + 1;
          overflow = true;
        }
      }
      breakPending = !lineSplits && hasPageBreak(b.node);
      if (target == null) continue;
      // Keep with next: headings at the bottom of a page move with the block they introduce.
      let m = k;
      if (overflow) {
        // (In newspaper columns, also from the top of a column that isn't its page's first, moved there already: Word
        // takes a paragraph on with the one it keeps with column by column, to the next page's first column at most.
        // The generated documents' headings in two and three columns, tools/doc-generator.)
        const midSheet = !!cols && p % cols.num !== 0;
        while (m > 0 && keep(blocks[m - 1].node) && (midSheet ? blocks[m - 1].top + shift >= contentTop(p) - 1 : !starts.has(blocks[m - 1].i) && blocks[m - 1].top + shift > contentTop(p) + 1)) m--;
      }
      const moved = blocks[m];
      // With "page break before" or a new section Word keeps the paragraph's space before.
      // At a new section's page, less the space after of the paragraph ending the section before (sample S097:
      // a Title spaced 9pt before, under a section ended by a 3pt-after paragraph, 6pt down in Word's own layout; under
      // a 6pt one, 3pt; varied in copies of it, a space before up to that space after leaves none).
      const kept = forced && sectStart && m === k && k > 0 ? Math.max(0, spaceBefore(moved.node) - spaceAfter(blocks[k - 1].node)) : spaceBefore(moved.node);
      const delta = contentTop(target) + (forced ? kept : tails.get(moved.i) || 0) - (moved.top + shift);
      starts.set(moved.i, (cols || carried.has(moved.i) ? starts.get(moved.i) || 0 : 0) + delta); // (in columns it may already carry a split's push; after a frame moved on, that push)
      shift += delta;
      if (rec.has(m)) rec.set(m, { top: moved.top + shift, bottom: moved.bottom + shift }); // (the walk goes on after it, not from it)
      for (let n = m; n <= k; n++) {
        moveNotes(blocks[n].i, target);
        splits.delete(blocks[n].i);
        paraSplits.delete(blocks[n].i);
        colSplits.delete(blocks[n].i);
      }
      k = m; // re-check everything after the moved block
      breakPending = !lineSplits && hasPageBreak(moved.node);
      if (lineSplits && moved.node.type === schema.nodes.paragraph && (hasPageBreak(moved.node) || moved.bottom + shift > contentBottom(target) + 1)) {
        const split = splitParagraph(moved, shift, true, true);
        if (split && split.pushes.length) {
          paraSplits.set(moved.i, split);
          notesByLine(moved, shift, split);
          shift += split.total;
        }
        if (split) breakPending = split.breakAfter;
      }
      // A moved table may still need splitting on its new page.
      if (moved.rows && (moved.bottom + shift > contentBottom(target) + 1 || rowBreakIn(moved.node))) {
        const split = splitTable(moved, shift, true);
        if (split) {
          splits.set(moved.i, split);
          shift += split.total;
        }
      }
    }
    if (!cols && !regs.length) verticalAlign(doc, blocks, starts, pulls, splits, paraSplits);
    const last = blocks.length ? blocks[blocks.length - 1].bottom + shift : 0;
    pages = Math.max(1, pageOf(Math.max(last - 1, 0)) + 1);
    // (Notes continued past the body's last page: pages of their own, as Word adds.)
    if (noteFlow.size) pages = Math.max(pages, Math.max(...noteFlow.keys()) + 1);
    const { colCss, contHeads } = placeColumns(blocks, regs, starts, colSplits, paraSplits, splits);

    const { css, heads } = applyPushes(doc, dom, blocks, base, colRegs, starts, pulls, splits, paraSplits);
    css.push(`.${scope} > .odx-content{min-height:${pageTop(pages) - gap}px}`);
    if (cols || colRegs.length) dom.classList.remove('odx-measure');
    sheet.classList.remove('odx-bl-off');
    const frameCss = [...frameDy].map(([i, dy]) => `.${scope} > .odx-content > :nth-child(${i + 1}){--odx-fdy:${dy.toFixed(2)}px}`);
    let text = css.concat(colCss, frameCss, baselineCss(blocks, scope, sheet)).join('\n');
    text += `\n@page { size: ${page.width}px ${page.height}px; margin: 0; }\n@media print { .${scope} { height: ${pages * step(page.height) - 1}px; overflow: hidden; } }`; // the browser's own print margins would shrink and shift each sheet
    if (rules.textContent !== text) rules.textContent = text; // unchanged rules: no style recalculation
    drawSheets();
    drawHeads(heads.concat(contHeads));
    drawLineNumbers(view, lineNums, pageOf);
    drawBalloons(view, session, sheet, pageTop(pages) - gap);
    lastMs = performance.now() - t0;
    passes++;
    // A balanced columns section's cut changed: lay out again with it (a few times at most, should it not settle).
    if (balanceAgain && balancePasses < 4) { balancePasses++; schedule(); }
    // Blocks moved into columns of other widths: again at those widths (a few times at most, should it not settle).
    else if (colsAgain && colsPasses < 8) { colsPasses++; colsRepass = true; schedule(); }
    colsAgain = false;
  }

  let drawn = '';
  let pageBorder = null;
  /** w:pgBorders of the document's section, as a box drawn on every sheet. */
  function readPageBorder() {
    const x = (view.state.doc.attrs.sectPr || '').match(/<w:pgBorders\b[\s\S]*?<\/w:pgBorders>/)?.[0];
    if (!x) return null;
    const fromText = /w:offsetFrom="text"/.test(x);
    const el = document.createElement('div');
    el.className = 'odx-pgborder';
    const css = [];
    for (const side of ['top', 'left', 'bottom', 'right']) {
      const m = new RegExp(`<w:${side}\\b[^>]*>`).exec(x);
      if (!m || /w:val="(none|nil)"/.test(m[0])) continue;
      const a = (n, d) => (new RegExp(`w:${n}="([^"]*)"`).exec(m[0]) || [])[1] ?? d;
      const w = Math.max(1, Number(a('sz', 4)) / 8 * 4 / 3);
      const color = /^[0-9A-Fa-f]{6}$/.test(a('color', '')) ? `#${a('color')}` : '#000';
      const style = a('val') === 'double' ? 'double' : /dot/.test(a('val')) ? 'dotted' : /dash/.test(a('val')) ? 'dashed' : 'solid';
      const space = Number(a('space', 24)) * 4 / 3;
      const margin = side === 'top' ? page.top : side === 'bottom' ? page.bottom : side === 'left' ? page.left : page.right;
      css.push(`${side}:${fromText ? margin - space - w : space}px`, `border-${side}:${style === 'double' ? Math.max(3, w * 3) : w}px ${style} ${color}`);
      if (/w:shadow="(1|true|on)"/.test(m[0]) && (side === 'right' || side === 'bottom')) css.push('box-shadow:3px 3px 0 #000');
    }
    el.style.cssText = css.join(';');
    return el;
  }
  let refsInStories = null;
  let docSeq = 0;
  let seqDoc = null;
  function drawSheets() {
    if (seqDoc !== view.state.doc) { seqDoc = view.state.doc; docSeq++; }
    pageBorder = readPageBorder();
    // (Headers and footers with reference fields are drawn again as the document changes: a running head follows
    // its heading's text.)
    if (refsInStories == null) refsInStories = /\b(?:STYLEREF|PAGEREF|REF)\b/.test(Object.keys(session.pkg.files).filter((k) => /word\/(header|footer)\d*\.xml$/.test(k)).map((k) => session.pkg.text(k)).join(''));
    const key = `${refsInStories ? view.state.doc.content.size + ':' + docSeq + ':' : ''}${(view.state.doc.attrs.sectPr || '').match(/<w:pgBorders\b[\s\S]*?<\/w:pgBorders>/)?.[0] || ''}:${pages}:${margins.map((m) => m[0] + (m[2] ? 'f' : '')).join(',')}:${pitch}:${pageTop(pages)}:${[...notesOn].map(([p, l]) => p + '=' + l.map((r) => r.kind[0] + r.id + '.' + r.n).join('+')).join(',')}`;
    if (drawn === key) return;
    drawn = key;
    const frag = document.createDocumentFragment();
    for (let p = 0; p < pages; p++) {
      const s = document.createElement('div');
      s.className = 'odx-sheet';
      // (With what its header's reference fields need: STYLEREF, PAGEREF and REF are worked out for this page.)
      const numText = { ...numberOf(p), resolve: (instr) => referenceValue(view, instr, { page: p }) };
      const num = numText.n;
      if (pageBorder) s.append(pageBorder.cloneNode());
      // Columns with a line between (w:cols w:sep): a thin rule midway in each gap, the text's height, as Word draws it.
      if (cols?.sep) {
        const m = marginsAt(p);
        for (let k = 1; k < cols.num; k++) {
          const line = document.createElement('div');
          line.className = 'odx-colsep';
          const mid = cols.xs[k] - cols.spaces[k - 1] / 2; // (midway in the gap before column k, from the text's left)
          const x = cols.rtl ? page.width - page.left - page.right - mid : mid;
          line.style.cssText = `position:absolute;left:${(page.left + x).toFixed(1)}px;top:${m.top}px;height:${pageH(p) - m.top - m.bottom}px;border-left:0.75pt solid #000`;
          s.append(line);
        }
      }
      s.style.cssText = `top:${pageTop(p)}px;height:${pageH(p)}px${pageW(p) !== page.width ? `;width:${pageW(p)}px` : ''}`;
      // The section's first page, even pages, or the default header and footer.
      let at = margins[0];
      for (const m of margins) if (m[0] <= p) at = m; else break;
      const hf = at[1].hf;
      if (hf) {
        // Even and odd go by the page's number, not its place (a section restarting at 1 makes the second sheet odd).
        const type = hf.titlePg && p === at[0] && at[2] ? 'first' : hf.even && num % 2 === 0 ? 'even' : 'default';
        const header = type === 'default' ? hf.header.default : hf.header[type];
        const footer = type === 'default' ? hf.footer.default : hf.footer[type];
        if (header) {
          const h = story(header, 'odx-header', numText, pages);
          // (In the section's own side margins: sample S106's 0.75in sections under a 1in last one.)
          h.style.cssText = `top:${hf.headerDist}px;left:${at[1].left ?? page.left}px;right:${at[1].right ?? page.right}px`;
          s.append(h);
          // Its watermark (reader.js): on the page itself, centred on the margins or where it says, behind the text.
          for (const img of h.querySelectorAll('img[style*="--odx-wm"]')) {
            const st = img.style, w = parseFloat(st.width) || 0, ht = parseFloat(st.height) || 0;
            const L = at[1].left ?? page.left, R = at[1].right ?? page.right, m = marginsAt(p);
            const spot = (pos, rel, from, size, room) => (pos === 'center' ? (rel === 'page' ? 0 : from) + ((rel === 'page' ? room : room - from - size.end) - size.self) / 2 : (rel === 'page' ? 0 : from) + Number(pos));
            const v = (k) => st.getPropertyValue(k).trim();
            st.left = `${spot(v('--odx-wm-h'), v('--odx-wm-hrel'), L, { end: R, self: w }, pageW(p))}px`;
            st.top = `${spot(v('--odx-wm-v'), v('--odx-wm-vrel'), m.top, { end: m.bottom, self: ht }, pageH(p))}px`;
            s.prepend(img);
          }
        }
        if (footer) {
          const f = story(footer, 'odx-footer', numText, pages);
          f.style.cssText = `bottom:${hf.footerDist}px;left:${at[1].left ?? page.left}px;right:${at[1].right ?? page.right}px`;
          s.append(f);
        }
      }
      const flow = noteFlow.get(p);
      if (flow) {
        const box = document.createElement('div');
        box.className = 'odx-notes';
        let top = pageH(p) - marginsAt(p).bottom - notesH.get(p);
        // Endnotes start where the text ends, as in Word (its separator line under the last paragraph's space after),
        // not at the page's foot as footnotes do: the generated documents' endnotes, tools/doc-generator.
        const last = view.dom.lastElementChild;
        if (last && !flow.skip && flow.items.every((r) => r.kind === 'endnote')) {
          const y = last.getBoundingClientRect().bottom + (parseFloat(getComputedStyle(last).marginBottom) || 0) - view.dom.getBoundingClientRect().top - pageTop(p) + ENDNOTE_LEAD;
          if (y > 0 && y < top) top = y;
        }
        box.style.cssText = `top:${top}px;left:${page.left}px;right:${page.right}px`;
        const sep = document.createElement('div');
        sep.className = 'odx-note-sep';
        // (Notes carried on from the page before: the part not yet shown, cut where this page's room ends.)
        const clip = document.createElement('div');
        clip.style.cssText = `overflow:hidden;height:${flow.h}px`;
        const inner = document.createElement('div');
        inner.style.marginTop = `${-flow.skip}px`;
        inner.append(...flow.items.map((r) => noteEl(r.note, r.kind, r.n)));
        clip.append(inner);
        box.append(sep, clip);
        s.append(box);
      }
      frag.append(s);
    }
    layer.replaceChildren(frag);
    clipExactRows(layer); // (drawn now, so they can be measured)
    fitRowspans(layer);
    // Header and footer tabs to their stops (center, right...), as the body's: they had stayed plain tab characters,
    // which collapse to a space (sample S078's centred "Company Proprietary").
    for (const hf of layer.querySelectorAll('.odx-hf')) layoutTabsIn(hf, session, (p) => { const n = hfNodes.get(p); if (!n) throw new Error('no node'); return n; });
    hfShifts(layer, sheet); // the print-only baseline shifts of header and footer text, as the body's
    const gf = document.createDocumentFragment();
    if (gap) for (let p = 0; p + 1 < pages; p++) {
      const g = document.createElement('div');
      g.style.cssText = `top:${pageTop(p) + step(pageH(p))}px;height:${gap}px`;
      gf.append(g);
    }
    gaps.replaceChildren(gf);
  }

  const schedule = () => {
    if (!frame) frame = requestAnimationFrame(layout);
  };
  // While typing, wait for a pause so pagination never slows keystrokes down.
  const scheduleIdle = () => {
    clearTimeout(timer);
    timer = setTimeout(schedule, 250);
  };
  // Re-run when anything changes size (fonts and images loading, window resizes).
  // Re-run when something changes size on its own: web fonts or images
  // finishing loading, or the window resizing. (Not a ResizeObserver on the
  // content: our own min-height changes would feed back into it.)
  const onLoad = (e) => { if (e.target.tagName === 'IMG') scheduleIdle(); };
  view.dom.addEventListener('load', onLoad, true);
  // Long documents skip laying out off-screen blocks (content-visibility) and
  // paginate on estimated heights; when a block is first drawn its real height
  // replaces the estimate, so paginate again.
  const onShown = (e) => { if (!e.skipped) scheduleIdle(); };
  view.dom.addEventListener('contentvisibilityautostatechange', onShown, true);
  window.addEventListener('resize', scheduleIdle);
  if (document.fonts) fontsSettled().then(() => !destroyed && schedule());
  // A web font that arrives later (a stand-in such as Carlito, requested only when the text using it is drawn) changes
  // what was measured in the fallback font: tabs, widths and pages are worked out again.
  const onFonts = () => { if (!destroyed) scheduleIdle(); };
  document.fonts?.addEventListener('loadingdone', onFonts);
  let destroyed = false;
  schedule();

  // Printing: sheets sit edge to edge (no gap), each exactly one printed page.
  const setGap = (g) => {
    if (g === gap) return;
    gap = g;
    pitch = step(page.height) + gap;
    drawn = '';
    layout();
  };
  // Sections with different page sizes are printed from copies of the sheet, one per page (printcopies.js).
  let printBox = null;
  let printZoom = '';
  const beforePrint = () => {
    // (Printed at full size, whatever the zoom on screen.)
    printZoom = sheet.style.zoom;
    sheet.style.zoom = '';
    // Word updates the body's fields as it prints (page references, cross-references): so do we, from the pages.
    updateFields(view, { all: true });
    setGap(0);
    // The print is taken right after this handler: finish the layout now rather than on later frames, which a page
    // printed in the background may not get (its pages kept their on-screen gaps).
    // (At least once: the fields' update can draw tables again, at their grid's widths, and the layout before had passed.)
    passed = false;
    for (let k = 0; k < 8 && !passed; k++) {
      if (frame) cancelAnimationFrame(frame);
      layout();
    }
    // Lifted floats read where pages are from the layout before (liftFloats): settle them on the gapless pages.
    if (view.dom.querySelector('img[style*="--odx-lrel"], img[style*="--odx-arel"]')) for (let k = 0; k < 2; k++) { if (frame) cancelAnimationFrame(frame); layout(); }
    if (engineOn()) buildPrintLines(sheet, view.dom, (el) => shiftFor(el, sheet), scope);
    printBox?.remove();
    printBox = printCopies(sheet, pages, { pageW, pageH, pageTop });
    if (printBox) document.body.append(printBox);
  };
  const afterPrint = () => { removePrintLines(sheet); printBox?.remove(); printBox = null; setGap(GAP); if (printZoom) sheet.style.zoom = printZoom; printZoom = ''; };
  window.addEventListener('beforeprint', beforePrint);
  window.addEventListener('afterprint', afterPrint);

  return {
    update: scheduleIdle,
    balloons: () => unzoomed(() => drawBalloons(view, session, sheet, pageTop(pages) - gap)),
    /** Headers or footers changed: measure and draw them again. */
    restory() {
      sectCache = { doc: null, ends: new Map(), first: page };
      heights.clear();
      noteHeights.clear();
      drawn = '';
      schedule();
    },
    /** Page index at y px (from the top of the sheet), and its header/footer set and type. */
    pageAt(y) {
      const p = Math.max(0, Math.min(pages - 1, pageOf(y)));
      let at = margins[0];
      for (const m of margins) if (m[0] <= p) at = m; else break;
      const hf = at[1].hf;
      // (Even by the page's number, as drawSheets draws it: after a section restarting at 1 that isn't its place.)
      const type = !hf ? 'default' : hf.titlePg && p === at[0] && at[2] ? 'first' : hf.even && numberOf(p).n % 2 === 0 ? 'even' : 'default';
      let sect = 0;
      for (const [from, index] of sections) if (from <= p) sect = index; else break;
      return { p, y: y - pageTop(p), top: pageTop(p), hf, type, sect, m: at[2] && at[1].first ? at[1].first : at[1], page, pitch };
    },
    setGap,
    /** Page p's number and its format ({ n, fmt }), as its header shows it (sections may restart and format it). */
    numberAt: (p) => numberOf(p),
    /** The number of pages laid out. */
    get count() { return pages; },
    /** The print layer (each line placed exactly, with its print baseline shift) on screen, for the PDF export to read;
     *  returns a function removing it. */
    printLayer() {
      if (!engineOn()) return () => {};
      buildPrintLines(sheet, view.dom, (el) => shiftFor(el, sheet), scope);
      return () => removePrintLines(sheet);
    },
    pageTop: (p) => pageTop(p),
    get pages() { return pages; },
    get lastLayoutMs() { return lastMs; },
    get layoutPasses() { return passes; },
    destroy() {
      destroyed = true;
      view.dom.removeEventListener('load', onLoad, true);
      view.dom.removeEventListener('contentvisibilityautostatechange', onShown, true);
      window.removeEventListener('resize', scheduleIdle);
      document.fonts?.removeEventListener('loadingdone', onFonts);
      window.removeEventListener('beforeprint', beforePrint);
      window.removeEventListener('afterprint', afterPrint);
      clearTimeout(timer);
      if (frame) cancelAnimationFrame(frame);
      rules.remove();
      layer.remove();
      repeats.remove();
      gaps.remove();
    },
  };
}
