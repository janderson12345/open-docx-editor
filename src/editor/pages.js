// Print-layout view: splits the continuous editor column into Word-style pages.
//
// The editable content stays one ProseMirror column. After each change we
// measure the top-level blocks, decide which block starts each page, and push
// those blocks down: margin-tops set as node decorations (applyPushes), and a
// generated stylesheet for what decorations can't say (split tables' rows,
// columns' moves). The few writes the layout makes into the editor's own
// elements (tab widths, a floating table set in the flow, pictures drawn up)
// go through quiet(), so the editor doesn't take them for edits. A layer
// behind the text draws the page sheets, headers and footers.
//
// Paragraphs split line by line and tables row by row (rows inside their cells),
// with spacers inserted as view decorations. Each section's margins and page
// size apply from the page it starts on (setMargins).
import { runOnLead } from './joins.js';
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { schema } from './schema.js';
import { layoutTabs, layoutTabsIn } from './tabs.js';
import { fitRowspans, clipExactRows } from './rowfit.js';
import { storyDrawer, hfNodes, settleStoryShapes, flowHeight, layoutPtabs } from './hfstory.js';
import { printCopies } from './printcopies.js';
import { sectionBreaks, columnsIn, sectionXmls } from './sections.js';
import { drawLineNumbers } from './linenumbers.js';
import { fontsSettled } from './fontcheck.js';
import { referenceValue, updateFields } from './refs.js';
import { lineHeightOf, layoutRect } from './util.js';
import { drawBalloons } from './balloons.js';
import { noteRefs } from './notes.js';
import { createNoteBoxes, CONT_LEAD } from './notebox.js';
import { applySqueeze } from './justify.js';
import { applyLineFix, lineGrid, storyLineFix } from './lhfix.js';
import { applyWidths, fixStoryWidths } from './widths.js';
import { applyFitText } from './fittext.js';
import { applyLabelWords } from './lblword.js';
import { applyWrapBoth, drawnUpTo } from './wrapboth.js';
import { applyTableClear } from './tableclear.js';
import { settleHyphens } from './hyphenate.js';
import { fitAutofitColumns } from './tablefit.js';
import { baselineCss, blockShifts, cellShifts, hfShifts, shiftFor } from './baseline.js';
import { buildPrintLines, removePrintLines, engineOn } from '../engine/printlines.js';
import { twipsToPx } from '../ooxml/units.js';
import { placeVar, padOf } from '../ooxml/place.js';
import { styleHeaderRows } from '../ooxml/styles.js';

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
// The browser's own grid, steps a px: 64 in Chromium and WebKit (their layout unit), 60 in Firefox (its app units), found
// once by laying out a box 1/64px tall. (On 64ths in Firefox, a push was laid out at the nearest 60th and measured back
// a hair off, and the next pass put it a 64th the other way: a 122-page document's pushes went back and forth pass after
// pass, 28 passes to open there and 14 more to its full layout, against Edge's 10 and 3; 7 and 8 on its own grid.)
let perPx = 0;
const gridSteps = () => {
  if (perPx || typeof document === 'undefined' || !document.body) return perPx || 64;
  const box = document.createElement('div');
  box.style.cssText = 'position:absolute;left:0;top:0;width:1px;height:0.015625px;visibility:hidden;pointer-events:none';
  document.body.append(box);
  const h = box.getBoundingClientRect().height;
  box.remove();
  perPx = Math.abs(h - 1 / 60) < 1e-4 ? 60 : 64; // (Firefox reports it through a single-precision float: 0.0166626)
  return perPx;
};
// A push in px, on the browser's own grid: laid out exactly as set, and measured back the same at the next pass.
// (In tenths, one near the middle of two came to one or the other by the last bits of the sum or by where an earlier pass
// had put it, and a print came out by how many passes had run before it. To a thousandth of a grid step first: a value
// exactly on the grid stays there, whatever its last bits.)
const toGrid = (px) => { const g = gridSteps(); return Math.round(Math.round(px * g * 1000) / 1000) / g; };
// A length from the style as the browser lays it out: cut to its 64ths in Chromium, to the nearest 60th in Firefox.
const asLaidOut = (px) => { const g = gridSteps(); return g === 64 ? Math.trunc(px * 64) / 64 : Math.round(px * g) / g; };
// A box's edges as the browser laid them out. Firefox gives them through single-precision floats: 100,000px down a long
// document, a few thousandths of a pixel off its own 60ths, and summed over the pushes above a block (what a pass takes
// off to find where it would be) a 60th off from one pass to the next, so a split row's spacer that fell near half a
// 60th went one way and then the other, pass after pass (a 122-page document's tables: 6 passes more than in Edge).
// Put back on its grid there (exact to 262,144px down, where the float is still within half a 60th); elsewhere the
// browser's own rect, as it was.
const onGrid = (v) => Math.round(v * 60) / 60;
const snapRect = (r) => {
  const top = onGrid(r.top), bottom = onGrid(r.bottom), left = onGrid(r.left), right = onGrid(r.right);
  return { top, bottom, left, right, x: left, y: top, width: right - left, height: bottom - top };
};
const rectOf = (el) => { const r = el.getBoundingClientRect(); return gridSteps() === 60 ? snapRect(r) : r; };
const rectsOf = (range) => { const l = range.getClientRects(); return gridSteps() === 60 ? [...l].map(snapRect) : l; };
// A font's line and a paragraph's multiple as Word has them, unrounded, from the four decimals they're kept in: a font's
// line in 2048ths of an em (its units), a multiple in 240ths (w:line). Word fits a page's last line by their exact sum,
// no line rounded: its prints of 156 generated pages, Times New Roman and Calibri at 10 and 12pt, 1.08, 1.15 and 1.5
// lines, the space that still keeps the last line predicted to the twip in all 12 kinds. (A value that isn't one: as it is.)
const nearest = (v, per) => { const u = Math.round(v * per) / per; return Math.abs(u - v) < 0.00006 ? u : v; };
const exactEm = (v) => nearest(v, 2048);
const exactMult = (v) => nearest(v, 240);
const spacer = (h, end = false) => () => {
  const el = document.createElement('span');
  el.className = end ? 'odx-push odx-push-end' : 'odx-push';
  el.contentEditable = 'false';
  el.style.height = `${h}px`;
  return el;
};

/** Where a box `h` tall goes between `top0` and `bottom0` for Word's vertical align (top, center, bottom), else `off` below top0. */
const alignedY = (align, top0, bottom0, h, off) => (align === 'top' ? top0 : align === 'bottom' ? bottom0 - h : align === 'center' ? (top0 + bottom0 - h) / 2 : top0 + off);

/** Adds `d` to a map's number at `k` (0 when it has none): the pushes the layout pass adds up per block. */
const addTo = (map, k, d) => map.set(k, (map.get(k) || 0) + d);
// A block's table: the element itself, or the table inside its wrapper (prosemirror-tables' tableWrapper, a floating
// table's frame); a table in a cell, only as a child or in a tableWrapper (a note's: notebox.js ownTable).
const tableOf = (el) => (el.tagName === 'TABLE' ? el : el.querySelector('table'));
const cellTable = (el) => (el.tagName === 'TABLE' ? el : el.classList.contains('tableWrapper') ? el.querySelector(':scope > table') : null);
// A shape with text on both sides, set out of the flow (reader.js anchorCss, --odx-both): it takes no line of its own,
// so it is no line when a paragraph is split over pages, nor where one starts. Counted as one, it made the lines beside
// it a single line the page couldn't split, the paragraph moved on whole; Word breaks such a paragraph as any other,
// the shape staying with its first lines (wrap-word3 vt-thin: Text 7's three lines beside its rectangle at page 1's
// foot, the rest on page 2 the column's width).
// A paragraph with nothing on its line: no content, or only drawings out of the line (floating or set where they are,
// reader.js anchorCss), as a section's closing paragraph holding the cover page's pictures (Word kept such a paragraph,
// 12px past the footer's room, on its page; moved on alone, it made a blank page before the next section).
// (Nor do empty markers, its bookmarks: a sample's section-break paragraph after a form's table, holding two, had a page of
// its own, every page after it a page late. Nor a field's marks with nothing drawn, a table of contents' end mark in the
// paragraph ending its section: Word's print of a sample, the contents filling their page to the footer and that
// paragraph's line past it, no page added.)
const lineless = (node) => {
  let only = true;
  node.forEach((c) => {
    const marker = c.type === schema.nodes.opaque_inline && !c.attrs.text && c.attrs.width == null;
    if (!marker && !(c.type === schema.nodes.image && /(?:^|;)\s*(?:position:absolute|float:(?:left|right))/.test(c.attrs.place || ''))) only = false;
  });
  return only;
};
const bothShape = (e) => e.tagName === 'IMG' && !!e.style.getPropertyValue('--odx-both') && getComputedStyle(e).position === 'absolute';

const GAP = 24; // px between sheets
const TIE = 0; // (was 0.33px: fitted when spacing was cut to the browser's 64ths of a pixel, a third of a pixel short a page)
// Guards, not fitted values: loops that walk columns or pages stop here, should a layout never settle.
const MAX_CARRY = 64; // columns a paragraph or table is carried over, and flow pages a note band is looked for past its page
const MAX_FLOW_PAGES = 8192; // flow pages walked to find each sheet's note band (placeColumns)
const NOTE_RUNAWAY = 500; // pages of notes past the last reference (sumNotes)
// The least room a page gives its notes, whatever its text takes (since 0.5.0; about two lines of a 10pt note), and a
// table set down near its page's foot ends the page's text above it when less than this (two lines) is left under it.
const MIN_ROOM = 40;
// Clip paths cut a block top and bottom only: they reach this far out at the sides (italic overhang, a list's number
// in the margin) and above the top, and the far edge of a split table's clip.
const CLIP_SIDE = 200, CLIP_TOP = 20, CLIP_FAR = 99999, CLIP_DEEP = 9999999;
// The height a one-section document's columns are cut to before its endnotes: an even share plus this, so the first
// column takes the odd line (balanceEndCols; 9 until 448a68b, raised to 12 for the generated documents' endnote-content
// in two and three columns). A balanced columns section measures half its line instead (endColumns).
const END_COLS_EXTRA = 12;

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
    touchNotes(Math.max(0, from - 1)); // (the notes' bands from there on: their pages' room changed)
  };
  // The margins entry in force on page p: the last section start at or before it.
  const entryAt = (p) => {
    let e = margins[0];
    for (const m of margins) if (m[0] <= p) e = m; else break;
    return e;
  };
  // A section's first page can have its own (first-page) header and footer, so its own margins.
  const marginsAt = (p) => {
    const e = entryAt(p);
    return e[2] && e[0] === p && e[1].first ? e[1].first : e[1];
  };
  // Sections can have their own page height, so pages don't all sit one pitch apart.
  const pageH = (p) => entryAt(p)[1].height || page.height;
  // A page's number and its format: one more than the page before, or the section's start where it restarts.
  // (One walk over the pages and the section starts together, not a search of the starts for each page.)
  const numberOf = (p) => {
    let num = 0, fmt, j = 0;
    for (let k = 0; k <= p; k++) {
      while (j + 1 < margins.length && margins[j + 1][0] <= k) j++;
      const st = margins[j];
      num = st[0] === k && (st[2] || k === 0) && st[1].numStart != null ? st[1].numStart : num + 1;
      fmt = st[1].numFmt;
    }
    return { n: num, fmt };
  };
  // And their own page width (a landscape section among portrait ones).
  const pageW = (p) => entryAt(p)[1].width || page.width;
  // (`ms`: other page setups than the pages' own, placeColumns' by sheet.)
  const pageTop = (p, ms = margins) => {
    let y = 0;
    for (let k = 0; k < ms.length; k++) {
      const from = ms[k][0];
      const to = k + 1 < ms.length ? Math.min(ms[k + 1][0], p) : p;
      if (to > from) y += (to - from) * (step(ms[k][1].height || page.height) + gap);
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
  // (With notes at its foot, the body's text keeps clear of the separator by a few px more: Word's separator is a line
  // of its own, and a line of body text ending 4px above our note box went to the next page in Word.)
  // (In columns, not for endnotes: they come after the text, under its columns as they end, or on the next page.)
  const bandH = (p) => {
    const b = bandOf(p);
    if (colGeo(p) && noteFlow.get(b)?.items.every((r) => r.kind === 'endnote')) return 0;
    return Math.max(notesH.get(b) || 0, colGeo(p) ? notesHPrev.get(b) || 0 : 0);
  };
  // Footnotes at the foot of page p (not endnotes, which follow the text).
  const footBand = (p) => !!bandH(p) && !!noteFlow.get(bandOf(p))?.items.some((r) => r.kind === 'footnote');
  // Over a band of notes, a paragraph's last line keeps its space after above the band, as in Word: in columns, and over
  // footnotes in one, where multiple spacing's extra under a line doesn't pass the band's top either, as it may the
  // margin (longfn-word2 fnhuge page 2, where Word moves Text 6's last two lines on; longfn-word fnlate's third page).
  const bandAfter = (p, node) => (bandH(p) && (colGeo(p) || footBand(p)) ? spaceAfter(node) : 0);
  const contentBottom = (p) => {
    const h = bandH(p);
    return Math.min(pageTop(p) + pageH(p) - marginsAt(p).bottom - (h ? h + noteGap() : 0), cuts.get(p) ?? Infinity);
  };
  // Over the band, the room its separator takes past the NOTE_SEP drawn in it (notebox.js noteSepRoom): 4px for 11pt
  // Calibri, single spaced; 5.5px for 12pt; 0 for small fonts.
  let gapPx = null;
  const noteGap = () => (gapPx ??= Math.max(0, notes.noteSepRoom() - NOTE_SEP));
  // Footnotes: which go at the bottom of each page, and the room they take there.
  let notesOn = new Map(); // page -> [{ kind, id, n }]
  let notesH = new Map(); // page -> px
  let cuts = new Map(); // page -> y: a table set down near its page's foot ends the text there (pagination's band)
  // Footnotes in newspaper columns sit in one band across the page's foot, each column on the sheet cut at its top, as
  // Word (2013 on) lays them out: they're kept by that band, keyed by the flow page of the sheet's first column. (Out of
  // columns, a band is its page.)
  let curRegs = []; // this pass's columns sections (colRegions), as laid out so far
  let notesHPrev = new Map(); // band -> px, as the last pass ended: a column laid out before a note further on in its sheet
  let notesHPrev2 = new Map(); // and the pass before
  // Bands of footnotes in columns too tall for their page, as the last pass ended: band -> { cut (px from the columns'
  // top: the sheet's text up to the reference evened out across them), top (where the notes start, the same) }.
  let heldPrev = new Map();
  const colGeo = (v) => (cols ? { P: 0, ...cols } : curRegs.find((r) => r.P != null && v >= r.P && (r.vEnd == null || v <= r.vEnd)));
  // The sheet a flow page is on (columns sections take several flow pages a sheet), and the first flow page from `v`
  // on whose sheet has the parity an odd- or even-page section break asks for (sheets counted from 1).
  const sheetOfFlow = (v) => {
    if (cols) return Math.floor(v / cols.num);
    let o = 0;
    // (A section of columns the walk hasn't placed, its first block a frame or a table set down on its page, is passed
    // over: those after it still count. placeColumns had a copy of this that did; this one stopped there.)
    for (const r of curRegs) { if (r.P == null) continue; if (v < r.P) break; if (r.vEnd == null || v <= r.vEnd) return r.S0 + Math.floor((v - r.P) / r.num); o = r.offAfter; }
    return v - o;
  };
  const parityPage = (v, parity) => {
    const odd = (sheetOfFlow(v) + 1) % 2 === 1;
    if (odd === (parity === 'oddPage')) return v;
    return cols ? v + cols.num - (v % cols.num) : v + 1;
  };
  // (Word 2010 and before, compatibility mode under 15: each column has a note area of its own, under it, its column
  // alone cut at it: Word's print of a compat-14 document of two columns.)
  const perCol = () => !(session.compatMode >= 15);
  const bandOf = (v) => { const g = !perCol() && colGeo(v); return g ? g.P + Math.floor((v - g.P) / g.num) * g.num : v; };
  // A flow page's column in its section of columns: { c, x (from the text's left), w }, or null.
  const colAt = (v) => {
    const g = colGeo(v);
    if (!g) return null;
    const c = (((v - g.P) % g.num) + g.num) % g.num;
    return { c, x: colX(g, c), w: g.widths[c] };
  };
  // Column c's left edge from the text's: a right-to-left section (w:bidi) fills its columns from the right, as in Word.
  const colX = (g, c) => (g.rtl ? page.width - page.left - page.right - g.xs[c] - g.widths[c] : g.xs[c]);
  const nextBand = (b) => { let q = b + 1; while (bandOf(q) !== q && q < b + MAX_CARRY) q++; return q; };
  // The band's own columns: the section's, or as many as w15:footnoteColumns says (1: one band the text's width), equal
  // widths, in the section's direction; null for one.
  const bandCols = (b, first) => {
    if (first && first.inCols === false) return null; // (its first note's reference is above a continuous break into columns)
    if (perCol()) { const at = colAt(b); return at ? { m: 1, gap: 0, w: at.w, rtl: false } : null; } // (its column's own)
    const g = colGeo(b); const m = g ? g.fnCols ?? g.num : 1;
    return m > 1 ? { m, gap: g.fnGap ?? g.space, rtl: !!g.rtl } : null;
  };
  let sheetBand = null; // sheet -> its band (columns), or null: the same
  const NOTE_SEP = 14; // separator line and the space around it
  const ENDNOTE_LEAD = 4; // Word's separator under the text is a line of the Normal style, a little taller than NOTE_SEP
  // Endnotes starting a page of their own: their separator 7pt under the top margin, as Word prints it (fncols-word end2-ltr).
  const ENDNOTE_TOP = 6;
  const notes = createNoteBoxes({ page, sheet, session, lineFix: (el, nodeOf) => { if (!lineGrid(view.state.doc)) storyLineFix(el, nodeOf); } });
  const { noteEl, noteSepRtl, noteHeight, bandLayout, bandBox } = notes;
  // The notes drawn at each page's foot: those carried on from the pages before (a note longer than the room a page has
  // for it continues at the foot of the next, as in Word), then the page's own; `skip` px of the first already shown.
  let noteFlow = new Map(); // page -> { items, skip, h }
  const NOTE_MIN_BODY = 48; // the room a page keeps for its body text (its reference's line, at least)
  // Worked out again from the first band whose notes changed since the last time (touchNotes), not from the first band
  // of all: a band's result depends only on the bands before it, through the notes they carry on. Every reservation in
  // the layout pass works the bands out, so a long document with a note on every page had paid for all its pages on
  // each one (fncols-word many1, 400 notes: see the commit). `sumSteps`: each band worked out, in order, with what it
  // carries on.
  let sumSteps = [];
  let sumFrom = -Infinity; // the first band to work out again (Infinity: none)
  const touchNotes = (p) => { if (p < sumFrom) sumFrom = p; };
  const sumNotes = () => {
    if (sumFrom === Infinity) return;
    let k = 0;
    while (k < sumSteps.length && sumSteps[k].p < sumFrom) k++;
    for (let j = k; j < sumSteps.length; j++) { noteFlow.delete(sumSteps[j].p); notesH.delete(sumSteps[j].p); }
    sumSteps.length = k;
    if (!k) { notesH = new Map(); noteFlow = new Map(); }
    sumFrom = Infinity;
    const withNotes = [...notesOn.keys()].filter((p) => notesOn.get(p).length).sort((a, b) => a - b);
    if (!withNotes.length) return;
    const last = withNotes[withNotes.length - 1];
    const was = k ? sumSteps[k - 1] : null;
    let { carry, skip, skipKey, skipH, endShown } = was ? was.out : { carry: [], skip: 0, skipKey: '', skipH: 0, endShown: false };
    // (skipKey: the width `skip` was measured at (a band's columns: c<px>), skipH its note's height there; endShown:
    // endnotes in columns drawn on a page before.)
    const start = !was ? withNotes[0] : carry.length ? nextBand(was.p) : withNotes.find((q) => q > was.p);
    if (start == null) return;
    // (Each band's step recorded as it ends, whichever way: what it carries on.)
    let stepP = null;
    const done = () => { if (stepP != null) sumSteps.push({ p: stepP, out: { carry, skip, skipKey, skipH, endShown } }); stepP = null; };
    for (let p = start; carry.length || p <= last; done(), p = nextBand(p)) {
      const items = [...carry, ...(notesOn.get(p) || [])];
      if (!items.length) continue;
      stepP = p;
      // (Measured at this page's text width, out of columns: a landscape section's among portrait ones.)
      if (!colGeo(p)) {
        const m = marginsAt(p);
        const tw = pageW(p) - (m.left ?? page.left) - (m.right ?? page.right);
        const wide = Math.abs(tw - (page.width - page.left - page.right)) > 1;
        for (const r of items) {
          if (!r.note) continue;
          r.h = noteHeight(r.note, r.kind, r.n, wide ? tw : null);
          r.wKey = wide ? String(Math.round(tw)) : '';
        }
      }
      let room = Math.max(MIN_ROOM, pageH(p) - marginsAt(p).top - marginsAt(p).bottom - NOTE_MIN_BODY - NOTE_SEP);
      // (Its notes cut under the text when one's reference line stayed on a page too short for it: splitParagraph.)
      const part = (notesOn.get(p) || []).filter((r) => r.partial && r.at != null);
      const under = (list) => { const at0 = Math.max(...list.map((r) => r.at)), pv = pageOf(at0); room = Math.max(MIN_ROOM, Math.min(room, pageTop(pv) + pageH(pv) - marginsAt(pv).bottom - at0 - NOTE_SEP)); };
      // Too long for the page: in Word it takes all the room under the text with its reference (that text's page ends
      // there), not only what the page keeps for its body. (`at` is in the flow: in columns, its own column's page.)
      const ats = (notesOn.get(p) || []).filter((r) => r.at != null);
      // In columns of its own: the band's balanced height; too tall for the page, each column filled to the room under
      // the text and the rest carried to the next page's band (balanced there, if it can be), as Word does it.
      // (Endnotes in a one-section document of columns are in its columns too, laid out the same way.)
      const allEnd = items.every((r) => r.note && r.kind === 'endnote');
      const bc = items.every((r) => r.note && r.kind === 'footnote') ? bandCols(p, items[0])
        : allEnd && cols && !cols.unequal && cols.num > 1 ? { m: cols.num, gap: cols.space, rtl: !!cols.rtl } : null;
      if (bc && allEnd) {
        // Endnotes in columns: under the text where it ends (its columns balanced before them, balanceEndCols), or, with
        // less than two lines' room left there, from the next page's top, as in Word (fncols-word end2-ltr). On a page of
        // their own their columns run left to right, and a page going on with them has a continuation separator over its
        // first column (endlong-ltr, endlong-rtl page 3); drawSheets.
        const own = (notesOn.get(p) || []).length > 0;
        if (own) {
          const sh = colsBal && Math.floor(p / cols.num) === colsBal.sheet;
          if (sh) room = Math.max(0, pageTop(p) + pageH(p) - marginsAt(p).bottom - (contentTop(p) + colsBal.bottom + ENDNOTE_LEAD) - NOTE_SEP);
          else if (ats.length) { const at0 = Math.max(...ats.map((r) => r.at)), pv = pageOf(at0); room = Math.max(0, pageTop(pv) + pageH(pv) - marginsAt(pv).bottom - at0 - NOTE_SEP); }
          if (room < MIN_ROOM) { carry = items; skip = 0; skipKey = ''; skipH = 0; continue; }
        } else room = pageH(p) - marginsAt(p).top - marginsAt(p).bottom - NOTE_SEP - ENDNOTE_TOP;
        const lay = bandLayout(items, bc);
        const wKey = `c${Math.round(lay.w)}`;
        const from = skip && skipKey !== wKey ? skip * ((lay.notes[0][1] - lay.notes[0][0]) / (skipH || 1)) : skip;
        // (Going on from endnotes on the page before, split there or not: Word's continuation separator.)
        const cont = !own && endShown;
        const fit = notes.bandFill(lay, bc.m, from, cont ? room + NOTE_SEP + ENDNOTE_TOP : room, cont ? CONT_LEAD : 0);
        noteFlow.set(p, { items, skip: from, h: fit.h, cols: bc, lay, fit, own, cont });
        endShown = true;
        notesH.set(p, fit.h + NOTE_SEP);
        carry = items.filter((r, j) => lay.notes[j][1] > fit.end + 0.5);
        const c0 = carry.length ? lay.notes[items.indexOf(carry[0])] : null;
        skip = c0 ? fit.end - c0[0] : 0;
        skipKey = wKey;
        skipH = c0 ? c0[1] - c0[0] : 0;
        if (p > last + NOTE_RUNAWAY) { done(); break; }
        continue;
      }
      if (bc) {
        const lay = bandLayout(items, bc);
        const wKey = `c${Math.round(lay.w)}`;
        // (What of the first note was shown on the page before, measured at that page's width: in proportion here.)
        const from = skip && skipKey !== wKey ? skip * ((lay.notes[0][1] - lay.notes[0][0]) / (skipH || 1)) : skip;
        let fit = notes.bandFill(lay, bc.m, from, room);
        let held = false;
        if (fit.end < lay.total - 0.5) {
          // (Footnotes: Word evens out the sheet's text up to the reference across its columns and starts the band under
          // the lowest of them, worked out a pass behind: heldBands. Until then, under the reference's text.)
          const hp = !allEnd && heldPrev.get(p);
          if (hp && ats.length) room = Math.max(MIN_ROOM, pageTop(p) + pageH(p) - marginsAt(p).bottom - contentTop(p) - hp.top - NOTE_SEP);
          else if (part.length) under(part);
          else if (ats.length) under(ats);
          fit = notes.bandFill(lay, bc.m, from, room);
          held = !!(part.length || ats.length);
        }
        noteFlow.set(p, { items, skip: from, h: fit.h, cols: bc, lay, fit, held: held && !allEnd });
        notesH.set(p, (held && fit.end < lay.total - 0.5 ? Math.max(room, fit.h) : fit.h) + NOTE_SEP);
        carry = items.filter((r, j) => lay.notes[j][1] > fit.end + 0.5);
        const c0 = carry.length ? lay.notes[items.indexOf(carry[0])] : null;
        skip = c0 ? fit.end - c0[0] : 0;
        skipKey = wKey;
        skipH = c0 ? c0[1] - c0[0] : 0;
        if (p > last + NOTE_RUNAWAY) { done(); break; }
        continue;
      }
      // (Carried on from a band in columns: what was shown there, in proportion at this page's width.)
      if (skip && skipKey !== (items[0].wKey ?? '')) skip *= (items[0].h || 0) / (skipH || 1);
      const total = items.reduce((t, r) => t + r.h, 0) - skip;
      if (part.length) under(part);
      if (total <= room) { noteFlow.set(p, { items, skip, h: total }); notesH.set(p, total + NOTE_SEP); carry = []; skip = 0; continue; }
      if (ats.length && !part.length) under(ats);
      // Cut at the last paragraph end that fits (or, a paragraph taller than the room, at the room).
      // (A line fits when its text ends above the margin: the half leading the browser puts under it is above it in Word,
      // longfn-word2 fnland page 4. Half a px for rounding.)
      let at = 0, cut = 0;
      for (const r of items) {
        const bs = notes.boundsOf(r.note?.doc, r.wKey ?? '') || [r.h];
        const inks = notes.inksOf(r.note?.doc, r.wKey ?? '') || bs;
        bs.forEach((b, i) => { if (at + Math.min(b, inks[i] ?? b) - skip <= room + 0.5 && at + b - skip > 0) cut = at + b - skip; });
        at += r.h;
      }
      if (!cut) cut = room;
      noteFlow.set(p, { items, skip, h: cut });
      // (Under its reference's text, cut at a line's end: the band starts at the text's end, the room left below the cut, as in
      // Word (longfn-word fnhuge: not at the foot with the room above it).)
      // (A cut past the room by its last line's half leading: the band still starts under the text.)
      notesH.set(p, (ats.length ? room : Math.min(cut, room)) + NOTE_SEP);
      const rest = [];
      let acc = 0, next = 0;
      for (const r of items) {
        if (acc + r.h > skip + cut + 0.5) { if (!rest.length) next = skip + cut - acc; rest.push(r); }
        acc += r.h;
      }
      carry = rest;
      skip = next;
      skipKey = rest[0]?.wKey ?? '';
      skipH = rest[0]?.h || 0;
      if (p > last + NOTE_RUNAWAY) { done(); break; }
    }
    done();
  };
  let blockNotes = new Map(); // block index -> [{ kind, id, n, h }]
  const placed = new Map(); // note key -> page
  // (`at`: where the block holding the references ends, when known: a note too long for its page has the room below it.)
  // A note into band p's list (placed: its band).
  const putNote = (r, p) => {
    touchNotes(p);
    placed.set(r.kind + r.id, p);
    if (!notesOn.has(p)) notesOn.set(p, []);
    notesOn.get(p).push(r);
  };
  // A note out of the band it's in; false when it's in none.
  const takeNote = (r) => {
    const from = placed.get(r.kind + r.id);
    if (from == null) return false;
    touchNotes(from);
    notesOn.set(from, (notesOn.get(from) || []).filter((x) => x !== r));
    placed.delete(r.kind + r.id);
    return true;
  };
  // A placed note moved to band `to`, and the bands' heights worked out again; false when it was already there or in none.
  const moveNote = (r, to) => {
    const from = placed.get(r.kind + r.id);
    if (from == null || from === to) return false;
    takeNote(r);
    putNote(r, to);
    sumNotes();
    return true;
  };
  const reserveNotes = (i, v, at = null) => {
    const p = bandOf(v);
    const inCols = !!cols || curRegs.some((g) => i >= g.lo && i <= g.hi);
    let added = false;
    for (const r of blockNotes.get(i) || []) {
      r.inCols = inCols;
      // (In columns, where its block ends now: moved on to another column, a note's block ends elsewhere, which a band
      // too tall for its page goes by: heldBands.)
      if (placed.has(r.kind + r.id)) { if (at != null && inCols && placed.get(r.kind + r.id) === p && r.at !== at) { r.at = at; touchNotes(p); added = true; } continue; }
      r.at = at;
      putNote(r, p);
      added = true;
    }
    if (added) sumNotes();
  };
  // A screen y in block b, as it would be laid out without pagination: less where the content starts, the block's own
  // undone pushes (b.undo) and our spacers above it (paragraphLines, rewrapCarry, notesByLine).
  const naturalOf = (b) => {
    const base = rectOf(view.dom).top + b.undo;
    const pushes = [...b.el.querySelectorAll('.odx-push')].map((e) => rectOf(e));
    return (y) => y - base - pushes.reduce((sum, r) => sum + (r.top + r.height / 2 <= y ? r.height : 0), 0);
  };
  // A footnote goes on the page its reference's line does: in a paragraph split over pages, not always the paragraph's
  // first (corpus 3's c3-01: a reference in a paragraph's lines carried to page 2 has its note there, as in Word).
  const notesByLine = (b, shift0, split) => {
    const list = blockNotes.get(b.i);
    if (!list || !split.pushes.length) return;
    const natural = naturalOf(b);
    for (const r of list) {
      if (r.kind !== 'footnote') continue;
      const el = view.nodeDOM(r.pos);
      if (!el || el.nodeType !== 1) continue;
      const y0 = rectOf(el).top;
      const nat = natural(y0);
      moveNote(r, bandOf(pageOf(nat + shift0 + split.pushes.reduce((sum, x) => sum + (x.pos <= r.pos ? x.delta : 0), 0))));
    }
  };
  // In columns, a footnote whose reference is in the lines a paragraph carries into the next column goes in that
  // column's band (its sheet's, when the column is on the next page), as in Word.
  const notesByCol = (b, p, cs) => {
    const list = blockNotes.get(b.i);
    if (!list) return false;
    const top = rectOf(b.el).top;
    const parts = carriedOf(cs);
    let moved = false;
    for (const r of list) {
      // (A footnote by its reference's line, in the column of the part that holds it; endnotes, after the text, with
      // its last lines.)
      let to;
      if (r.kind === 'footnote') {
        const el = view.nodeDOM(r.pos);
        if (!el || el.nodeType !== 1) continue;
        const y = rectOf(el).top - top;
        const part = parts.filter((q) => y >= q.cut - 1).pop();
        if (!part) continue;
        to = bandOf(p + part.off);
      } else if (r.kind === 'endnote') to = bandOf(p + parts[parts.length - 1].off);
      else continue;
      if (moveNote(r, to)) moved = true;
    }
    return moved;
  };
  // Whether a footnote's reference is in a paragraph's lines from `y0` to `y1` (px from its top).
  const refsBetween = (b, y0, y1) => (blockNotes.get(b.i) || []).some((r) => {
    const el = r.kind === 'footnote' && view.nodeDOM(r.pos);
    if (!el || el.nodeType !== 1) return false;
    const y = rectOf(el).top - rectOf(b.el).top;
    return y >= y0 - 1 && y < y1 - 1;
  });
  const moveNotes = (i, v) => {
    // (Out of the bands they're in, and reserved again on `to`'s: reserveNotes works the heights out.)
    const to = bandOf(v);
    for (const r of blockNotes.get(i) || []) if (placed.get(r.kind + r.id) !== to) takeNote(r);
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
      const lo = n ? all[n - 1][0] + 1 : 0;
      // (Not balanced when it holds a column break of its own: Word leaves its columns as the break puts them.)
      let colBreak = false;
      doc.forEach((node, _, i) => { if (i >= lo && i <= hi && !colBreak && breakKindOf(node) === 'column') colBreak = true; });
      const cont = !!next && /<w:type\b[^>]*w:val="continuous"/.test(next[1]);
      out.push({ lo, hi, ...c, cont, balance: cont && !colBreak });
    });
    return out;
  };
  let colTops = new Map(); // flow page -> its content top, for the columns beside the first one on a columns section's first sheet
  const balanced = new Map(); // section's lo -> { sheet (of its own, from 0), h }: the column height its last sheet is cut to
  let balancedDoc = null;
  let balancePasses = 0;
  // Passes run again for floats set down that reach past the foot (crossFoot), since the document last changed: once a
  // table, but a chain of such tables, each crossing only after the one before goes in the flow, is a pass each; capped.
  let crossDoc = null;
  let crossPasses = 0;
  const CROSS_MAX = 8;
  let narrowed = new Map(); // block index -> px its column (of their own widths) is narrower than the widest, as laid out
  let colsAgain = false;
  let colsPasses = 0;
  let colsRepass = false; // the coming pass is one of those
  // Where the walk last put the shapes with text on both sides (their paragraphs' pushes, their drawing up or down):
  // moved, the lines beside them are split afresh (wrapboth.js) and the pages walked again, a few times at most.
  let shapesAt = '';
  let shapesPasses = 0;
  let shapesDoc = null; // (the count starts again when the document changes)
  // Splitting paragraphs line by line needs each page-crossing paragraph laid
  // out; in very long documents (content-visibility skips off-screen blocks)
  // that costs too much, so there paragraphs move to the next page whole.
  // (On screen. For print every block is laid out, editor.css, and the paragraphs split as in any document: a document
  // of 3,162 paragraphs had printed 128 pages for Word's 122, each page ending where a whole paragraph did.)
  let printing = false;
  const splitsNow = () => (!sheet.classList.contains('odx-large') || printing) && !cols;
  let lineSplits = splitsNow();
  let lastPass = 0; // when the last layout pass ran (performance.now)

  const layer = document.createElement('div');
  layer.className = 'odx-sheets';
  sheet.prepend(layer);
  const rules = document.createElement('style');
  document.head.appendChild(rules);

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

  // A floating picture in the paragraph (not one placed on the page) reaching past its page's foot. Word keeps it on its
  // paragraph's page, into the bottom margin as far as the page's edge, and moves it up from there so it ends at the
  // edge (its PDFs: floats 80 to 300px tall 10px below a paragraph near the foot, none moved to the next page, those
  // past the edge drawn up to it). One taller than the page's text area goes to the next page with its paragraph, as does
  // a text box grown to fit its text past the foot (one grown to fit a long table: the page before holds the title only). Returns { past } for that, or { up },
  // how far to draw the floats up (px), each by its index in the paragraph.
  const floatsPast = (b, shift, p) => {
    const er = rectOf(b.el);
    const up = new Map();
    [...b.el.querySelectorAll('img')].forEach((img, n) => {
      // (Not a picture wrapped top and bottom drawn as a float, reader.js --odx-shapeband: it isn't drawn up as one.)
      if (getComputedStyle(img).float === 'none' || img.style.getPropertyValue('--odx-shapeband')) return;
      const r = rectOf(img);
      // (Where it is laid out: less the drawing up a pass before gave it.)
      const ty = parseFloat((img.style.translate || '').split(/\s+/)[1]) || 0;
      const bottom = b.top + shift + (r.bottom - ty - er.top);
      if (bottom <= contentBottom(p) + 1) return;
      // (A text box grown to fit its text, past the foot at all: Word moves it on, where it keeps a shape of the
      // size it's saved at; sample p-textbox__landscape of the generated documents, 20px past the foot.)
      if (r.height > pageHeight(p) + 1 || placeVar(img.getAttribute('style'), 'odx-grown') != null) { up.set(-1, 0); return; }
      const edge = pageTop(p) + pageH(p);
      // (The edge kept for the lines it wraps above it: wrapboth.js drawnUpTo.)
      if (bottom > edge + 0.5) { up.set(n, bottom - edge); drawnUpTo.set(img, edge); }
    });
    return up.has(-1) ? { past: true } : { up };
  };
  // A picture wrapped top and bottom, placed down from its paragraph (reader.js anchorCss: a block at its offset), reaching
  // past its page's foot: Word keeps it on its paragraph's page, moved up to end at the bottom margin (in a document for
  // Word 2013 on, compatibility mode 15) or at the page's edge (modes 14 and below), and the text after it goes on above
  // it, the page ending at its top; what doesn't fit there goes on to the next page. Not moved above its paragraph's top:
  // there, the paragraph goes on to the next page with it, as before (coverpic-word: pictures 200 to 405pt high, 300 to
  // 600pt below a paragraph 196pt down, each drawn up to the foot with the eight lines after it above; anchors lower down
  // the page, moved on; sample 1387's cover picture, 315pt below its paragraph, at the foot of page 1 in Word's print).
  // Only one 15px or more below its paragraph's top, under the paragraph's own line there, as Word has it (one nearer
  // takes the line under it, bandFloats: that line doesn't fit either, and it goes on). One wrapped square as wide as the
  // column too, a band (anchorCss): coverpic-word a9, as a1. Returns { img, up,
  // top } (how far to draw it up, and its top on the page, drawn up), or null.
  const TB_LINE = 15; // px (as reader.js bandFloats)
  const tbPast = (b, shift, p) => {
    const imgs = b.el.querySelectorAll(':scope > img.odx-wrap-block');
    if (imgs.length !== 1) return null;
    const img = imgs[0];
    const er = rectOf(b.el);
    const r = rectOf(img);
    // (Where it is laid out: less the drawing up a pass before gave it, tbOwnMt. In the flow or out of it, it's at its offset.)
    const rel = r.top + (tbUpNow.get(img) || 0) - er.top;
    if (rel < TB_LINE) return null;
    const paraTop = b.top + shift;
    const top = paraTop + rel;
    const limit = pageTop(p) + pageH(p) - (session.compatMode >= 15 ? marginsAt(p).bottom ?? page.bottom : 0);
    const over = top + r.height - limit;
    if (over <= 0.5 || top - over < paraTop - 0.5) return null;
    return { img, up: over, top: top - over };
  };
  // A shape with text on both sides (bothShape) in the paragraph reaching past its page's bottom margin: Word keeps it on
  // the page with the paragraph's first lines, drawn up to end at the margin (not the page's edge, as a float), and
  // splits the paragraph as any other, its lines beside the shape where it's drawn, those on the next page the column's
  // width (wrap-word4 foot-h*: rectangles 30 to 80pt high at their paragraph's top, stepped down to the foot). Drawn up
  // over the text of the lines above, the paragraph goes on to the next page whole instead (`past`), its shape where
  // its offset puts it there (foot-h50-long Case 14: a filler line above reaching under it; Word saves where it drew a
  // shape, so its file has that rectangle 596.75pt below its paragraph, at page 2's foot, as wrap-word3 vt-thin has
  // Text 20's). And one above the page's top margin is drawn down to start there (sq-thick: Text 6's rectangle, saved
  // 7.2pt above its paragraph, at page 2's top margin). How far to draw each up (px, down when below 0), by its index
  // among the paragraph's pictures.
  const PAST_MIN = 1; // px a shape is drawn up past where it's saved before the text above it can move its paragraph on
  const MAX_ABOVE = 50; // blocks above looked at for text a shape drawn up would reach (a bound, not a fitted value)
  const MAX_SCANS = 32; // shapes of one paragraph that look (a bound, not a fitted value)
  const shapesPast = (b, shift, p) => {
    const er = rectOf(b.el);
    const up = new Map();
    let past = false;
    // (At most MAX_SCANS shapes of a paragraph look for text above them: a crafted paragraph of thousands of such shapes,
    // none over the text, scanned the same blocks for each on every pass.)
    let scans = 0;
    [...b.el.querySelectorAll('img')].forEach((img, n) => {
      if (!bothShape(img)) return;
      const r = rectOf(img);
      if (r.height >= contentBottom(p) - contentTop(p)) return; // (taller than the page's text area: as it is)
      const ty = parseFloat((img.style.translate || '').split(/\s+/)[1]) || 0;
      const rel = r.top - ty - er.top; // (its top from its paragraph's, as laid out)
      const over = b.top + shift + rel + r.height - contentBottom(p);
      const under = contentTop(p) - (b.top + shift + rel);
      // (Its room from the margin it's drawn to, as a float's from its edge: wrapboth.js drawnUpTo.)
      if (under > 0.5) { up.set(n, -under); drawnUpTo.set(img, contentTop(p) + r.height); return; }
      if (over <= 0.5) return;
      up.set(n, over);
      drawnUpTo.set(img, contentBottom(p));
      // (Only when drawn up further than the half point or so the page's lines may sit lower than Word's: one Word drew up
      // is saved there, wrap-word3 sq-none's Text 6 7.2pt above its paragraph, split over the pages as in Word.)
      if (over <= PAST_MIN || past || ++scans > MAX_SCANS) return;
      // (The text above it there: of the blocks above on the page, a line reaching below its top, across its room.)
      const v = getComputedStyle(img).getPropertyValue('--odx-both').trim().split(/\s+/).map((x) => parseFloat(x) || 0);
      const left = r.left - (v[3] || 0), right = r.right + (v[1] || 0);
      // (From where the paragraph is less the push the last pass gave it: moved on to the next page, the text above was
      // that page's gap away, and it came back, and went on again.)
      const own = er.top - (applied.get(b.i)?.extra || 0);
      const top = own + rel - over;
      const range = document.createRange();
      let sib = b.el.previousElementSibling;
      for (let k = 0; sib && k < MAX_ABOVE && !past; k++, sib = sib.previousElementSibling) {
        const sr = rectOf(sib);
        if (sr.bottom <= top) break;
        const walker = document.createTreeWalker(sib, NodeFilter.SHOW_TEXT);
        for (let t = walker.nextNode(); t && !past; t = walker.nextNode()) {
          if (!t.data.trim()) continue;
          range.selectNodeContents(t);
          for (const q of rectsOf(range)) if (q.width > 0 && q.bottom > top + 0.25 && q.top < own && q.right > left && q.left < right) { past = true; break; }
        }
      }
    });
    return { up, past };
  };
  const HF_GAP = 60, HF_FAR = 0.3; // a header or footer too tall for the page: how near the middle Word cuts it; from how near the far edge
  /**
   * Rows that have to start a new page, Word-style: a row that doesn't fit moves
   * to the next page, below a copy of the header rows. Returns null when the
   * first row doesn't fit either (the whole table moves instead).
   */
  // A table's footnotes by the row of its own that holds the reference: row index -> [{ r, dy (the reference's middle
  // below the row's top, our spacers in its cell taken out), cellPos (where its cell starts) }].
  const notesOfRows = (b) => {
    const out = new Map();
    const body = b.table.tBodies[0];
    if (!body) return out;
    for (const r of blockNotes.get(b.i) || []) {
      if (r.kind !== 'footnote') continue;
      const el = view.nodeDOM(r.pos);
      if (!el || el.nodeType !== 1) continue;
      let td = el.closest('td');
      while (td && td.parentElement?.parentElement !== body) td = td.parentElement.closest('td');
      if (!td) continue;
      const j = [...body.rows].indexOf(td.parentElement);
      if (j < 0) continue;
      const q = rectOf(el);
      const y = (q.top + q.bottom) / 2;
      let cellPos;
      try { cellPos = view.posAtDOM(td, 0); } catch { cellPos = r.pos; }
      if (!out.has(j)) out.set(j, []);
      out.get(j).push({ r, dy: y - rectOf(td.parentElement).top - undoIn(td, y), cellPos });
    }
    return out;
  };
  function splitTable(b, shift, tableFirstOnPage, chainAtTop = false) {
    const { rows } = b;
    let { headerRows } = b;
    // Before Word 2013's layout (compatibility mode 14 and earlier) header rows that don't fit at a page's foot break
    // there as other rows do, and none repeats on the next page; a table of only header rows doesn't move on either
    // (hdr314-word, hdrall14-word and the same in mode 12: the table stays, its rows going on line by line. A sample's
    // mode-14 table of six header rows had moved to the next page.)
    const oldLayout = !(session.compatMode >= 15);
    if (headerRows && oldLayout && !tableFirstOnPage && rows[headerRows - 1].bottom + shift + (rows[headerRows - 1].line || 0) > contentBottom(pageOf(rows[0].top + shift))) {
      headerRows = 0;
      b.headerRows = 0;
    }
    // (Repeated, the header rows keep their own line under them in that layout, a wider one than the table's other lines
    // taking that much more room above the next row; from mode 15 the line is the next row's own, as on the first page.
    // Word's prints of ..\hdrline-word, 32 generated tables, header lines of 4 to 24 eighths of a point, the cell's or a
    // table style's, the next row split or whole; sample 3314's header-row pages 2px high.)
    // (Less the line a row draws over itself at a page's top, as the table's first row does.)
    const hdrLine = headerRows && oldLayout ? Math.max(0, (rows[headerRows - 1].line || 0) - (rows[0].topLine || 0)) : 0;
    const headH = headerRows ? rows[headerRows - 1].bottom - rows[0].top + hdrLine : 0;
    const pushes = new Map();
    const offs = new Map();
    const cellPushes = [];
    const cellTops = []; // cells of rows split here that are top-aligned for it (splitRow)
    const bands = [];
    const closes = []; // where a row broken across pages gets its closing border
    const inRow = new Set();
    const heads = [];
    let extra = 0;
    const extraAt = [];
    // (A floating table's rows don't keep with the next: Word breaks it at the page's foot as one without keeps, set down
    // from its text or not; generated documents, every paragraph in its 11 rows keeping with the next.)
    const keeps = floating(b.node) ? () => false : rowKeeps;
    // Every row keeping with the next and the table longer than a page: the keeps can't all hold, and its rows split at
    // a page's foot as they would without them.
    let keepsHold = true;
    if (rows.length && rows[rows.length - 1].bottom - rows[0].top > pageHeight(pageOf(rows[0].top + shift) + 1)) {
      keepsHold = false;
      for (let j = 0; j < b.node.childCount && !keepsHold; j++) if (!keeps(b.node.child(j))) keepsHold = true;
    }
    // Its footnotes go on the page of the row that holds their reference, not all on the table's first page, as in Word
    // (tblnotes-word: a reference in each of 40 rows, each note at the foot of its row's page; a sample's staffing tables
    // had every note of a three-page table on its first, the rows pushed on by them).
    // A table whose rows are all header rows doesn't start at a page's foot: Word 2013 and later move it to the next page
    // when it doesn't fit, also when it's longer than a page, and breaks it there, repeating none (hdrall-word: 5, 30 and 60 rows
    // under 44, 30 and 10 lines of text; a sample's travel table, each of its 20 rows a header row, had started half way
    // down a page).
    if (b.allHeader && !oldLayout && !tableFirstOnPage && !chainAtTop) return null;
    const rowNotes = notesOfRows(b);
    const home = new Map(); // note -> the band it was in, to put back when the table isn't split after all
    for (const list of rowNotes.values()) for (const x of list) { home.set(x.r, placed.get(x.r.kind + x.r.id)); takeNote(x.r); }
    if (home.size) sumNotes();
    const placeNote = (r, pg) => {
      if (placed.get(r.kind + r.id) === bandOf(pg)) return false;
      takeNote(r);
      putNote(r, bandOf(pg));
      return true;
    };
    const placeRow = (j, pg) => {
      let moved = false;
      for (const x of rowNotes.get(j) || []) if (placeNote(x.r, pg)) moved = true;
      if (moved) sumNotes();
    };
    const unplaceFrom = (j) => {
      let moved = false;
      for (const [q, list] of rowNotes) if (q >= j) for (const x of list) if (takeNote(x.r)) moved = true;
      if (moved) sumNotes();
    };
    // (A row with references fits when it does with their notes on its page; else they are taken out again.)
    const fitsWithNotes = (j, pg, bottom) => {
      if (!rowNotes.has(j)) return bottom <= contentBottom(pg);
      placeRow(j, pg);
      if (bottom <= contentBottom(pg)) return true;
      unplaceFrom(j);
      return false;
    };
    const restore = () => {
      if (!home.size) return null;
      for (const [r, p] of home) { takeNote(r); if (p != null) putNote(r, p); }
      sumNotes();
      return null;
    };
    for (let j = 0; j < rows.length; j++) {
      extraAt[j] = extra;
      if (rows[j].hidden) continue; // (no box: measureBlocks)
      const top = rows[j].top + shift + extra;
      const bottom = rows[j].bottom + shift + extra;
      const pg = pageOf(top);
      const brk = j > headerRows && j > 0 && b.node.child(j) && rowBreakAt(b.node, j) && top > contentTop(pg) + headH + 1;
      if (!brk && fitsWithNotes(j, pg, bottom + (rows[j].line || 0))) continue;
      // Word breaks a row inside its cells, line by line, unless the row says it can't split, or (sample S072) the room left
      // on the page is less than the row's minimum height (w:trHeight "at least"): then the row moves whole. The room has
      // to hold its top line and its own bottom line as well (rowmin-word: Word's prints of 36 tables, a 32.25pt row of four
      // lines at a page's foot, its cells with top and bottom lines of their own: with none it splits from 32.25pt of room;
      // with 0.5pt lines, from 33.25pt, moving whole at 33pt; with 1pt lines, from 34.25pt, whole at 33.75pt; with 3pt
      // lines, whole up to 35.25pt. The same 1pt row with no bottom line of its own, the next row's top line under it,
      // splits at 33.75pt: table-row-foot. A sample's 645tw rows under 1pt top lines and 0.5pt bottom ones, 44.2px from
      // the foot, split there and every page after it was a row out.)
      const tr = b.node.child(j)?.attrs.trPr || '';
      const minM = /<w:trHeight\b[^>]*w:val="(\d+(?:\.\d+)?)"/.exec(tr);
      const minH = minM && !/<w:trHeight\b[^>]*w:hRule="(?:exact|auto)"/.test(tr) ? twipsToPx(minM[1]) : 0;
      // (Less a little for the lines' widths, kept to 1/64px: a row with room for exactly its height and its lines splits,
      // rowmin-04-4. With no lines the room is compared as it is: a sample's row with room for exactly its 100px moved on
      // as in Word, the rows above it 1.3px higher than Word's.)
      // With no line under it, its own or the next row's, the row's part there takes the table's bottom line, and that has
      // to fit too (why-tests/fix2711 b-*: a 45pt row under a 0.25pt bottom line moves on with 0.2px over its height,
      // splits with 0.5px; under a 3pt line, moves with 3.5px, splits with 4.5px; modes 14 and 15. A sample's row with
      // 0.13px less than its 60px plus that line split where Word moves it, the pages after it a row out.)
      const lines = (rows[j].topLine || 0) + (rows[j].ownLine || (rows[j].line ? 0 : rows[j].tableLine || 0));
      const short = minH && contentBottom(pg) - top - lines < minH - (lines ? 0.1 : 0) && minH <= pageHeight(pg + 1);
      // A row taller than a page breaks across pages in Word all the same, also one that can't split: at the next page's
      // top when it can't start here (tallrow-word: rows of 60 lines, under 20 lines of text and at a page's foot with
      // room for one line or two, cantSplit or not; a sample's 2,000px cantSplit row ran past the page's foot and over
      // the next one's top margin).
      const tallRow = rows[j].bottom - rows[j].top > pageHeight(pg + 1);
      const first = j === 0 ? tableFirstOnPage : top <= contentTop(pg) + (pushes.size || inRow.size ? headH : 0) + 1;
      const force = tallRow && first && !brk && j >= headerRows && lineSplits;
      const canSplit = force || (!brk && j >= headerRows && lineSplits && !short && !(keepsHold && j > headerRows && j < rows.length - 1 && keeps(b.node.child(j))));
      let rs = null;
      let tried = false;
      // Keep with next: the rows above that keep with this one (rowKeeps) go to the next page with it, as in Word; when
      // they reach the table's first row, the table moves (sample: Heading 1 in its cells). Unless this row can split
      // here, some of its lines left on the page: that is all the keep asks for, and Word leaves the run where it is
      // (keeprow-word: a keep-with-next heading row over a row of ten lines stays at the page's foot with two or more of
      // its lines, as without the keep; with room for one, widow control taking it on, the heading row goes with it; a
      // row that can't split takes it on always. A sample's table heading row had gone on with its first row, a page of
      // text 186px low after it).
      if (!brk) {
        let s = j;
        while (s > headerRows && keeps(b.node.child(s - 1)) && !pushes.has(s - 1) && !inRow.has(s - 1) && pageOf(rows[s - 1].top + shift + extraAt[s - 1]) === pg) s--;
        // (Stopped by a row above that keeps with this run but was itself moved here or carried on from the page before:
        // the run already starts the page, and moving it again gains nothing; Word breaks it where it runs out, as a
        // sample's 14 such rows over four pages, each page had been left half empty.)
        const fromTop = s > headerRows && keeps(b.node.child(s - 1)) && (pushes.has(s - 1) || inRow.has(s - 1));
        if (s < j && canSplit) { rs = splitRow(b, j, shift + extra, headH, force); tried = true; }
        if (s < j && !rs) {
          const topS = rows[s].top + shift + extraAt[s];
          const firstS = fromTop || (s === 0 ? tableFirstOnPage || chainAtTop : topS <= contentTop(pg) + (pushes.size || inRow.size ? headH : 0) + 1);
          // (Only the run above this row has to fit the next page: a row taller than a page splits there, and Word takes the
          // heading row on with it all the same, kntall-word.)
          const tall = rows[j].top - rows[s].top + headH > pageHeight(pg + 1);
          if (!firstS && !tall && s <= headerRows && !tableFirstOnPage) return restore();
          if (!firstS && !tall && s > headerRows) {
            const d = contentTop(pg + 1) + headH - topS;
            pushes.set(s, d);
            offs.set(s, extraAt[s]);
            if (!headerRows) closes.push(contentTop(pg + 1) + 1);
            heads.push(contentTop(pg + 1));
            extra = extraAt[s] + d;
            unplaceFrom(s);
            placeRow(s, pg + 1);
            j = s;
            continue;
          }
        }
      }
      if (!tried && canSplit) rs = splitRow(b, j, shift + extra, headH, force);
      if (rs) {
        cellPushes.push(...rs.pushes);
        cellTops.push(...(rs.tops || []));
        for (const [q, cut] of rs.crossings) {
          bands.push([cut - (b.tableTop + shift), contentTop(q + 1) + headH - (b.tableTop + shift)]);
          heads.push(contentTop(q + 1));
          closes.push(cut);
          if (!headerRows) closes.push(contentTop(q + 1) + 1); // and its top border where it carries on
        }
        inRow.add(j);
        // (A split row's notes by the line holding the reference: the page its part with that line is on.)
        for (const x of rowNotes.get(j) || []) {
          const down = rs.pushes.filter((q) => q.pos >= x.cellPos && q.pos <= x.r.pos).reduce((t, q) => t + q.delta, 0);
          placeNote(x.r, pageOf(rows[j].top + shift + extra + x.dy + down));
        }
        if (rowNotes.has(j)) sumNotes();
        extra += rs.growth;
        continue;
      }
      if (first) { placeRow(j, pg); continue; } // (nothing gained moving it on)
      if (j <= headerRows) return restore();
      // (Less the room over its text that isn't the row's own, a wider line from the row above, --odx-ptc, styles.js: at
      // a page's top the row's text starts under its own top line only. Word's prints of six generated tables of rows
      // that can't split, moved whole: under a top line of their own the text starts that much down the page, with
      // only bottom lines, at the page's top.)
      const moved = b.table.tBodies[0]?.rows[j];
      const lift = moved ? Math.min(...[...moved.cells].filter((td) => !td.hasAttribute('data-odx-phantom')).map((td) => {
        const ptc = td.style.getPropertyValue('--odx-ptc');
        return ptc ? Math.max(0, (parseFloat(td.style.getPropertyValue('--odx-pt')) || 0) - (parseFloat(ptc) || 0)) : 0;
      })) : 0;
      const d = toGrid(contentTop(pg + 1) + headH - top - (Number.isFinite(lift) ? lift : 0));
      pushes.set(j, d);
      offs.set(j, extra); // everything above this row has moved down by this much (pushed and split rows)
      if (!headerRows) closes.push(contentTop(pg + 1) + 1); // the moved row's top border, where it now starts: at the page's top, as Word draws it (it was a pixel above)
      heads.push(contentTop(pg + 1));
      placeRow(j, pg + 1);
      extra += d;
      if (tallRow) j--; // (to break it at its new place: above)
    }
    return pushes.size || inRow.size ? { rows: pushes, offs, heads, total: extra, cellPushes, cellTops, bands, inRow, closes } : null;
  }

  // (Not in a document of mode 14 or earlier: there Word leaves one line of a cell's paragraph at a page's foot, or
  // carries one over, widow control or not; in mode 15 it doesn't. Its prints of ..\cellwidow-word, 24 generated
  // tables of a three-line paragraph in a cell, room for one, two or three lines; sample 3314's three-line cell row.)
  const cellWidow = (p) => {
    if (!(session.compatMode >= 15)) return false;
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
  // (A floating table's keeps don't hold: Word breaks its rows as without them; generated documents.)
  const floating = (table) => /<w:tblpPr\b/.test(table.attrs.tblPr || '');
  // (Nor a row of exact height: it moves whole, its text cut to the row or not. Word's prints of 24 generated tables at
  // a page's foot, rows of three lines: automatic rows break line by line, exact ones never, also 0.5in ones that cut
  // their text. Such a row can't grow to span the page's foot either: broken, the rows after it stayed where they
  // were and ran off the page, a sample's 21-row table printed over the footer and every page after it out of place.)
  const exactRow = (row) => /<w:trHeight\b[^>]*w:hRule="exact"/.test(row.attrs.trPr || '');
  // (Nor, in mode 14 and earlier, a header row: it moves whole, between rows the table still breaks. Word's prints of
  // ..\cellwidow-word hdr2-m14, hdrall-m14: a header row of three lines with room for one or two goes to the next page;
  // sample 2133's table of header rows.)
  const headerRow = (row) => !(session.compatMode >= 15) && /<w:tblHeader(?![^>]*w:val="(?:0|false|off)")/.test(row.attrs.trPr || '');
  const cantSplit = (row, table) => /<w:cantSplit(?![^>]*w:val="(?:0|false|off)")/.test(row.attrs.trPr || '') || exactRow(row) || headerRow(row) || (!(table && floating(table)) && keepsWhole(row));

  // Our spacers in a cell (not in the tables in it), and how much they have made a row of a table in a cell taller:
  // what's below them is measured less that, at its natural place.
  // (A cell's blocks that have a box. A hidden one, tracked or vanished text, reads as at the window's top: scrolled down
  // the document, that is far below the row, and the row measured thousands of px tall, by how far the page was scrolled.)
  const shownKids = (el) => [...el.children].filter((e) => e.getClientRects().length);
  const lastShown = (el) => {
    for (let e = el.lastElementChild; e; e = e.previousElementSibling) if (e.getClientRects().length) return e;
    return null;
  };
  const spacersIn = (td) => [...td.querySelectorAll('.odx-push')].filter((e) => e.closest('td') === td).map((e) => rectOf(e));
  const grown = (ntr) => {
    if (!ntr.querySelector('.odx-push')) return 0;
    let nat = -Infinity;
    for (const ntd of ntr.cells) {
      const last = lastShown(ntd);
      if (!last) continue;
      const lb = rectOf(last).bottom;
      nat = Math.max(nat, lb - undoIn(ntd, lb) + (parseFloat(getComputedStyle(ntd).paddingBottom) || 0));
    }
    return Number.isFinite(nat) ? Math.max(0, rectOf(ntr).bottom - nat) : 0;
  };
  const tablesIn = (td) => [...td.children].map(cellTable).filter(Boolean);
  function undoIn(td, y) {
    let u = spacersIn(td).reduce((sum, r) => sum + (r.bottom <= y + 0.5 ? r.height : 0), 0);
    for (const t of tablesIn(td)) for (const ntr of t.rows) {
      if (rectOf(ntr).bottom > y + 0.5) break;
      u += grown(ntr);
    }
    return u;
  }
  /**
   * Lines of each cell of row j that move to the next page(s), measured like
   * paragraphs. `rowShift` is how far the row has moved from its natural place.
   * Returns null when the row should move whole (nothing of it fits, or it can't split).
   */
  function splitRow(b, j, rowShift, headH, force = false) {
    if (!force && cantSplit(b.node.child(j), b.node)) return null;
    const tr = b.table.tBodies[0].rows[j];
    const base = rectOf(view.dom).top + b.rows[j].undo;
    const pushes = [];
    const crossings = new Map(); // page -> bottom of the row's part that stays on it
    const asTop = new Set(); // the row's own centred or bottom-aligned cells, laid out from its top as it splits
    let keeps = false;
    let newBottom = -Infinity;
    let endTd = null; // the cell that ends lowest
    // The lines of a run of paragraphs in one cell from `extra0` on: which move to the next page, as a paragraph's.
    function runLines(lines, extra0, padTop, padBottom, line, seq0, out, after = 0) {
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
        // (A paragraph carried whole to the next page keeps there what its space before has over the space after of the
        // paragraph left behind, inside the row as Word sets it: ..\splitspace-word, 195tw before after none 13px down,
        // after 80tw 7.7px, after 240tw none, 40tw after 80tw none; sample 1113 page 22, sample 1283's table text none.)
        const prevP = lines[n].p.previousElementSibling;
        const prevAfter = prevP && prevP.tagName === 'P' ? Math.max(0, parseFloat(getComputedStyle(prevP).marginBottom) || 0) : 0;
        const sb = whole && seq0 + n > 0 ? Math.max(0, (parseFloat(getComputedStyle(lines[n].p).marginTop) || 0) - prevAfter) : 0;
        const delta = toGrid(contentTop(pg + 1) + headH + padTop + lead + sb - top);
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
      // (The cell ends below its last paragraph's space after: margins don't pass through a cell's edge.)
      if (lines.length) out.end = Math.max(out.end, lines[lines.length - 1].bottom + rowShift + extra + after + padBottom);
      return extra;
    }
    // One cell from `extra0` on (`undoOut`: the natural offset the cells round it give a place in it). Its paragraphs line
    // by line; a table in it row by row, each of its cells as one of these, the row as tall as its tallest (Word breaks a
    // table in a cell across pages too). false: the row has to move whole.
    function plan(td, undoOut, extra0, padTop0, padBottom0, line, seq0, out) {
      // A cell centred or bottom-aligned inside a nested table would shift as lines move: such a row moves whole. (A cell
      // of the row itself is laid out from the row's top as it splits: asTop, below.)
      if (!asTop.has(td) && !/^(top|baseline)$/.test(getComputedStyle(td).verticalAlign)) return false;
      const cs = getComputedStyle(td);
      // (At the next page's top the cell's own room only, --odx-ptc where it is less than its padding: styles.js.)
      const ptc = td.style.getPropertyValue('--odx-ptc');
      const padTop = padTop0 + (ptc ? parseFloat(ptc) || 0 : parseFloat(td.style.getPropertyValue('--odx-pt')) || parseFloat(cs.paddingTop) || 0);
      const padBottom = padBottom0 + (parseFloat(cs.paddingBottom) || 0);
      const undoAt = (y) => undoOut(y) + undoIn(td, y);
      let extra = extra0;
      let seq = seq0;
      let lines = [];
      const flush = (after = 0) => {
        const e = runLines(lines, extra, padTop, padBottom, line, seq, out, after);
        if (e === false) return false;
        extra = e;
        seq += lines.length;
        lines = [];
        return true;
      };
      for (const el of shownKids(td)) {
        if (el.tagName === 'P') {
          const r = rectOf(el);
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
        const t = cellTable(el);
        if (!t || !t.tBodies[0] || !flush()) return false; // anything else in the cell: move the row whole
        for (const ntr of t.tBodies[0].rows) {
          if (/<w:cantSplit|<w:trHeight\b[^>]*w:hRule="exact"/.test(ntr.dataset?.trpr || '')) return false;
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
      const lastEl = lastShown(td);
      if (!flush(lastEl && lastEl.tagName === 'P' ? Math.max(0, parseFloat(getComputedStyle(lastEl).marginBottom) || 0) : 0)) return false;
      return extra;
    }
    // A cell centred or set at the bottom (as its document has it: a split row's cells are top-aligned by a class, below):
    // Word breaks the row all the same, line by line. Content that fits in the part of the row on its first page is
    // aligned in that part (centred between the row's top and the last line kept on the page, or ending with that line);
    // longer content runs from the row's top as a top-aligned cell's does (48 generated documents: a two-line label
    // cell, top, centred and at the bottom, beside 25 to 100 lines, from 5 and 30 lines down the page).
    const rowTop = b.rows[j].top + rowShift;
    const pg0 = pageOf(rowTop);
    const aligned = [];
    const tops = [];
    for (const td of tr.cells) {
      if (td.hasAttribute('data-odx-phantom')) continue; // an empty grid slot, nothing to move
      const line = parseFloat(td.style.getPropertyValue('--odx-bdb')) || 0; // the line under the row, the next row's: it has to fit too
      const out = { atStart: false, kept: false, end: -Infinity };
      const va = /^(middle|bottom)$/.exec(td.style.verticalAlign || '')?.[1];
      let off = 0; // how far the browser's own alignment has the cell's content from where a top-aligned cell's starts
      if (va) {
        const kids = shownKids(td).filter((e) => !e.classList.contains('odx-push'));
        if (!kids.length) continue;
        const cs = getComputedStyle(td);
        const padTop = parseFloat(td.style.getPropertyValue('--odx-pt')) || parseFloat(cs.paddingTop) || 0;
        const padBottom = parseFloat(cs.paddingBottom) || 0;
        const first = rectOf(kids[0]), last = rectOf(kids[kids.length - 1]);
        // (With its first paragraph's space before and its last's space after, which Word centres with the text: a sample's
        // one-line label beside six lines kept on the page, 2pt before and after, was 2.4px low without them.)
        const space = Math.max(0, parseFloat(getComputedStyle(kids[0]).marginTop) || 0) + Math.max(0, parseFloat(getComputedStyle(kids[kids.length - 1]).marginBottom) || 0);
        const content = last.bottom - first.top - undoIn(td, last.bottom) + space;
        const tdTop = rectOf(td).top;
        off = /^(top|baseline)$/.test(cs.verticalAlign) ? 0 : Math.max(0, first.top - tdTop - (parseFloat(cs.borderTopWidth) || 0) - (parseFloat(cs.paddingTop) || 0));
        tops.push(td);
        if (rowTop + padTop + content + padBottom + line <= contentBottom(pg0) + 1) { aligned.push({ td, va, content, padTop, padBottom, first: kids[0] }); continue; }
        asTop.add(td);
      }
      if (plan(td, () => off, 0, 0, 0, line, 0, out) === false) return null;
      // The row splits only when a cell that breaks keeps some of its lines here (as in Word);
      // if one would have to move whole, so does the row.
      if (out.atStart) return null;
      if (out.kept) keeps = true;
      if (!Number.isFinite(out.end)) continue;
      for (const [q, v] of crossings) if (pageOf(out.end - 1) === q) crossings.set(q, Math.max(v, out.end)); // cells that end on that page
      if (out.end > newBottom) endTd = td;
      newBottom = Math.max(newBottom, out.end);
    }
    if (!keeps || !pushes.length) return null;
    // The row's part on its last page is as tall as the row's minimum height (w:trHeight "at least") all the same: Word
    // gives each part the whole of it (why-tests/fix2711 r2, r4, r6, m2, m6: a 45pt row's last two lines on the next page,
    // the text after the table 23px lower than under those lines alone, modes 15 and 14). The room under the cell that
    // ends lowest, after its last paragraph. (A sample's row parts ended at their text, the page's rows from there high.)
    const trPr = b.node.child(j)?.attrs.trPr || '';
    const minM = /<w:trHeight\b[^>]*w:val="(\d+(?:\.\d+)?)"/.exec(trPr);
    const minH = minM && !/<w:trHeight\b[^>]*w:hRule="(?:exact|auto)"/.test(trPr) ? twipsToPx(minM[1]) : 0;
    // Also where only the cell's empty last paragraph goes on there: Word gives that part the whole minimum height too
    // (why-tests rows, emp-*: 51 lines and an empty paragraph that doesn't fit, rows of at least 300 to 1500 twips, the
    // empty paragraph 8, 11 and 16pt: the next row as far down as under the minimum, or under the paragraph where that
    // is taller; samples 1544, 0577, 0621 and 2363 gain with it). The room after the last paragraph that takes room (in
    // one drawn at no height, w:hideMark's, it was hidden, though the next pass still took it off). (Its line kept:
    // editor.css odx-push-end.)
    let lastP = endTd && lastShown(endTd);
    while (lastP && lastP.tagName === 'P' && !(rectOf(lastP).height > 0)) lastP = lastP.previousElementSibling;
    const q = pageOf(newBottom - 1);
    if (minH && Number.isFinite(newBottom) && lastP?.tagName === 'P') {
      const cs = getComputedStyle(endTd);
      const room = (parseFloat(endTd.style.getPropertyValue('--odx-pt')) || parseFloat(cs.paddingTop) || 0) + minH + (parseFloat(cs.paddingBottom) || 0);
      const need = toGrid(contentTop(q) + headH + room - newBottom);
      let end = null;
      try { end = view.state.doc.resolve(view.posAtDOM(lastP, 0)).end(); } catch { /* a paragraph the document no longer has */ }
      if (q > pg0 && need > 0.5 && end != null) {
        pushes.push({ pos: end, delta: need, end: true });
        newBottom += need;
      }
    }
    // The aligned cells' content in the row's part on its first page: a spacer above it, the cell top-aligned.
    const cut0 = crossings.get(pg0);
    for (const a of aligned) {
      if (!Number.isFinite(cut0) || a.first.tagName !== 'P') continue;
      const room = Math.min(cut0, contentBottom(pg0)) - a.padBottom - (rowTop + a.padTop) - a.content;
      const down = toGrid(a.va === 'middle' ? room / 2 : room);
      if (down > 0.5) pushes.push({ pos: view.posAtDOM(a.first, 0), delta: down });
    }
    return { pushes, tops: tops.map((td) => view.posAtDOM(td, 0) - 1), crossings: [...crossings].map(([q, v]) => [q, Number.isFinite(v) ? Math.min(v, contentBottom(q)) : b.rows[j].top + rowShift]), growth: Math.max(0, newBottom - (b.rows[j].bottom + rowShift)) };
  }

  const widowControl = (node) => node.attrs.view.widowControl ?? session.styleProp(node.attrs.styleId, 'widowControl') ?? session.styles.defaults.pPr.widowControl ?? true; // on when nothing says (Word's own default, though the OOXML spec says off)
  // (A paragraph run into through hidden marks keeps lines as the first of them does: joins.js runOnLead.)
  const keepLines = (node) => { const n = runOnLead(node) || node; return n.attrs.view.keepLines ?? session.styleProp(n.attrs.styleId, 'keepLines'); };

  /**
   * The paragraph's lines in natural (unpaginated) coordinates: line boxes
   * estimated from the text fragments, with our own spacers taken out.
   */
  function paragraphLines(b) {
    const natural = naturalOf(b);
    const frags = [];
    const walker = document.createTreeWalker(b.el, NodeFilter.SHOW_TEXT);
    const range = document.createRange();
    for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      if (t.parentElement.closest('.odx-push')) continue;
      range.selectNodeContents(t);
      for (const r of rectsOf(range)) if (r.height > 0) frags.push(r);
    }
    // (A turned picture by its box, not its turned outline: the room it takes on its line.)
    // (Not a shape with text on both sides: bothShape.)
    // (Nor a drawing out of the flow, in front of the text or behind it: it takes no room on a line. Sample 2466's cover
    // text box, 176px under the page's last paragraph, had been taken for a second line of it, and the paragraph moved
    // on whole wherever it ended a hair past the margin, as in Firefox.)
    b.el.querySelectorAll('img:not(.ProseMirror-separator), br:not(.ProseMirror-trailingBreak)').forEach((e) => {
      if (bothShape(e)) return;
      if (e.tagName === 'IMG' && !e.classList.contains('odx-wrap-block') && getComputedStyle(e).position === 'absolute') return;
      const r = layoutRect(e);
      if (r.height > 0) frags.push(r);
    });
    // A paragraph ending in a line break has an empty last line (drawn by the editor's own break): Word counts it
    // for widow control (three lines and the break split two and two).
    const trailing = frags.length ? b.el.querySelector('br.ProseMirror-trailingBreak') : null;
    let trailRect = null;
    if (trailing) { const r = rectOf(trailing); if (r.height > 0) { frags.push(r); trailRect = r; } }
    // (Drawn by the paragraph's ::after in place of that break, editor.css odx-endbreak: the line at its content's foot.)
    if (frags.length && !trailRect && (b.el.classList.contains('odx-endbreak') || b.el.classList.contains('odx-tabend'))) {
      const cs = getComputedStyle(b.el);
      const er = rectOf(b.el);
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
    const breaks = [...b.el.querySelectorAll('.odx-page-break')].map((e) => rectOf(e).bottom);
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
        if ((n.tagName === 'IMG' || n.tagName === 'BR') && !bothShape(n) && onLine(rectOf(n))) return view.posAtDOM(n.parentNode, [...n.parentNode.childNodes].indexOf(n));
        continue;
      }
      range.selectNodeContents(n);
      const rects = [...rectsOf(range)];
      if (!rects.some((r) => r.top >= top - 2 && r.top < top + 2) && !rects.some(onLine)) continue;
      // first character on that line (character tops only grow through a text node)
      let lo = 0;
      let hi = n.length - 1;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        range.setStart(n, mid);
        range.setEnd(n, mid + 1);
        const r = rectsOf(range);
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

  /** Whether a paragraph's page break comes before any of its text (it starts with the break). */
  function leadingBreak(el) {
    const br = el.querySelector('.odx-page-break');
    if (!br) return false;
    const range = document.createRange();
    range.setStart(el, 0);
    range.setEndBefore(br);
    return !range.toString().trim();
  }

  /** A table set down at a place on its page (styles.js): { y, rel, above, below } in px, or null. */
  function floatAt(b) {
    if (!b.rows) return null;
    const t = tableOf(b.el);
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
  let roundOver = 0; // the walk's: what the page so far takes more than exactly, at the block being fitted (roundOff)
  const roundMemo = new WeakMap(); // paragraph node -> { h, own }
  /**
   * What a paragraph's box, and the room between it and the block above it on its page, take in the browser more than
   * exactly (px, either sign): a line's height is laid out as the nearest 64th of a pixel, and so is each paragraph's
   * spacing (props.js spacePt), where Word adds twips and parts of them up. Its lines all of one height, or nothing
   * for the box (a taller run or a picture in a line: not known). The room above as the nearest twip to what is there.
   */
  function roundOff(b, above) {
    if (b.node.type !== schema.nodes.paragraph) return 0;
    const h = b.bottom - b.top;
    let own = 0;
    const memo = roundMemo.get(b.node);
    if (memo && memo.h === h) own = memo.own;
    else {
      const cs = getComputedStyle(b.el);
      // (Less the room a list label adds over the first line: what it takes more than exactly is counted by itself, below.)
      const room = parseFloat(b.el.style.getPropertyValue('--odx-lbl-own')) || 0;
      const box = h - room - (parseFloat(cs.paddingTop) || 0) - (parseFloat(cs.paddingBottom) || 0) - (parseFloat(cs.borderTopWidth) || 0) - (parseFloat(cs.borderBottomWidth) || 0);
      const near = parseFloat(lineHeightOf(cs));
      const n = near > 0 ? Math.round(box / near) : 0;
      if (n >= 1 && box > 0) {
        // A line's exact height: the font's own times the spacing's multiple, or the exact height set (in px or pt).
        const each = box / n;
        const set = /^([\d.]+)(px|pt)$/.exec(cs.getPropertyValue('--odx-lhv').trim());
        const font = exactEm(parseFloat(cs.getPropertyValue('--odx-lh')) || 0) * (parseFloat(cs.fontSize) || 0) * exactMult(parseFloat(cs.getPropertyValue('--odx-mult')) || 1);
        // (An exact height is whole twips; its style value is cut to four decimals, 18.6667px for 14pt, and over 46 lines
        // that lost 0.0015px, enough to send a last line ending exactly on the margin to the next page.)
        const setPx = set ? Number(set[1]) * (set[2] === 'pt' ? 4 / 3 : 1) : 0;
        const twips = Math.round(setPx * 15) / 15;
        // (The set height before the font's: 10pt Times New Roman's line, 15.332px, is within a rounding of an exact 11.5pt,
        // 15.333px, and counted as it a page of 55 such lines came a twip short, keeping a last line Word moves on.)
        const exact = [Math.abs(twips - setPx) < 0.0002 ? twips : setPx, font].find((v) => v > 0 && Math.abs(v - each) < 0.012);
        if (exact) own = n * (each - exact) + (room ? room - (parseFloat(b.el.style.getPropertyValue('--odx-lbl-up')) || room) : 0);
        else if (!room) own = mixedOwn(b.el, cs, box);
      }
      roundMemo.set(b.node, { h, own });
    }
    const gap = above ? b.top - above.bottom : 0;
    return own + (gap > 0.004 && gap < 200 ? gap - Math.round(gap * 15) / 15 : 0);
  }

  /**
   * roundOff's part for a single-spaced paragraph of lines in several sizes of its own font (no pictures, tabs, raised or
   * lowered text, or other fonts): the box less each line's exact height, the font's own line at the line's largest size.
   * Word adds those up unrounded; the browser lays each line out to a 64th of a pixel, and sample 2466's cover titles (a
   * 48pt break over 36pt text, 36pt over 20pt) came 0.08px taller over two paragraphs, its last line a hair past the
   * margin where Word keeps it. 0 when unsure: a line not accounted for leaves the box off by more than its rounding.
   */
  function mixedOwn(el, cs, box) {
    const em = exactEm(parseFloat(cs.getPropertyValue('--odx-lh')) || 0);
    const mult = parseFloat(cs.getPropertyValue('--odx-mult')) || 1;
    if (!(em > 0) || mult !== 1 || /(px|pt)$/.test(cs.getPropertyValue('--odx-lhv').trim())) return 0;
    if (el.querySelector('img:not(.ProseMirror-separator), svg, math, .odx-tab, .odx-opaque-inline:not(:empty), sup, sub, .odx-strut')) return 0;
    const frags = [];
    const range = document.createRange();
    const family = cs.fontFamily;
    for (const e of el.querySelectorAll('span.odx-t, br')) {
      const ecs = getComputedStyle(e);
      if (ecs.fontFamily !== family) return 0;
      if (e.classList.contains('odx-sp0')) continue;
      const fs = parseFloat(ecs.fontSize) || 0;
      // (A break's line counts where its break is, with no box of its own too: Firefox gives a line holding only a break
      // none, and the line was lost from the count.)
      if (e.tagName === 'BR') { const r = rectOf(e); if (r.height > 0 || r.top || r.left) frags.push({ top: r.top, bottom: Math.max(r.bottom, r.top + 3), fs }); continue; }
      range.selectNodeContents(e);
      for (const r of rectsOf(range)) if (r.height > 0) frags.push({ top: r.top, bottom: r.bottom, fs });
    }
    if (!frags.length || frags.length > 400) return 0;
    frags.sort((x, y) => x.top - y.top);
    const lines = [];
    for (const f of frags) {
      const last = lines[lines.length - 1];
      if (last && f.top < last.bottom - 2) { last.bottom = Math.max(last.bottom, f.bottom); last.fs = Math.max(last.fs, f.fs); } else lines.push({ ...f });
    }
    if (lines.length < 2 || new Set(lines.map((l) => l.fs)).size < 2) return 0; // (one size: the uniform rule's, or not known)
    // (Each size as Word has it, to the half point: Firefox's 20pt is 26.6563px. And up to 0.05px a line either way: its
    // 1/60px grid and the runs' own line factors leave that much; a line not accounted for is off by whole pixels. In
    // Firefox 2466's 48/36pt title came 0.094px over its four lines, and with 0.02 a line nothing was counted.)
    const own = box - lines.reduce((t, l) => t + em * (Math.round(l.fs * 1.5) / 1.5), 0);
    return Math.abs(own) <= 0.05 * lines.length ? own : 0;
  }

  /**
   * How far a paragraph's last line may hang below the page's text area: Word places the extra space of
   * "multiple" line spacing below the text and doesn't count it when fitting a line (1px for rounding).
   */
  function lineSlack(el) {
    const cs = getComputedStyle(el);
    const mult = parseFloat(cs.getPropertyValue('--odx-mult')) || 1;
    const lh = parseFloat(lineHeightOf(cs));
    // (From the font's own units where the line is that font's: what Word adds, unrounded.)
    const own = exactEm(parseFloat(cs.getPropertyValue('--odx-lh')) || 0) * (parseFloat(cs.fontSize) || 0);
    const m = exactMult(mult);
    if (mult > 1 && own > 0 && Math.abs(own * m - lh) < 0.02) return 1 + own * (m - 1);
    return 1 + (mult > 1 && lh ? lh * (1 - 1 / mult) : 0);
  }

  /**
   * Widow and orphan control where a paragraph of `n` lines is cut before line `at` (its lines from `start` on this page
   * or column): two lines kept for the next one, and none left alone here, as in Word. The line to cut before, or -1 when
   * the paragraph has to move whole (one line would stay behind at the top of the paragraph). (A cell's paragraphs keep
   * their own rule: splitRow's runLines.)
   */
  const widowCut = (at, n, firstOnPage, start = 0) => {
    if (n - at < 2 && at - 1 > start) at = n - 2; // keep two lines for the next page or column
    if (start === 0 && at - start < 2 && !firstOnPage) return -1; // a single line can't stay behind
    return at;
  };

  /**
   * A paragraph's first line's text offset in its box (`lead0`, px: the text's top less the content box's) and the part
   * of it a line carried to a new page or column keeps above its text (`lead`): under "multiple" spacing the browser's
   * half of the extra goes above the text, where Word's text sits at the line's top, and so (corpus 2's double-spaced
   * Cambria carried to page 13 8px high without it) under double spacing; under "at least" spacing too (corpus 2's
   * 18pt-at-least Calibri carried to page 11 3px high without it). `labelPad`: less a list label's padding above its
   * line, the bullet line's own (sample S071's bullet carried to page 8, 1.9px low); pages only (splitParagraph).
   */
  function lineLead(b, lines, labelPad) {
    const pcs = getComputedStyle(b.el);
    const lh1 = pcs.getPropertyValue('--odx-lhlen').trim() === '1';
    const mult = lh1 ? 1 : parseFloat(pcs.getPropertyValue('--odx-mult')) || 1;
    // (The room a label reaching above the text adds over the first line, --odx-lbl-own: pararules.js labelLine; an
    // empty item's is its label's padding.)
    const lblUp = labelPad && b.el.classList.contains('odx-li') ? Math.max(parseFloat(getComputedStyle(b.el, '::before').paddingTop) || 0, parseFloat(pcs.getPropertyValue('--odx-lbl-own')) || 0) : 0;
    const lead0 = lines[0].screenTop - (rectOf(b.el).top + parseFloat(pcs.paddingTop) + parseFloat(pcs.borderTopWidth)) - lblUp;
    const atLeast = lh1 && parseFloat(pcs.getPropertyValue('--odx-runmult')) !== 0;
    const lead = (mult > 1 || atLeast) && Math.abs(lead0) <= 20 ? lead0 : 0;
    return { lead0, lead };
  }

  /** Where a paragraph breaks at its column's bottom: { cut (px below its top), height, delta (for what follows) } or null. */
  function colSplit(b, shift, firstOnPage, p) {
    const lines = paragraphLines(b);
    if (!lines || lines.length < 2 || keepLines(b.node)) return null;
    const slack = lineSlack(b.el) - 1 + roundOver;
    const after = bandAfter(p, b.node);
    // (As splitParagraph fits a line at a page's foot: a line's bottom here is the next line's text top, its half of
    // the leading under "multiple" spacing below the line's box, taken back off; and a line carried on keeps that
    // half above its text, as the paragraph's first line does.)
    // (Its own limits, not splitParagraph's: the inset only up to 6px (75e5009), and no list label's padding taken off.)
    const { lead0, lead } = lineLead(b, lines, false);
    const inset = lead0 > 0 && lead0 <= 6 ? lead0 : 0;
    let at = lines.findIndex((l, i) => l.bottom - (i < lines.length - 1 ? inset : 0) + shift + (i === lines.length - 1 ? after : 0) > contentBottom(p) + slack);
    if (at <= 0) return null;
    if (widowControl(b.node)) {
      at = widowCut(at, lines.length, firstOnPage);
      if (at < 0) return null;
    }
    if (at <= 0) return null;
    const cut = lines[at].top - b.top;
    // A paragraph longer than the next column too goes on in the one after, and so on: cut again at each column's
    // foot, two lines kept on each side (columns of equal widths).
    const parts = [cut];
    let from = at;
    const cg = cols || colGeo(p); // (in a section of columns among others too)
    for (let c = 1; cg && !cg.unequal && c < MAX_CARRY; c++) {
      const room = contentBottom(p + c) - contentTop(p + c) + slack;
      const top = lines[from].top - lead;
      let j = lines.findIndex((l, i) => i > from && l.bottom - (i < lines.length - 1 ? inset : 0) - top + (i === lines.length - 1 ? bandAfter(p + c, b.node) : 0) > room);
      if (j < 0) break;
      if (widowControl(b.node)) {
        if (lines.length - j < 2) j = lines.length - 2;
        if (j - from < 2) j = from + 2;
      }
      if (j <= from || j >= lines.length) break;
      parts.push(lines[j].top - b.top);
      from = j;
    }
    // (A line carried to the next column sits as the paragraph's first line does: a text box taller than its line, a
    // Title's 28pt in Aptos Display, starts above it, as in Word: the first line in column 2 3.4px low without it.)
    // (Up to 6px, where a page's is 2: fitted on that Title in bba1165; splitParagraph's 2 on sample S062's headings.)
    const above = !lead && lead0 < 0 && lead0 >= -6 ? lead0 : 0;
    return { cut, parts: parts.length > 1 ? parts : undefined, lift: lead, height: b.bottom - b.top, delta: contentTop(p + parts.length) + lead + above - (lines[from].top + shift) };
  }
  // A split block's carried parts: [{ off (columns on), cut, stop (px from its top, null: its foot), h }].
  const carriedOf = (cs) => (cs.parts
    ? cs.parts.map((cut, i) => { const stop = cs.parts[i + 1] ?? null; return { off: i + 1, cut, stop, h: (stop ?? cs.height) - cut + (cs.lift || 0) + (cs.head || 0) }; })
    : [{ off: cs.to || 1, cut: cs.cut, stop: null, h: (cs.rewrap ? cs.rewrap.h : cs.height - cs.cut) + (cs.head || 0) }]);

  // Columns of their own widths: a paragraph going on in a column of another width is laid out again at that width, as
  // Word does, so its lines there aren't the first column's. The text from its cut is copied out (`cs.rewrap`: the copy,
  // its width and height), and what follows moves by how much taller or shorter that makes it.
  const colWidthAt = (g, v) => g.widths[(((v - (g.P ?? 0)) % g.num) + g.num) % g.num];
  function rewrapCarry(b, cs, g, p) {
    const grow = colWidthAt(g, p + 1) - colWidthAt(g, p);
    if (Math.abs(grow) < 0.5) return;
    // The first character at or below the cut, in the paragraph's own layout (as paragraphLines measures it).
    const natural = naturalOf(b);
    const y = b.top + cs.cut - 1;
    const range = document.createRange();
    const box = (t, i) => { range.setStart(t, i); range.setEnd(t, i + 1); return [...rectsOf(range)].find((r) => r.height > 0); };
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
    copy.style.removeProperty('--odx-push');
    copy.style.margin = '0';
    copy.style.textIndent = '0';
    copy.classList.add('odx-contlines'); // (no list number or bullet: editor.css)
    copy.querySelectorAll('.ProseMirror-trailingBreak, .odx-push').forEach((e) => e.remove());
    const width = widthIn(b, g, p + 1);
    const h = heightAt(copy, width);
    cs.rewrap = { el: copy, width, h };
    cs.delta += h - (cs.height - cs.cut) - (cs.lift || 0); // (a fresh copy: its first line has its own lead)
  }
  /** A block's width in flow page v's column: as laid out, less the narrowing it has now, plus that column's. */
  function widthIn(b, g, v) {
    const now = parseFloat(getComputedStyle(b.el).getPropertyValue('--odx-coln')) || 0;
    return rectOf(b.el).width + now - (g.width - colWidthAt(g, v));
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
    const box = rectOf(el);
    if (lines) {
      const range = document.createRange();
      const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      for (let t = walker.nextNode(); t; t = walker.nextNode()) {
        range.selectNodeContents(t);
        for (const r of rectsOf(range)) {
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
    copy.style.removeProperty('--odx-push');
    copy.style.margin = '0';
    const lines = [];
    const bottom = b.top + heightAt(copy, widthIn(b, g, p), lines);
    let at = lines.findIndex((l) => b.top + l.bottom + shift > contentBottom(p) + roundOver + lineSlack(b.el) - 1);
    if (lines.length < 2 || keepLines(b.node)) at = -1;
    if (at > 0 && widowControl(b.node)) {
      at = widowCut(at, lines.length, firstOnPage);
    }
    return { bottom, splits: at > 0 };
  }

  // A table in columns goes on in the next column row by row: cut above the first row that doesn't fit, its rows from
  // there drawn again at the next column's top (as a paragraph's lines are, colSplit), under its header rows repeated.
  function colSplitTable(b, shift, p) {
    const table = tableOf(b.el);
    const trs = table && table.tBodies[0] ? [...table.tBodies[0].rows] : [];
    if (!trs.length) return null;
    const top0 = rectOf(b.el).top;
    const room = contentBottom(p) + 1 - (b.top + shift); // how far down the block the column ends
    const at = trs.findIndex((tr) => rectOf(tr).bottom - top0 > room);
    if (at < 0) return null;
    // Inside the row, between its cells' lines, as Word breaks a row line by line (when the row can split and its cells
    // sit at the top): the lowest place above the column's foot that no line of any cell straddles.
    let cut = null;
    const node = b.node.child(at);
    const rowTop = rectOf(trs[at]).top - top0;
    if (lineSplits && node && !cantSplit(node, b.node) && ![...trs[at].cells].some((td) => !/^(top|baseline)$/.test(getComputedStyle(td).verticalAlign))) {
      const base = rectOf(view.dom).top + b.undo;
      const spans = [];
      for (const el of trs[at].querySelectorAll('p')) {
        const r = rectOf(el);
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
    const h = headerRowsOf(b.node);
    const hb = h && h < trs.length ? rectOf(trs[h]).top - top0 : 0;
    const head = hb && cut >= hb - 0.5 ? hb : 0; // (the header rows drawn again above what goes on, once the cut is past them)
    // A table longer than the next column too goes on in the one after, and so on: cut again at a row's top at each
    // column's foot (columns of equal widths), its header rows again above each part.
    const parts = [cut];
    const cg = cols || colGeo(p);
    for (let c = 1; cg && !cg.unequal && c < MAX_CARRY; c++) {
      const from = parts[parts.length - 1];
      const fit = contentBottom(p + c) - contentTop(p + c) - head + 1;
      const j = trs.findIndex((tr) => rectOf(tr).bottom - top0 - from > fit);
      if (j < 0) break;
      const next = rectOf(trs[j]).top - top0;
      if (next <= from + 0.5) break; // (a row taller than a column: it stays whole)
      parts.push(next);
    }
    const lastCut = parts[parts.length - 1];
    return { cut, head, parts: parts.length > 1 ? parts : undefined, height: b.bottom - b.top, delta: contentTop(p + parts.length) + head - (b.top + shift + lastCut) };
  }

  /**
   * The line a paragraph that keeps with the next block is cut before when that block moves on and the paragraph may
   * split: its last line goes with the block, its last two under widow control, where as many stay behind; else null.
   */
  function keepCut(b) {
    if (b.node.type !== schema.nodes.paragraph || keepLines(b.node) || hasPageBreak(b.node)) return null;
    const lines = paragraphLines(b);
    if (!lines || lines.some((l) => l.forced || l.blank)) return null;
    const take = widowControl(b.node) ? 2 : 1;
    return lines.length >= 2 * take ? lines.length - take : null;
  }

  /**
   * Lines that must start a new page: after a page break, or where the page
   * runs out (with Word's widow/orphan control and keep-lines-together).
   * Returns null when the paragraph should move to the next page whole.
   */
  function splitParagraph(b, shift, blockFirstOnPage, moved = false, keepAt = -1) {
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
    const slack = lineSlack(b.el) - 1 - tie(b.el) + roundOver;
    // A line carried to the next page under "multiple" spacing keeps the browser's half of the extra above its text, as
    // the paragraph's first line does (the print shift takes it back): Word's text sits at the line's top either way.
    const { lead0, lead } = lineLead(b, lines, true);
    // And under any spacing, a text box a little taller than its line starts above it: a line the page ran out on goes
    // over with its box, not its text, at the page's top, as a paragraph's first line does (sample S071's single-spaced
    // Calibri, page 7 1.3px low). (Not one after a page break: sample S062's headings, 1.3px high with it.)
    const above = !lead && lead0 < 0 && lead0 >= -2 ? lead0 : 0;
    // A line's bottom is the next line's text top, which sits below that line's box by the text's offset in it: taken
    // off when fitting (doc 04: the second line of 4.4 ends 959.3px down a 960px page, as Word fits it).
    // (Up to the half of "multiple" spacing's extra, as for lead: sample S058's double-spaced 13pt Times, its
    // 10px offset left on, moved page 1's last line to page 2.)
    const inset = lead0 > 0 && (lead0 <= 6 || (lead && lead0 <= 20)) ? lead0 : 0;
    // The block's own footnotes, each with its reference's screen top: a line before a reference needs no room for its
    // note (Word keeps the lines above a long note's reference on the page, the reference's line going on with the note).
    const ownNotes = (blockNotes.get(b.i) || []).filter((r) => r.kind === 'footnote').map((r) => {
      const e = view.nodeDOM(r.pos);
      // (Its middle: a superscript mark's top is above its line's.)
      const q = e && e.nodeType === 1 ? rectOf(e) : null;
      return { r, y: q ? (q.top + q.bottom) / 2 : null };
    }).filter((x) => x.y != null);
    const bottomFor = (pg, n) => {
      const base = contentBottom(pg);
      const end = lines[n + 1] ? lines[n + 1].screenTop : Infinity;
      const band = bandOf(pg);
      const later = ownNotes.filter((x) => x.y >= end - 1 && placed.get(x.r.kind + x.r.id) === band).map((x) => x.r);
      if (!later.length) return base;
      const list = notesOn.get(band) || [];
      notesOn.set(band, list.filter((x) => !later.includes(x)));
      touchNotes(band);
      sumNotes();
      const without = contentBottom(pg);
      notesOn.set(band, list);
      touchNotes(band);
      sumNotes();
      return without;
    };
    const partialFits = (pg, n, bottom) => {
      const band = bandOf(pg);
      const end = lines[n + 1] ? lines[n + 1].screenTop : Infinity;
      const mine = (x) => placed.get(x.r.kind + x.r.id) === band;
      const inLine = ownNotes.filter((x) => mine(x) && x.y >= lines[n].screenTop - 1 && x.y < end - 1).map((x) => x.r);
      if (!inLine.length) return false;
      const later = new Set(ownNotes.filter((x) => mine(x) && x.y >= end - 1).map((x) => x.r));
      const others = (notesOn.get(band) || []).filter((r) => !later.has(r) && !inLine.includes(r));
      const at = bottom + (n === lines.length - 1 ? spaceAfter(b.node) : 0) + 11.5;
      const under = pageTop(pg) + pageH(pg) - marginsAt(pg).bottom - at - NOTE_SEP;
      // (Its first two lines: a note's paragraph keeps two lines together at a page's foot, as body text does, longfn-word2
      // fn4long's page 1, where Word moves the reference line on rather than leave one line of its note.)
      const bounds = notes.boundsOf(inLine[0].note?.doc, inLine[0].wKey ?? '') || [];
      const first = bounds[1] ?? inLine[0].h;
      // (And what of a note goes on from the page before: it comes first in the band, longfn-word2 fnhuge page 2.)
      const flow = noteFlow.get(band);
      const own = notesOn.get(band) || [];
      const carried = flow ? Math.max(0, flow.items.filter((r) => !own.includes(r)).reduce((t, r) => t + r.h, 0) - (flow.skip || 0)) : 0;
      if (under < carried + others.reduce((t, r) => t + r.h, 0) + first) return false;
      for (const r of inLine) { r.partial = true; r.at = at; }
      touchNotes(band);
      sumNotes();
      return true;
    };
    for (let n = 0; n < lines.length; n++) {
      const top = lines[n].top + shift + extra;
      const bottom = lines[n].bottom + shift + extra - (n < lines.length - 1 ? inset : 0);
      const pg = pageOf(top);
      // (Its last line's space after above a band of notes: bandAfter.)
      const fitBottom = bottom + (n === lines.length - 1 ? bandAfter(pg, b.node) : 0) + (footBand(pg) ? Math.max(0, slack) : 0);
      // (Only asked for a line that wouldn't fit with all of them.)
      const room = (k) => (fitBottom > contentBottom(pg) + slack && ownNotes.length ? bottomFor(pg, k) : contentBottom(pg));
      let at = -1;
      if (lines[n].forced) {
        // A paragraph that starts a page with a page break: the break already happened (Word doesn't add a blank page).
        if (n === 1 && moved && pushes.length === 0 && leadingBreak(b.el)) continue;
        at = n;
      }
      // (Its last lines taken on to stay with the block it keeps with, which has moved: keepCut.)
      else if (n === keepAt && n > pageStart) at = n;
      else if (fitBottom > room(n) + slack && n > pageStart && !(ownNotes.length && partialFits(pg, n, bottom))) {
        if (keepLines(b.node) && pushes.length === 0) return blockFirstOnPage ? { pushes, total: 0 } : null;
        at = n;
        if (widow) {
          at = widowCut(at, lines.length, blockFirstOnPage, pageStart);
          if (at < 0) return null;
        }
      } else if (fitBottom > room(n) + slack && n === 0 && !blockFirstOnPage && !lines[1]?.forced) {
        // A paragraph holding nothing but a page break stays at its page's foot when its line fits there without the
        // paragraph's own space before, which Word doesn't count for it: what follows starts the next page (breakAfter),
        // where taking the paragraph over left a page blank. Word's PDFs of 124 generated documents: 45 lines of a
        // 46-line page, then the break's paragraph with 0 to 35pt before, two pages each; with that much after the
        // line above it instead, or 46 lines, three.
        const own = lines.length === 1 && !b.node.textContent && hasPageBreak(b.node) ? Math.min(spaceBefore(b.node), parseFloat(getComputedStyle(b.el).marginTop) || 0) : 0;
        if (own > 0 && fitBottom - own <= room(n) + slack) continue;
        return null;
      }
      // (A first line that only holds a page break stays at the bottom of its page, as in Word: the text after it starts the next.)
      if (at < 0) continue;
      const lineTop = lines[at].top + shift + extra;
      // (Opening with the break, the paragraph starts the page there: its space before as at a page break before, breakSpace.)
      const opens = lines[at].forced && at === 1 && leadingBreak(b.el);
      const delta = contentTop(pageOf(lines[at].forced ? lineTop : top) + 1) + lead + (lines[at].forced ? (opens ? breakSpace(b.node, b.i > 0 ? view.state.doc.maybeChild(b.i - 1) : null) : 0) : above) - lineTop;
      // A closing break's empty line that doesn't fit: no spacer before it (after a block, Edge gives that line no box).
      // It stays, hidden in the margin, and the next block starts the page that much lower, as if the line went first, as
      // in Word (sample S041: page 3 opens with it).
      if (lines[at].blank && !lines[at].forced) { tail = lines[at].bottom - lines[at].top; continue; }
      const pos = lineStart(b.el, lines[at].screenTop);
      if (pos == null || delta <= 0) continue;
      pushes.push({ pos, delta: toGrid(delta) });
      extra += toGrid(delta);
      pageStart = at;
      n = at;
    }
    // A page break with nothing after it on its line: the next block starts the new page.
    const breakAfter = hasPageBreak(b.node) && !lines.some((l) => l.forced);
    return { pushes, total: extra, breakAfter, tail };
  }

  let headKey = '';
  // A paragraph's part carried to the next column, as drawn there: block element -> { copy, cut, rewrap } (for its
  // line numbers, which go beside the lines where they're seen).
  let carriedParts = new WeakMap();
  let blockCols = new WeakMap(); // block element -> its column ({ x from the text's left, w }), in columns
  const nodeIds = new WeakMap(); // block node -> a number (a block whose content changed is a new node)
  let idSeq = 0;
  // A line at a table's page break, `y` px down the layer: in the colour and at the width of the top line of the row
  // there (its first cell's; else the table's first cell's), a hairline at its own width on paper (--odx-hair). False
  // for a table with no lines.
  function breakRule(frag, t, r, origin, y) {
    const at = [...t.rows].find((tr) => { const q = rectOf(tr); return q.top - origin.top <= y + 1.5 && q.bottom - origin.top > y + 1.5; });
    const cell = (at && at.cells[0]) || t.querySelector('td') || t;
    const shadow = getComputedStyle(cell).boxShadow;
    if (!shadow || shadow === 'none') return false;
    let color = null, w = 1;
    for (const m of shadow.matchAll(/(rgba?\([^)]*\))\s+(-?[\d.]+)px\s+(-?[\d.]+)px/g)) {
      if (!color) color = m[1]; // (the first line named, unless a top line is)
      if (Number(m[2]) === 0 && Number(m[3]) > 0) { color = m[1]; w = Number(m[3]); break; }
    }
    const line = document.createElement('div');
    line.className = 'odx-repeat';
    // (As an inset shadow at the box's top, as the cells' own lines are: a box's background is drawn to whole pixels, a
    // hairline a full pixel thick on paper.)
    // (A hairline only where the cell's own line is one, styles.js: a line of a whole pixel, 0.75pt, is drawn at that
    // on paper too. Taken for a hairline, a table's last line on a page printed at 0.63px under rows whose lines
    // are 1px, and in Word's print 0.88px.)
    const hair = w < 1 || getComputedStyle(cell).getPropertyValue('--odx-hairtop').trim() === '1';
    const h = hair ? 'max(0.667px, var(--odx-hair, 1px))' : `${w}px`;
    line.style.cssText = `top:${y}px;left:${r.left - origin.left}px;width:${r.width}px;height:${Math.max(4, Math.ceil(w) + 2)}px;box-shadow:inset 0 ${h} 0 0 ${color || '#000'}`;
    frag.append(line);
    return true;
  }
  function drawHeads(list) {
    // (Unchanged repeats are left as they are: redrawing a long document's every pass took a third of its layout time.)
    const key = list.map((h) => { if (!nodeIds.has(h.b.node)) nodeIds.set(h.b.node, ++idSeq); return `${nodeIds.get(h.b.node)}:${h.b.i}:${h.b.top.toFixed(1)}:${h.top}:${h.rule ? 1 : 0}:${h.dy ?? ''}:${h.dx ?? ''}:${h.cut ?? ''}:${h.stop ?? ''}:${h.end ?? ''}:${h.rewrap ? `${h.rewrap.width.toFixed(1)}/${h.rewrap.h.toFixed(1)}` : ''}`; }).join(',');
    if (!list.length && !headKey) return;
    if (key === headKey && repeats.childElementCount) return;
    headKey = key;
    carriedParts = new WeakMap();
    const origin = rectOf(layer);
    const frag = document.createDocumentFragment();
    for (const { b, top, rule, cont, cut, stop, end, dx, dy, rewrap, col } of list) {
      if (rewrap) {
        // Its lines laid out again at the next column's width (rewrapCarry): from the top of that copy.
        const r = rectOf(b.el);
        const wrap = document.createElement('div');
        wrap.className = 'odx-repeat';
        wrap.style.cssText = `top:${(r.top - origin.top + dy + cut).toFixed(2)}px;left:${(r.left - origin.left + dx).toFixed(2)}px;width:${rewrap.width}px`;
        const host = document.createElement('div');
        host.className = 'odx-content';
        const copy = rewrap.el.cloneNode(true);
        host.append(copy);
        wrap.append(host);
        carriedParts.set(b.el, [{ copy, cut, stop: null, rewrap: true, col }]);
        frag.append(wrap);
        continue;
      }
      if (cont) {
        // A paragraph's lines past its column's bottom, again at the next column's top (columns: drawn, not editable there).
        const r = rectOf(b.el);
        const wrap = document.createElement('div');
        wrap.className = 'odx-repeat';
        wrap.style.cssText = `top:${(r.top - origin.top + dy).toFixed(2)}px;left:${(r.left - origin.left + dx).toFixed(2)}px;width:${r.width}px`;
        const host = document.createElement('div');
        host.className = 'odx-content';
        const copy = b.el.cloneNode(true);
        copy.removeAttribute('contenteditable');
        copy.style.removeProperty('margin-top');
        copy.style.removeProperty('--odx-push'); // (its page push, which editor.css makes an !important margin: margin 0 alone wouldn't take)
        copy.style.margin = '0';
        const to = stop ?? end; // (only down to there: the next part's cut, or a table's header rows)
        copy.style.clipPath = `inset(${cut.toFixed(2)}px ${-CLIP_SIDE}px ${to != null ? (r.height - to).toFixed(2) : -CLIP_TOP}px ${-CLIP_SIDE}px)`;
        copy.querySelectorAll('.ProseMirror-trailingBreak, .odx-push').forEach((e) => e.remove());
        host.append(copy);
        if (end == null) carriedParts.set(b.el, [...(carriedParts.get(b.el) || []), { copy, cut, stop, rewrap: false, col }]);
        wrap.append(host);
        frag.append(wrap);
        continue;
      }
      const t = b.table;
      if (rule) {
        // The bottom border of a row's part left behind at a page break, in the table's line colour.
        const r = rectOf(t);
        if (!breakRule(frag, t, r, origin, top - 1)) continue; // a borderless table gets no break rule
        continue;
      }
      const copy = t.cloneNode(true);
      copy.removeAttribute('contenteditable');
      const body = copy.tBodies[0];
      while (body.rows.length > b.headerRows) body.lastElementChild.remove();
      copy.querySelectorAll('.column-resize-handle, .selectedCell').forEach((n) => n.classList.remove('selectedCell') || (n.classList.contains('column-resize-handle') && n.remove()));
      const r = rectOf(t);
      const wrap = document.createElement('div');
      wrap.className = 'odx-repeat';
      wrap.style.cssText = `top:${top}px;left:${r.left - origin.left}px;width:${r.width}px`;
      wrap.append(copy);
      frag.append(wrap);
      // The line under the repeated header rows: the top line of the row that follows them on this page, whose own is
      // drawn at its cell's top, above its push (Word's print of a sample: a black line under the header row at the top
      // of the table's second page; there was none).
      let hh = 0;
      for (let k = 0; k < b.headerRows && k < t.rows.length; k++) hh += rectOf(t.rows[k]).height;
      breakRule(frag, t, r, origin, top + hh);
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
  const floatUp = new Map(); // block index -> { el, up: Map(picture's index in it -> px drawn up) }, this pass's (floatsPast)
  let drawnUp = new Set(); // pictures drawn up now
  // A float drawn up (floatsPast) is drawn there, but the browser still lays its paragraph's lines out round where it
  // is: on the next page, the lines it went on to were kept off its side too, where Word has them the column's width
  // (wrap-word2 wl-up, wl-upj, wr-up: page 2's first lines narrowed by a picture drawn on page 1). Its box ends that much
  // higher (a margin below it less by as much), at the page's edge, as drawn. Its own margin there kept, by picture.
  const ownMb = new WeakMap();
  const liftBox = (img, dy) => {
    if (getComputedStyle(img).float === 'none') return; // (a shape set out of the flow takes no room in it)
    if (!ownMb.has(img)) ownMb.set(img, img.style.marginBottom);
    const mb = dy > 0 ? `calc(${ownMb.get(img) || '0px'} - ${dy.toFixed(2)}px)` : ownMb.get(img);
    if (img.style.marginBottom !== mb) img.style.marginBottom = mb;
  };
  const shapeUp = new Map(); // block index -> { el, up }: its shapes with text on both sides drawn up (shapesPast), this pass's
  const tbUp = new Map(); // block index -> { img, up, top }: its picture wrapped top and bottom drawn up (tbPast), this pass's
  let tbOut = new Set(); // such pictures set out of the flow now
  // (Drawn up by its top margin less, not a translation: its box where it was drawn past the page's foot, Edge printed an
  // empty page more for it. Its own top margin kept, and how far it's drawn up now.)
  const tbOwnMt = new WeakMap();
  const tbUpNow = new WeakMap();
  let tbPasses = 0;
  let tbDoc = null; // (the count starts again when the document changes)
  let applied = new Map(); // block index -> { extra px we added above it }
  const keep = (node) => node.type === schema.nodes.paragraph && (node.attrs.view.keepNext ?? session.styleProp(node.attrs.styleId, 'keepNext'));
  const paraBreakBefore = (node) => node.type === schema.nodes.paragraph && (node.attrs.view.pageBreakBefore ?? session.styleProp(node.attrs.styleId, 'pageBreakBefore'));
  // A table row starts a page when the first paragraph of its first cell says to, as in Word (the row moves, under its
  // header rows; a paragraph in another cell, or after the first, doesn't), unless the row above keeps with it (a
  // paragraph in it keeps with the next); a table whose first row does, starts one.
  const rowBreaksBefore = (row) => !!row.firstChild?.firstChild && paraBreakBefore(row.firstChild.firstChild);
  // A row keeps with the next when the first paragraph of its first cell does, as in Word (keeprow-word: a heading row
  // of two paragraphs, only the second kept with the next, stays at the page's foot with its next row moved on, as
  // without a keep; only the first, it goes on with it; the same in a row of two cells, kept in the second or the first).
  const rowKeeps = (row) => {
    let first = null;
    row.firstChild?.descendants((n) => { if (first) return false; if (n.type === schema.nodes.paragraph) { first = n; return false; } return true; });
    return !!first && !!keep(first);
  };
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
  // The space before a paragraph keeps at the top of the page it starts by asking for one (page break before) or by
  // opening with a page break: in a document of Word 2010's compatibility or before, what it has over the space after of
  // the paragraph above (that went on the page before); from Word 2013's layout, none. Word's prints of 19 generated
  // documents, why-tests/fix2711 p1-pj: 10pt before under 10pt after, none; 20pt under 10pt, 10pt; under 5pt, 15pt;
  // 10pt under none, 10pt; mode 15, none in each. (After a page break in a paragraph of its own: none, both modes.) A
  // sample's mode-14 heading opening with its page's break, 10pt before under a 6pt-after caption, sat 5.3px high.
  const breakSpace = (node, prev) => (session.compatMode >= 15 ? 0 : Math.max(0, spaceBefore(node) - (prev ? spaceAfter(prev) : 0)));
  // How many rows a table starts with that repeat at each page's top (w:tblHeader).
  // (A row that says nothing of its own is one when its table's style makes its rows header rows: Word's prints of 12
  // generated documents, the style's setting with the first row's own on, off or absent. Every row a header row: none
  // repeats here; Word then repeats the first on the second page only.)
  const headerRowsOf = (table) => {
    const byStyle = styleHeaderRows(session.styles, table.attrs.tblPr);
    let h = 0;
    while (h < table.childCount && isHeader(table.child(h), byStyle)) h++;
    return h;
  };
  const isHeader = (row, byStyle = false) => (/<w:tblHeader\b/.test(row.attrs.trPr || '') ? /<w:tblHeader(?![^>]*w:val="(?:0|false|off)")/.test(row.attrs.trPr) : byStyle);
  // (In newspaper columns a page break goes on at the next page's first column, a column break at the next column.)
  const breakKindOf = (node) => {
    let kind = null;
    if (node.type === schema.nodes.paragraph) node.forEach((c) => { if (c.type === schema.nodes.hard_break && c.attrs.type === 'page') kind = 'page'; else if (c.type === schema.nodes.hard_break && c.attrs.type === 'column' && !kind) kind = 'column'; });
    return kind;
  };
  const hasPageBreak = (node) => {
    let found = false;
    if (node.type === schema.nodes.paragraph) node.forEach((c) => { if (c.type === schema.nodes.hard_break && (c.attrs.type === 'page' || c.attrs.type === 'column')) found = true; });
    return found;
  };

  // Header and footer line heights as the body's (lhfix.js), not on a document grid of lines.
  const storyLines = (el) => { if (!lineGrid(view.state.doc)) storyLineFix(el, (p) => hfNodes.get(p)); };
  const heights = new Map();
  // (`dist`: a header's distance from the page's top, where it is drawn: its shapes placed against the page go by it,
  // hfstory.js dropBelowLogo.)
  const storyHeight = (doc, footer = false, dist = 0) => {
    if (!doc) return 0;
    const key = footer || !dist ? doc : `${dist}`;
    let byDoc = heights.get(doc);
    if (!byDoc) heights.set(doc, (byDoc = new Map()));
    if (!byDoc.has(key)) {
      const el = story(doc, 'odx-header', 1, 1);
      if (footer) el.dataset.odxStory = 'footer'; // (its shapes as a footer's: hfstory.js settleStoryShapes)
      el.style.cssText = `visibility:hidden;top:${footer ? 0 : dist}px;left:${page.left}px;right:${page.right}px`;
      sheet.append(el);
      storyLines(el);
      settleStoryShapes(el); // (a shape the body wraps round, below its paragraph: not in its height)
      clipExactRows(el); // (its exact rows as drawn: a cell's overflow doesn't grow them)
      fitRowspans(el);
      // Its tabs to their stops, as when it is drawn (below): left as plain tab characters, each kept at its full width
      // now that header and footer spaces are (editor.css), a footer line of a page number, a tab and text that just
      // fits wrapped here alone, the footer measured a line taller than it is drawn and every page a line short.
      // (Its text at Word's widths first, as the body's (widths.js): the tabs go after it.)
      fixStoryWidths(el);
      try { layoutTabsIn(el, session, (p) => { const n = hfNodes.get(p); if (!n) throw new Error('no node'); return n; }); } catch { /* measured as it is */ }
      layoutPtabs(el); // (positional tabs after the ordinary ones)
      byDoc.set(key, flowHeight(el)); // (what floats in it doesn't count: hfstory.js flowHeight)
      el.remove();
    }
    return byDoc.get(key);
  };
  let sectCache = { doc: null, ends: new Map(), first: page };

  /**
   * Floats placed against a page or its margins (reader.js anchorCss: --odx-lrel), not their paragraph: Word draws them
   * on the page their paragraph is on, at that height, and that page's lines wrap round them, those above the paragraph
   * too. Each is drawn again at the start of the page's first paragraph, below a hairline-wide float as tall as the drop
   * to its height, and the original is hidden. Positions are read from the current layout (one pass behind: the next
   * pass settles them).
   */
  // Where a block's first line is (px, screen): its box's top, unless pagination moved its lines to the next page by a
  // spacer inside it (a paragraph whose box still starts at the foot of the page before: pagefloat-word pf-early's anchor,
  // its picture lifted onto the wrong page).
  const firstLineTop = (el) => {
    const top = rectOf(el).top;
    const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.data.trim() && !n.parentElement.closest('.odx-lift, .odx-push') ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP) });
    const t = w.nextNode();
    if (!t) return top;
    const rg = document.createRange();
    rg.setStart(t, 0); rg.setEnd(t, Math.min(1, t.data.length));
    const q = rectsOf(rg)[0];
    return q ? Math.max(top, q.top) : top;
  };
  // The page each float placed against the page went on, for this document (its image node -> flow page): never an
  // earlier one again, and its paragraph starts there. Word's way out of the loop where the float, lifted onto its
  // paragraph's page, wraps the text above so far that the paragraph moves to the next page: the float goes on with it,
  // the page before keeping its room (pagefloat-word pf-early: Text 7 starts page 2 under its picture, page 1 ends
  // short). Reset when the document changes.
  let liftMemo = new WeakMap();
  let liftMemoDoc = null;
  /** The latest page a float placed against the page in this paragraph went on (liftMemo), or -1. */
  const liftedPage = (node) => {
    if (liftMemoDoc !== view.state.doc) return -1;
    let out = -1;
    node.forEach((c) => { if (c.type === schema.nodes.image && placeVar(c.attrs.place, 'odx-lrel') != null) out = Math.max(out, liftMemo.get(c) ?? -1); });
    return out;
  };
  const liftFloats = (base) => {
    if (liftMemoDoc !== view.state.doc) { liftMemo = new WeakMap(); liftMemoDoc = view.state.doc; }
    const out = [];
    view.state.doc.descendants((node, pos) => {
      if (node.type !== schema.nodes.image || placeVar(node.attrs.place, 'odx-lrel') == null) return true;
      const img = view.nodeDOM(pos);
      const host = img && img.nodeType === 1 ? img.closest('.odx-content > *') : null;
      if (!host || host.parentElement !== view.dom) return false;
      const prevPg = liftMemo.get(node);
      const pg = Math.max(pageOf(firstLineTop(host) - base), prevPg ?? -1);
      liftMemo.set(node, pg);
      const place = node.attrs.place;
      const rel = placeVar(place, 'odx-lrel');
      const off = parseFloat(placeVar(place, 'odx-ly')) || 0;
      const align = placeVar(place, 'odx-lyalign');
      const h = Number(node.attrs.height) || 0;
      const m = marginsAt(pg);
      const top0 = pageTop(pg) + (rel === 'page' ? 0 : m.top);
      const bottom0 = pageTop(pg) + pageH(pg) - (rel === 'page' ? 0 : m.bottom);
      const y = alignedY(align, top0, bottom0, h, off);
      // (The page's first paragraph by where its first line is, and the drop from the page's text top, where pagination puts
      // that line: measured from the paragraph's box, a pass behind, the lift and the push chased each other and never
      // settled, pagefloat-word pf-early's picture drawn on the page before.)
      const first = [...view.dom.children].find((e) => e.tagName === 'P' && firstLineTop(e) - base >= contentTop(pg) - 1);
      if (!first || pageOf(firstLineTop(first) - base) !== pg) return false;
      // (A shape drawn with room round its outline, reader.js padPlace: its picture that much larger and higher.)
      const pad = padOf(place);
      const drop = Math.round((y - pad - Math.max(contentTop(pg), rectOf(first).top - base)) * 10) / 10;
      const at = view.posAtDOM(first, 0);
      const cs = getComputedStyle(img);
      const side = cs.float === 'right' ? 'right' : 'left';
      const src = img.getAttribute('src') || '';
      const style = `float:${side};clear:${side};width:${node.attrs.width + 2 * pad}px;height:${h + 2 * pad}px;margin:${Math.min(0, drop)}px ${cs.marginRight} ${cs.marginBottom} ${cs.marginLeft}`;
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
      if (node.type !== schema.nodes.image || placeVar(node.attrs.place, 'odx-arel') == null) return true;
      const img = view.nodeDOM(pos);
      const host = img && img.nodeType === 1 ? img.closest('.odx-content > *') : null;
      const cb = img && img.offsetParent;
      if (!host || host.parentElement !== view.dom || !cb) return false;
      const pg = pageOf(firstLineTop(host) - base);
      const place = node.attrs.place;
      const rel = placeVar(place, 'odx-arel');
      const off = parseFloat(placeVar(place, 'odx-ay')) || 0;
      const align = placeVar(place, 'odx-ayalign');
      const h = Number(node.attrs.height) || 0;
      const m = marginsAt(pg);
      const p0 = pageTop(pg), p1 = p0 + pageH(pg);
      const [top0, bottom0] = rel === 'page' ? [p0, p1] : rel === 'topMargin' ? [p0, p0 + m.top] : rel === 'bottomMargin' ? [p1 - m.bottom, p1] : [p0 + m.top, p1 - m.bottom];
      const y = alignedY(align, top0, bottom0, h, off);
      const top = Math.round((y - (rectOf(cb).top - base)) * 10) / 10;
      // (Less the room drawn round a shape's outline: reader.js padPlace.)
      const pad = padOf(place);
      out.push(Decoration.node(pos, pos + node.nodeSize, { style: `top:${top - pad}px;margin-top:0` }, { key: `onpage:${top - pad}` }));
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
      const r = rectOf(el);
      const a = applied.get(i);
      // (A floating block's margin moves only itself: what follows it in the flow is where it was.)
      const float = a && a.extra && outOfFlow(el);
      if (a && !float) undo += a.extra;
      const block = { i, node, el, top: r.top - base - undo - (float ? a.extra : 0), undo: undo + (float ? a.extra : 0) };
      if (node.type === schema.nodes.table) {
        // Rows too, so a table can continue on the next page.
        const table = tableOf(el);
        const trs = table && table.tBodies[0] ? [...table.tBodies[0].rows] : [];
        if (table && trs.length === node.childCount && !cols && !inRegion(i)) { // (in columns, tables move whole)
          block.table = table;
          block.tableTop = rectOf(table).top - base - undo;
          let prevBottom = block.tableTop;
          block.rows = trs.map((tr, j) => {
            // A row that isn't shown (a tracked deletion printed without markup) has no box: it takes no room, where the
            // row above it ends. (Measured as it was, at the window's top, the table had no rows and never split: a
            // sample's table with a deleted row ran past the page's foot to the sheet's edge, as Word moves it on.)
            if (!tr.getClientRects().length) return { top: prevBottom, bottom: prevBottom, undo, line: 0, topLine: 0, ownLine: 0, hidden: true };
            // A pushed row's extra padding is inside it: it moves its bottom, not its top.
            const rr = rectOf(tr);
            const rowTop = rr.top - base - undo;
            const d = a && a.rows && a.rows.get(j);
            if (d) undo += d;
            const contentUndo = undo;
            let bottom = rr.bottom;
            if (a && a.inRow && a.inRow.has(j)) {
              // Split inside its cells: the row's natural bottom is its tallest cell without our spacers.
              bottom = -Infinity;
              for (const td of tr.cells) {
                const last = lastShown(td);
                if (!last) continue;
                const spacers = undoIn(td, Infinity); // (in a table in the cell, a row's tallest cell's)
                // (Its last paragraph's space after is in the cell too: margins don't pass through a cell's edge.)
                const after = Math.max(0, parseFloat(getComputedStyle(last).marginBottom) || 0);
                bottom = Math.max(bottom, rectOf(last).bottom - spacers + after + (parseFloat(getComputedStyle(td).paddingBottom) || 0));
              }
              if (!(bottom > rr.top)) bottom = rr.bottom;
            }
            // The line under the row is the next row's (styles.js), but a row ending a page has to have room for it.
            const line = Math.max(0, ...[...tr.cells].map((td) => parseFloat(td.style.getPropertyValue('--odx-bdb')) || 0));
            // (And the widest line over it and the widest bottom line a cell of it draws itself, --odx-tl and --odx-bob: a row
            // of "at least" height splits only with room for both.)
            const topLine = Math.max(0, ...[...tr.cells].map((td) => parseFloat(td.style.getPropertyValue('--odx-tl')) || 0));
            const ownLine = Math.max(0, ...[...tr.cells].map((td) => parseFloat(td.style.getPropertyValue('--odx-bob')) || 0));
            // (And the table's bottom line, --odx-tbb, which Word draws under the row's part at a page's foot.)
            const tableLine = Math.max(0, ...[...tr.cells].map((td) => parseFloat(td.style.getPropertyValue('--odx-tbb')) || 0));
            const row = { top: rowTop, bottom: bottom - base - undo, undo: contentUndo, line, topLine, ownLine, tableLine };
            undo += rr.bottom - bottom;
            prevBottom = row.bottom;
            return row;
          });
          if (block.rows.some((x) => !x.hidden && x.bottom <= x.top) || block.rows.every((x) => x.hidden)) block.rows = null; // not laid out (content-visibility)
          const h = headerRowsOf(node);
          block.headerRows = h < node.childCount ? h : 0;
          block.allHeader = h > 1 && h >= node.childCount;
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
        blockNotes.get(i).push({ ...r, h: noteHeight(note, r.kind, r.n), note, bi: i });
      }
    }
  }

  // The end of a columns section, at block k (top: where it would go): its last flow page and sheet. Balanced, what
  // follows goes under its lowest column there (on the flow page of its last column, so the sheet's other columns stay
  // where they are): `moved` is how far block k moves down (also added to `starts`); `again`, that the column height
  // its last sheet is cut to changed, so layout runs again.
  function endColumns(re, k, b, top, blocks, rec, sectEnds, starts) {
    const breaksFirst = (node) => node.type === schema.nodes.paragraph && node.firstChild?.type === schema.nodes.hard_break && node.firstChild.attrs.type === 'column';
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
      // (An empty section-break paragraph: Word's columns end above it. Unless it's all the section holds: then it keeps
      // its line, as in one column.)
      if (j > re.k0 && !blocks[j].node.childCount && sectEnds.get(blocks[j].i)) continue;
      const parts = x.parts || (x.carry ? [{ off: 1, h: x.carry, stop: null }] : []);
      // (A paragraph that starts with a column break leaves nothing in the column before: its text is all in the next.
      // Its first line, the break's, counted there, what followed a sample's code lists went a line low: sample 1113.)
      const lead = parts.length && breaksFirst(blocks[j].node);
      if (!lead) {
        bottoms.set(v, Math.max(bottoms.get(v) ?? -Infinity, x.bottom));
        ends.set(v, Math.max(ends.get(v) ?? -Infinity, x.bottom + (x.carry ? 0 : spaceAfter(blocks[j].node)))); // (a paragraph cut there has none)
      }
      // (Its parts carried on, each at a column's top: the last ends with its space after.)
      for (const q of parts) {
        const y = contentTop(v + q.off) + q.h;
        bottoms.set(v + q.off, Math.max(bottoms.get(v + q.off) ?? -Infinity, y));
        ends.set(v + q.off, Math.max(ends.get(v + q.off) ?? -Infinity, y + (q.stop == null ? spaceAfter(blocks[j].node) : 0)));
        vEnd = Math.max(vEnd, v + q.off);
      }
    }
    const sheet = Math.floor((vEnd - re.P) / re.num);
    re.vEnd = vEnd;
    touchNotes(Math.max(0, re.P, vEnd - re.num)); // (the bands from its last sheet on: its columns end there)
    re.offAfter = vEnd - (re.S0 + sheet);
    for (const v of [...colTops.keys()]) if (v > vEnd) colTops.delete(v);
    // (What follows on the same page starts under the lowest column; balanced, the columns are also cut even.)
    if (re.balance || re.cont) {
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
      addTo(starts, b.i, d);
      moved = d;
      if (!re.balance) return { moved, again };
      // The height to cut its columns to: an even share, and half a line more so the first column takes the odd one.
      const line = parseFloat(lineHeightOf(getComputedStyle(blocks[re.k0].el))) || 18;
      const want = { sheet, h: total / re.num + line / 2 };
      const had = balanced.get(re.lo);
      // (Word raises the cut until the last column is no taller than it: a paragraph that keeps two lines a side may not
      // split at an even share, and moved whole it overfills the last column.)
      if (had && had.sheet === sheet && vEnd > first) {
        const last = (bottoms.get(vEnd) ?? base(vEnd)) - base(vEnd);
        if (last > had.h + 1) want.h = Math.max(want.h, had.h + line);
        else want.h = Math.max(want.h, had.h > want.h && had.raised ? had.h : want.h);
        want.raised = want.h > total / re.num + line / 2 + 0.5;
      }
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
    touchNotes(Math.max(0, P - 1)); // (the bands from there on are its sheets now)
    rs.y0 = y0;
    rs.off0 = n ? regs[n - 1].offAfter || 0 : 0;
    rs.S0 = P - rs.off0;
    for (let c = 1; c < rs.num; c++) colTops.set(P + c, pageTop(P + c) + (y0 - pageTop(P)));
    // A section set to start on the page after this one (a continuous one part way down it) starts past its columns:
    // its "next page" is no column of this page. (As the flow page after it, a continuous section's own margins and
    // footer were the page's second column's, and what followed the columns there was cut at them: sample 1113's form,
    // its page a line short, the pages after it a page late.)
    if (!newPage) {
      for (const m of margins) if (m[0] > P && m[0] < P + rs.num) m[0] = P + rs.num;
      for (const m of sections) if (m[0] > P && m[0] < P + rs.num) m[0] = P + rs.num;
    }
    const bal = rs.balance && balanced.get(rs.lo);
    if (bal) for (let c = 0; c < rs.num - 1; c++) { const v = P + bal.sheet * rs.num + c; cuts.set(v, (v === P ? y0 : contentTop(v)) + bal.h); }
  }

  // Columns: each block moved up beside the column before it, the lines a block carries into the next column (drawn
  // again there), and the page count, sections and page setup counted in sheets. Each block's move, cut and column width
  // go on the block itself (`blockAttrs`: block index -> { cls, style }, node decorations, applyPushes), not as rules by
  // its place among its siblings (:nth-child): those made the browser restyle every block after one added.
  function placeColumns(blocks, regs, starts, colSplits, paraSplits, splits) {
    const colCss = [];
    const contHeads = [];
    const blockAttrs = new Map();
    const attrsOf = (i) => { if (!blockAttrs.has(i)) blockAttrs.set(i, { cls: [], style: [] }); return blockAttrs.get(i); };
    sheetBand = null;
    blockCols = new WeakMap();
    const live = regs.filter((r) => r.P != null);
    if (cols || live.length) {
      // Each block's column: its page in the flow. Move it up beside the column before it. A columns section's columns
      // count from the page it starts on; what follows it moves up by the flow pages its columns took beyond their sheets.
      const geo = (k) => (cols ? { P: 0, S0: 0, off0: 0, ...cols } : live.find((r) => k >= r.k0 && k < r.k1));
      const offAt = (k) => { let o = 0; for (const r of live) if (k >= r.k1 && r.offAfter != null) o = r.offAfter; return o; };
      const place = (k, v) => {
        const g = geo(k);
        if (!g) return { s: v - offAt(k), c: 0, w: 0, x: 0 };
        // (narrow: how much narrower than the widest its column is, columns of their own widths.)
        if (v < g.P) return { s: v - g.off0, c: 0, w: 0, x: 0, narrow: g.unequal ? g.width - g.widths[0] : 0, before: true };
        const c = (v - g.P) % g.num;
        return { s: g.S0 + Math.floor((v - g.P) / g.num), c, w: g.width + g.space, x: colX(g, c), cw: g.widths[c], narrow: g.unequal ? g.width - g.widths[c] : 0 };
      };
      // How far up a block on flow page v goes to its sheet, `up` flow pages above: from its flow page's top to its sheet's,
      // each sheet as tall as its own page setup. (One page's pitch for all, a landscape last section's, a portrait form's
      // second column stood 240px low, the page under it a page late: sample 1113. The heights of the `up` flow pages
      // just above it, its landscape pages after the form's columns stood 240px low each: the pages the columns took were
      // portrait ones.)
      const bySheetSetup = () => {
        const out = [];
        for (const [f, m, fresh] of margins) {
          const at = sheetOfFlow(f);
          if (out.length && out[out.length - 1][0] === at) out.pop(); // (the later one wins a sheet)
          out.push([at, m, fresh]);
        }
        return out;
      };
      const sheetSetup = cols ? null : bySheetSetup();
      const liftOf = (v, up) => (!up ? 0 : sheetSetup ? pageTop(v) - pageTop(v - up, sheetSetup) : pageTop(v) - pageTop(v - up));
      const narrowNow = new Map();
      let sh = 0;
      let lastS = blocks.length ? place(blocks.length - 1, pages - 1).s : 0;
      for (let k = 0; k < blocks.length; k++) {
        const b = blocks[k];
        sh += starts.get(b.i) || 0;
        const v = pageOf(b.top + sh);
        const { s, x, cw, narrow } = place(k, v);
        if (cw) blockCols.set(b.el, { x, w: cw }); // (its column's text edges, for its line numbers)
        const up = v - s; // pages of flow above this column's sheet
        // Its column's width (measured with it too, so a block that lands in another column is measured again there).
        if (narrow) { narrowNow.set(b.i, narrow); attrsOf(b.i).style.push(`--odx-coln:${narrow.toFixed(2)}px`); }
        if (up || x) { const a = attrsOf(b.i); a.cls.push('odx-colmv'); a.style.push(`--odx-colt:translate(${x.toFixed(1)}px,${-liftOf(v, up)}px)`); }
        const cs = colSplits.get(b.i);
        if (cs) {
          const a = attrsOf(b.i);
          a.cls.push('odx-colclip');
          a.style.push(`--odx-colclip:inset(${-CLIP_TOP}px ${-CLIP_SIDE}px ${(cs.height - cs.cut).toFixed(2)}px ${-CLIP_SIDE}px)`);
          const head = cs.head || 0;
          for (const part of carriedOf(cs)) {
            const v2 = v + part.off, q = place(k, v2), up2 = v2 - q.s;
            lastS = Math.max(lastS, q.s); // (the lines it carries over count for the page count)
            const va = colVa.get(v2) || 0; // (centred on the page: verticalAlignCols)
            contHeads.push({ b, cont: true, cut: part.cut, stop: part.stop, rewrap: cs.rewrap, col: q.cw ? { x: q.x, w: q.cw } : null, dx: q.x - x, dy: (contentTop(v2) + va + head + (cs.rewrap ? 0 : cs.lift || 0) - liftOf(v2, up2)) - (b.top + sh + part.cut - liftOf(v, up)) });
            if (head) contHeads.push({ b, cont: true, cut: 0, end: head, dx: q.x - x, dy: (contentTop(v2) + va - liftOf(v2, up2)) - (b.top + sh - liftOf(v, up)) });
          }
        }
        if (!cols) sh += (paraSplits.get(b.i)?.total || 0) + (splits.get(b.i)?.total || 0); // (what a block split over pages pushes below it)
      }
      pages = lastS + 1;
      // Sections start on sheets, not flow pages.
      // Each sheet's band of notes: the flow page of its first column.
      sheetBand = new Map();
      for (let v = 0; v < MAX_FLOW_PAGES; v++) {
        const s = sheetOfFlow(v);
        if (s >= pages) break;
        if (!sheetBand.has(s)) sheetBand.set(s, v);
      }
      if (!cols) {
        sections = sections.map(([f, n]) => [sheetOfFlow(f), n]);
        // And their page setup (headers, footers, size, numbering) with them: a section after a columns section had its
        // header and numbers on the sheet its flow page would have been.
        margins = bySheetSetup();
      }
      colCss.push(`.${scope} > .odx-content{height:${pageTop(pages) - gap}px;overflow:visible}`);
      // (Not while the flow is measured: odx-measure. As specific as the rules by position were.)
      colCss.push(`.${scope} > .odx-content:not(.odx-measure) > .odx-colmv{transform:var(--odx-colt)}`, `.${scope} > .odx-content:not(.odx-measure) > .odx-colclip{clip-path:var(--odx-colclip)}`);
      // Blocks that changed column width: lay out again at the new widths.
      const moved = narrowNow.size !== narrowed.size || [...narrowNow].some(([i, n]) => Math.abs((narrowed.get(i) ?? 0) - n) > 0.01);
      narrowed = narrowNow;
      if (moved) colsAgain = true;
    }
    return { colCss, contHeads, blockAttrs };
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
    for (const list of pulls.values()) for (const [i, d] of list) addTo(extra, i, d);
    // Where each block ends up, and its section.
    let acc = 0, sect = 0;
    const at = blocks.map((b) => {
      acc += extra.get(b.i) || 0;
      const split = (paraSplits.get(b.i)?.total || 0) + (splits.get(b.i)?.total || 0);
      const r = { b, top: b.top + acc, bottom: b.bottom + acc + split, split: split !== 0, sect, after: parseFloat(getComputedStyle(b.el).marginBottom) || 0 };
      // How low it reaches with what floats in it (a text box, a picture): Word centres those too.
      r.reach = r.bottom;
      const er = rectOf(b.el);
      for (const f of b.el.querySelectorAll('img')) {
        if (getComputedStyle(f).position !== 'absolute' && getComputedStyle(f).float === 'none') continue;
        r.reach = Math.max(r.reach, b.top + acc + (rectOf(f).bottom - er.top));
      }
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
      // (The lowest of its last block, with its space after, and anything floating on the page: a frame, a text box.)
      const slack = contentBottom(p) - Math.max(end.bottom + end.after, ...g.map((r) => r.reach + (outOfFlow(r.b.el) ? r.after : 0)));
      if (slack <= 0.5) continue; // (a full page: nothing to move)
      const moves = v === 'both' ? g.slice(1).map((r) => [r, slack / (g.length - 1)]) : [[g[0], v === 'center' ? (slack - (g[0].top - contentTop(p))) / 2 : slack]];
      let total = 0;
      for (const [r, d] of moves) { addTo(starts, r.b.i, d); total += d; }
      const next = at.find((r) => r.p > p); // (the next page's first block; an empty section-ending paragraph goes down with the rest)
      if (next && total) addTo(starts, next.b.i, -total);
    }
  }

  // A one-section document in columns: each block where it ends up (`at`: { b, top, v (its flow page) }), and each flow
  // page's column's lowest end (`ends`), the lines a block carries into the columns after it counted there; `withAfter`:
  // a block's bottom margin with it (vertical alignment counts it, the balance before endnotes doesn't).
  const colEnds = (blocks, starts, colSplits, withAfter) => {
    let acc = 0;
    const at = [];
    const ends = new Map(); // flow page -> how low its column's content reaches
    for (const b of blocks) {
      acc += starts.get(b.i) || 0;
      const cs = colSplits.get(b.i);
      const top = b.top + acc;
      const v = pageOf(top);
      at.push({ b, top, v });
      ends.set(v, Math.max(ends.get(v) ?? -Infinity, cs ? top + cs.cut : b.bottom + acc + (withAfter ? parseFloat(getComputedStyle(b.el).marginBottom) || 0 : 0)));
      if (cs) for (const q of carriedOf(cs)) ends.set(v + q.off, Math.max(ends.get(v + q.off) ?? -Infinity, contentTop(v + q.off) + q.h));
    }
    return { at, ends };
  };
  // A one-section document in newspaper columns centred (or set at the bottom) on its pages: each sheet's columns moved
  // down together by the room under its fullest column, as Word does (the text a column carries on into the next too).
  let colVa = new Map(); // flow page -> px its column's content is moved down
  function verticalAlignCols(doc, blocks, starts, colSplits) {
    const v = /<w:vAlign\b[^>]*w:val="(center|bottom)"/.exec(doc.attrs.sectPr || '')?.[1];
    if (!v || sectionXmls(doc).length > 1) return;
    const num = cols.num;
    const { at, ends } = colEnds(blocks, starts, colSplits, true);
    const sheets = new Set([...ends.keys()].map((q) => Math.floor(q / num)));
    for (const sh of sheets) {
      let slack = Infinity;
      for (let q = sh * num; q < (sh + 1) * num; q++) if (ends.has(q)) slack = Math.min(slack, contentBottom(q) - ends.get(q));
      if (!Number.isFinite(slack) || slack <= 0.5) continue;
      const first = at.find((r) => Math.floor(r.v / num) === sh);
      if (!first) continue;
      const d = v === 'center' ? (slack - (first.top - contentTop(first.v))) / 2 : slack;
      if (d <= 0.5) continue;
      addTo(starts, first.b.i, d);
      for (let q = sh * num; q < (sh + 1) * num; q++) colVa.set(q, d);
      const next = at.find((r) => Math.floor(r.v / num) > sh);
      if (next) addTo(starts, next.b.i, -d);
    }
  }

  // Columns before endnotes (a one-section document in columns): Word ends the text's columns at an even height on its
  // last sheet, and the endnotes start under the lowest of them, across the page (the generated documents'
  // endnote-content with columns: Word's second column starts with the first's last lines, the note under both).
  let colsBal = null; // { sheet, h, bottom }: the last sheet, the height its columns are cut to, where the lowest ends
  function balanceEndCols(blocks, starts, colSplits) {
    const has = [...blockNotes.values()].some((l) => l.some((r) => r.kind === 'endnote'));
    if (!has) { if (colsBal) { colsBal = null; colsAgain = true; } return; }
    const num = cols.num;
    const { ends } = colEnds(blocks, starts, colSplits, false);
    if (!ends.size) return;
    const lastV = Math.max(...ends.keys());
    const sheet = Math.floor(lastV / num);
    let total = 0, low = 0;
    for (let q = sheet * num; q < (sheet + 1) * num; q++) if (ends.has(q)) { total += ends.get(q) - contentTop(q); low = Math.max(low, ends.get(q) - contentTop(q)); }
    // (An even share, and half a line more so the first column takes the odd one, as for a balanced section.)
    const h = Math.min(total / num + END_COLS_EXTRA, contentBottom(sheet * num) - contentTop(sheet * num));
    const was = colsBal;
    if (!was || was.sheet !== sheet || Math.abs(was.h - h) > 1) { colsBal = { sheet, h, bottom: low }; colsAgain = true; }
    else colsBal.bottom = low;
  }

  // A band of footnotes in columns too tall for its page, as Word lays it out: the sheet's text up to the reference's
  // line evened out across its columns (the first taking the extra, as before endnotes), every column cut there, and the
  // band from under the lowest of them, the rest of its notes going on on the next page (fncols-word tall2a: two
  // paragraphs of ten lines, the second with a note of 50 sentences, side by side over the band; not one under the
  // other with the band under them). For the next pass, which lays the columns out again when it changed.
  function heldBands(blocks, starts, colSplits) {
    const now = new Map();
    const held = [...noteFlow].filter(([, f]) => f.held);
    if (held.length) {
      const { at, ends } = colEnds(blocks, starts, colSplits, false);
      for (const [b] of held) {
        const g = colGeo(b);
        if (!g || g.num < 2) continue;
        const r = (notesOn.get(b) || []).filter((x) => x.at != null && x.kind === 'footnote').sort((x, y) => x.at - y.at).pop();
        const e = r && at.find((q) => q.b.i === r.bi);
        if (!e) continue;
        // (Where the reference's block ends: in the last column it's carried into, when it's split.)
        const cs = colSplits.get(e.b.i);
        const parts = cs ? carriedOf(cs) : [];
        const lastPart = parts[parts.length - 1];
        const vRef = lastPart ? e.v + lastPart.off : e.v;
        const endRef = lastPart ? contentTop(vRef) + lastPart.h : e.top + (e.b.bottom - e.b.top);
        if (vRef < b || vRef >= b + g.num) continue;
        let total = endRef - contentTop(vRef), low = total;
        for (let v = b; v < vRef; v++) { const h = (ends.get(v) ?? contentTop(v)) - contentTop(v); total += h; low = Math.max(low, h); }
        // (Not lower than the reference's own lines when they start the sheet: nothing comes before them to even out,
        // tall2b's page 2, two lines carried over with the reference, a column of two lines beside them.)
        const partTop = lastPart ? contentTop(vRef) : e.top;
        const cut = Math.max(total / g.num + END_COLS_EXTRA, vRef === b && partTop <= contentTop(vRef) + 1 ? endRef - contentTop(vRef) + 1 : 0);
        const gapUnder = spaceAfter(e.b.node) + 11.5;
        now.set(b, { cut, top: Math.min(low, cut) + gapUnder, cuts: Array.from({ length: g.num }, (_, c) => [b + c, contentTop(b + c) + cut]) });
      }
    }
    for (const b of new Set([...now.keys(), ...heldPrev.keys()])) {
      const x = now.get(b), y = heldPrev.get(b);
      if (!x || !y || Math.abs(x.cut - y.cut) > 0.5 || Math.abs(x.top - y.top) > 0.5) colsAgain = true;
    }
    heldPrev = now;
  }

  // The walk's result applied: margin-tops (node decorations), split tables' padded rows and clip paths (CSS), spacer
  // widgets inside split paragraphs and cells, baseline shifts and lifted floats. Returns the CSS and the header rows to
  // draw again.
  const outOfFlow = (el) => { const cs = getComputedStyle(el); return cs.float !== 'none' || cs.position === 'absolute' || cs.position === 'fixed'; };
  // Whether a floating table leaves room beside it for text. Word wraps only short words into a strip under half an inch
  // (single letters from 18px, a four-letter word not in 22px), and none into 16px; the text after such a table goes
  // below it, and the table breaks across pages row by row as one in the flow does (generated documents: a table of 40
  // rows from mid-page, 24px narrower than the column, split at the page's foot; two in a row each split too). Laid out
  // in the flow, it does both here.
  const ROOM_BESIDE = 48;
  // Laid out in the flow, such a table keeps the gap a float leaves under it: its distance from the text (bottomFromText)
  // and no more. The space before the paragraph after it, which a float's text below it doesn't add (it starts beside the
  // float, as Word's anchor paragraph does), is taken back by a negative margin collapsing with it (samples: text 22px
  // and 30px low under such tables).
  const inFlow = new WeakMap(); // table node -> its margin-bottom in the flow
  // One placed down from the text it's anchored to (over 100tw, styles.js) stays a float: the text after it fills the
  // room above it, as in Word (sample: a table 451tw down, the text after it 30px low in the flow).
  function setDown(table) {
    const fp = /<w:tblpPr\s[^>]*>/.exec(table.attrs.tblPr || '')?.[0] || '';
    const y = /\sw:tblpY="(-?\d{1,6})"/.exec(fp);
    return !!y && +y[1] > 100 && /\sw:vertAnchor="text"/.test(fp) && !/\sw:tblpYSpec=/.test(fp);
  }
  // One placed up from the text it's anchored to (tblpY below 0) doesn't break where it starts on a page: Word moves it
  // to the next page whole when it doesn't fit in what's left, its anchor paragraph with it, and breaks it there only
  // when it's longer than a page. In a document of Word 2010's compatibility or before, only when there's no room
  // beside it (why-tests flt0898: tables of 10 to 70 rows, 1, 59 and 400tw up, 0 to 94pt beside them, all moved on in
  // mode 15; at 0 and 200tw down all broke at the foot. A sample's 31-row table had started half way down a page.)
  function upFromText(table) {
    const fp = /<w:tblpPr\s[^>]*>/.exec(table.attrs.tblPr || '')?.[0] || '';
    return /\sw:tblpY="-\d{1,6}"/.test(fp) && /\sw:vertAnchor="text"/.test(fp) && !/\sw:tblpYSpec=/.test(fp);
  }
  const upLift = (table) => Math.max(-1440, twipsToPx(+/\sw:tblpY="(-\d{1,6})"/.exec(table.attrs.tblPr || '')[1])); // (as styles.js draws it)
  const movesWhole = (b) => upFromText(b.node) && (session.compatMode >= 15 || !roomBeside(b.el));
  // (Unless it reaches past its page's foot: Word breaks it there row by row, and the text after it goes under it, as in
  // the flow, its drop kept above it; generated documents. The walk finds it so, crossFoot, and the pass runs again.)
  const crossFoot = new WeakSet();
  function inFlowGap(table, next) {
    const m = /<w:tblpPr\s[^>]*bottomFromText="(\d{1,6})"/.exec(table.attrs.tblPr || '');
    // (And a drop down from the text under 100tw, drawn that far down by its top (styles.js): the text or table after it
    // goes under it as drawn, in Word (new documents, tables 40, 80 and 99tw down with a table or a paragraph after
    // it: that had come 2-6px high, the drop's height).)
    const fp = /<w:tblpPr\s[^>]*>/.exec(table.attrs.tblPr || '')?.[0] || '';
    const y = /\sw:tblpY="(\d{1,6})"/.exec(fp);
    const drop = y && /\sw:vertAnchor="text"/.test(fp) && !setDown(table) ? twipsToPx(Math.min(+y[1], 31680)) : 0;
    const bft = m ? twipsToPx(Math.min(+m[1], 31680)) : 0; // (at most 22in)
    const b = next ? spaceBefore(next) : 0;
    return `${+(drop + Math.max(bft, b)).toFixed(2)}px`;
  }
  function roomBeside(t) {
    const host = t.parentElement;
    if (!host) return true;
    const hs = getComputedStyle(host), ts = getComputedStyle(t);
    const inner = host.clientWidth - (parseFloat(hs.paddingLeft) || 0) - (parseFloat(hs.paddingRight) || 0);
    const used = rectOf(t).width + (parseFloat(ts.marginLeft) || 0) + (parseFloat(ts.marginRight) || 0);
    return inner - used >= ROOM_BESIDE;
  }
  // A floating block's own margin-top (px), as it is without a push of ours: read while it has none (the push is the
  // margin while it is there, editor.css) and kept by node; one pushed before it was ever read, from its own style.
  const ownTops = new WeakMap();
  const ownTop = (b) => {
    if (!b.el.style.getPropertyValue('--odx-push')) ownTops.set(b.node, parseFloat(getComputedStyle(b.el).marginTop) || 0);
    else if (!ownTops.has(b.node)) ownTops.set(b.node, parseFloat(b.el.style.marginTop) || 0);
    return ownTops.get(b.node);
  };
  function applyPushes(doc, dom, blocks, base, colRegs, starts, pulls, splits, paraSplits, blockAttrs) {
    // Pushed blocks: margin-top = natural gap + delta. The natural gap is
    // measured between this block's top and the previous block's bottom.
    // Set on the blocks themselves (node decorations), not by rules on their place among their siblings
    // (:nth-child): with those, a paragraph added (Enter) made the browser restyle every block after it, 130-190 ms a
    // keystroke in a 443-page document.
    const posOf = new Map();
    doc.forEach((node, pos, i) => posOf.set(i, [pos, pos + node.nodeSize]));
    const blockDecos = [];
    const onBlock = (i, attrs, key) => { const r = posOf.get(i); if (r) blockDecos.push(Decoration.node(r[0], r[1], attrs, { key })); };
    doc.forEach((node, pos, i) => { if (unfloated.has(node)) onBlock(i, { style: `float:none !important${inFlow.has(node) ? `;margin-bottom:${inFlow.get(node)} !important` : ''}` }, 'nf'); });
    // (Columns' moves, cuts and widths, and frames' drops: placeColumns, the walk's frameDy.)
    for (const [i, a] of blockAttrs) {
      const style = a.style.join(';'), cls = a.cls.join(' ');
      onBlock(i, cls ? { class: cls, style } : { style }, `ca:${cls}:${style}`);
    }
    const css = [];
    applied = new Map();
    const index = new Map(blocks.map((b, k) => [b.i, k]));
    const extra = new Map(starts);
    for (const list of pulls.values()) for (const [i, d] of list) addTo(extra, i, d);
    for (const [i, delta] of extra) {
      const k = index.get(i);
      // (From the block before it in the flow: a margin doesn't count from a floating table's bottom, the text beside it
      // starts at its top. A paragraph after one, starting a page, went to 342px above the page: everything after it lost.)
      let q = k - 1;
      while (q >= 0 && outOfFlow(blocks[q].el)) q--;
      const gap = q >= 0 ? blocks[k].top - blocks[q].bottom : blocks[k].top - asLaidOut(parseFloat(getComputedStyle(dom).paddingTop) || 0); // the page's top padding is the first section's plain top margin, not one moved down for a tall header (as laid out, on the browser's grid: the 0.014px left over was added to the block's push at every pass, and it never settled)
      // A floating block's margin doesn't take the place of the gap above it, as a block's in the flow does (its margin
      // collapsing with the one above): the float stays under the space after the block before it, and under its own drop,
      // and its margin moves it from there by all it adds to the float's own. Set to gap + delta it stood a gap too low,
      // and remembered as moved by delta only, it was measured a gap lower at each pass: the paragraph after a table moved
      // on to the next page went down by that gap every time the pages were laid out (13.9px a pass in one document,
      // 2.7px in another), so a print came out by how many passes had run before it.
      const floats = outOfFlow(blocks[k].el);
      const from = floats ? ownTop(blocks[k]) : gap;
      const margin = toGrid(from + delta);
      applied.set(i, { extra: margin - from });
      // Under the bottom margin of the block before, a smaller margin collapses into it: give the difference as a negative
      // one. (After a balanced columns section, and wherever a block in the flow moves up: a table pulled back by 141.8px
      // under a paragraph with 4px after it stood at -137.8, was remembered at -141.8, and measured 4px lower at each pass,
      // so a print came out by how many passes had run before it.)
      const above = colRegs.length ? blocks[k - 1] : !floats && margin < gap ? blocks[q] : undefined;
      const mb = above ? asLaidOut(parseFloat(getComputedStyle(above.el).marginBottom) || 0) : 0;
      const set = margin < mb ? margin - mb : margin;
      // (As a property of its own, which editor.css makes the margin: a decoration's style is taken off name by name
      // when it goes, and one that set margin-top took the paragraph's own space before with it. A paragraph pushed to
      // the next page in one pass and back in a later one stood 14px high, its 10pt of space before gone.)
      onBlock(i, { style: `--odx-push:${set}px` }, `mt:${set}`);
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
        const dd = toGrid(d);
        rows.set(j, dd);
        bands.push([rel + 1, rel + dd]);
        extra += dd;
        // (This table's own rows: the same row of a table in one of its cells took the room too, six of them 3,950px in one
        // document, and the row measured that much taller, moved the table on, and came back at the next pass.)
        css.push(`${sel} tbody:not(td tbody) > tr:nth-child(${j + 1}) > td{padding-top:calc(var(--odx-pt, 0px) + ${dd}px) !important;--odx-trh-add:${dd}px}`);
        // (On top of the height a merged cell's extra gave the row, rowfit.js. Inside it, the push used that height up: the
        // row measured 17px shorter while pushed, came back, and the layout flipped between the two at every pass.)
        if (b.table?.tBodies[0]?.rows[j]?.style.getPropertyValue('--odx-rowfit')) css.push(`${sel} tbody:not(td tbody) > tr:nth-child(${j + 1}){height:calc(var(--odx-rowfit) + ${dd}px) !important}`);
      }
      for (const band of split.bands) bands.push(band);
      bands.sort((x, y) => x[0] - y[0]);
      const cut = bands.map(([a, z]) => `V${a.toFixed(1)} H${-CLIP_TOP} V${z.toFixed(1)} H${CLIP_FAR}`).join(' ');
      css.push(`${sel} table:not(td table){clip-path:path('M${-CLIP_TOP} ${-CLIP_TOP} H${CLIP_FAR} ${cut} V${CLIP_DEEP} H${-CLIP_TOP} Z')}`);
      const a = applied.get(i) || { extra: 0 };
      a.rows = rows;
      a.inRow = split.inRow;
      applied.set(i, a);
      if (b.headerRows) for (const t of split.heads) heads.push({ b, top: t });
      for (const t of split.closes) heads.push({ b, top: t, rule: true });
    }
    const decos = [...blockDecos];
    // A spacer goes with the text before it, unless that is tracked deletion: hidden in simple markup, it took the spacer
    // with it. The spacer added no height but was still taken off when measuring, so the paragraph read as shorter than
    // nothing, fitted, split again at the next pass, and the layout flipped between the two.
    const pushWidget = (pos, delta, end = false) => {
      let side = -1;
      try { if (doc.resolve(pos).nodeBefore?.marks.some((m) => m.type === schema.marks.del)) side = 1; } catch { /* a position the document no longer has */ }
      return Decoration.widget(pos, spacer(delta, end), { side, key: `push:${delta}:${side}${end ? ':end' : ''}`, ignoreSelection: true });
    };
    for (const split of splits.values()) {
      for (const { pos, delta, end } of split.cellPushes) decos.push(pushWidget(pos, delta, end));
      // (A split row's centred and bottom-aligned cells: from the top, their content placed by a spacer, splitRow.)
      for (const pos of split.cellTops || []) {
        const cell = pos >= 0 && pos < view.state.doc.content.size ? view.state.doc.nodeAt(pos) : null;
        if (cell && /^table_(cell|header)$/.test(cell.type.name)) decos.push(Decoration.node(pos, pos + cell.nodeSize, { class: 'odx-vtop' }, { key: 'vtop' }));
      }
      // A list item moved whole to the row's next page: its label, the paragraph's ::before, came before the spacer and
      // kept a line of its own above it, the item's text a line low on the new page (Word's prints: the item's first
      // line at the top of the cell's part there, bullet or not). The paragraph is told how far it's pushed, and takes
      // that as room above its label instead (editor.css, --odx-pushed).
      for (const { pos, delta, end } of split.cellPushes) {
        if (end) continue;
        let at = null;
        try { at = view.state.doc.resolve(pos); } catch { /* a position the document no longer has */ }
        if (at && at.parentOffset === 0 && at.parent.type.name === 'paragraph' && at.depth > 0) {
          // (Its own inline padding, which the room is added to: the rule above it in editor.css overrides it.)
          const own = view.nodeDOM(at.before())?.style?.paddingTop || '0px';
          decos.push(Decoration.node(at.before(), at.after(), { style: `--odx-pushed:${delta}px;--odx-pushpad:${own}` }, { key: `pushed:${delta}:${own}` }));
        }
      }
    }
    for (const [i, split] of paraSplits) {
      const a = applied.get(i) || { extra: 0 };
      a.inline = split.total;
      applied.set(i, a);
      for (const { pos, delta } of split.pushes) decos.push(pushWidget(pos, delta));
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
    // (With the rules: a split table's rows are pushed by those, so a pass that only moved a table's split changed them alone.)
    const state = `${pushSig}|${css.join('')}`;
    if (state !== pushState) pushesMoved = true;
    pushState = state;
    pushesKnown = true;
    const now = pushKey.getState(view.state)?.find().map((d) => `${d.from}:${d.to}:${d.spec.key}`).sort().join(',');
    if (pushSig !== now) {
      view.dispatch(view.state.tr.setMeta(pushKey, decos).setMeta('addToHistory', false));
      pushesMoved = true;
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
  let passed = false; // the last layout() got past its early returns, and left the pushes as it found them
  // Whether this pass changed the pushes, and how many passes in a row have for this document. A spacer can change the lines
  // it was measured from (a paragraph one line longer with it), so a pass that moved them is measured again; left at
  // that, the pages stayed as the first pass put them or as the second did, by whether anything else asked for a layout,
  // and a second print of the same page came out differently from the first.
  let pushesMoved = false;
  let settleRuns = 0;
  let settleDoc = null;
  const SETTLE_MAX = 24;
  // Pushes that come back to a set they've had: a floating table moved on to the next page takes the text beside it along,
  // which brings the table back, and both layouts are good ones. The same one of the round is kept, whichever pass got
  // there first, until the document, fonts, width or page gap change: left going, a print took whichever it found.
  let pushState = '';
  let pushRound = null;
  let fontArrivals = 0;
  const roundKey = () => `${fontArrivals}|${gap}|${view.dom.clientWidth}`;
  const roundHeld = () => !!pushRound?.done && pushRound.doc === view.state.doc && pushRound.key === roundKey();
  // The layout measures the pages at 100% whatever the zoom (zoom.js): every length it reads and sets is the page's own.
  const unzoomed = (fn) => {
    const z = sheet.style.zoom;
    if (!z) return fn();
    sheet.style.zoom = '';
    try { return fn(); } finally { sheet.style.zoom = z; }
  };
  // The wrap step's pass. A header's or footer's shape stays on its page while the text moves, so until the pages' pushes
  // are in, the text isn't where it will be beside it: lines split round a shape they don't reach on their own pages, and a
  // page's spacer put in among those lines held them short once they'd gone (20 characters short in a 41-page document, a
  // line more in the paragraph, and every page after it out). A shape in the body moves with its text, and counts from the start.
  let pushesKnown = false;
  const wrapStep = () => applyWrapBoth(view, pushesKnown ? layer : null, pushState);
  // (Every pass, early ones too: a drawing placed on its page by a pass that went on to stop early had kept the cut of
  // where it was before.)
  function layout() { return unzoomed(() => { const r = layoutPass(); quiet(clipToSheets); return r; }); }
  function layoutPass() {
    lastPass = performance.now();
    lineSplits = splitsNow();
    frame = 0;
    passed = false;
    if (!colsRepass) colsPasses = 0; // (a pass asked for from outside: the column widths may settle afresh)
    colsRepass = false;
    if (shapesDoc !== view.state.doc) { shapesDoc = view.state.doc; shapesPasses = 0; }
    const t0 = performance.now();
    quiet(() => layoutTabs(view, session));
    if (quiet(() => fitAutofitColumns(view))) return schedule(); // column widths changed: measure again once they're drawn
    if (applyLineFix(view)) return schedule(); // line heights changed: measure again once they're drawn
    if (applyWidths(view)) return schedule(); // hinted fonts' widths corrected: measure again once they're drawn
    if (applyFitText(view)) return schedule(); // fitted text spaced out to its width: measure again
    if (applyLabelWords(view)) return schedule(); // a list paragraph's overlong first word broken on the label's line: measure again
    // (And round the shapes wrapped round in the headers and footers drawn on the sheets: wrapboth.js storyZones.)
    if (wrapStep()) return schedule(); // a line split beside a shape with text on both sides: measure again
    if (applyTableClear(view)) return schedule(); // a table that doesn't fit beside a float moved under it: measure again
    // (Its own pass: a view update only lays out again when the document changed, and a squeeze that settled after
    // everything else (lines split beside a shape, wrapboth.js) had left the pages undrawn.)
    if (applySqueeze(view, session)) return schedule(); // lines changed: measure again once they're drawn
    if (quiet(() => fitRowspans(view.dom))) return schedule(); // merged cells' extra height moved to their last row: measure again
    passed = true;
    if (roundHeld()) return;
    const dom = view.dom;
    const doc = view.state.doc;
    // Word lets a floating table longer than a page run on in the flow: nothing wraps beside it.
    // Kept as a rule too: redrawing the table's decorations would reset an inline style.
    // (Each block's own element, as measureBlocks finds it: the DOM's children needn't line up with the document's.)
    const topEls = topElements(doc);
    doc.forEach((node, pos, i) => {
      const t = topEls ? topEls[i] : view.nodeDOM(pos);
      if (!t || t.nodeType !== 1) return;
      if (getComputedStyle(t).float === 'none') return;
      // (Not one placed down from the text, setDown, or with no room beside it, styles.js --odx-fnoroom: it stays a
      // float, so the text after it fills the room above it, as in Word, and its empty paragraphs go under it, editor.js.)
      const above = node.type === schema.nodes.table && /--odx-fnoroom/.test(t.getAttribute('style') || '');
      if (rectOf(t).height > pageHeight(0)) { quiet(() => { t.style.float = 'none'; }); unfloated.add(node); }
      else if (node.type === schema.nodes.table && !above && !(setDown(node) && !crossFoot.has(node)) && !roomBeside(t)) {
        const mb = inFlowGap(node, doc.maybeChild(i + 1));
        quiet(() => { t.style.float = 'none'; t.style.marginBottom = mb; });
        unfloated.add(node);
        inFlow.set(node, mb);
      }
    });
    const colRegs = colRegions(doc);
    const inRegion = (i) => colRegs.some((r) => i >= r.lo && i <= r.hi);
    if (cols || colRegs.length) dom.classList.add('odx-measure'); // measure the flow without the columns' moves
    sheet.classList.add('odx-bl-off'); // and without the print-only baseline shifts
    const base = rectOf(dom).top;
    const blocks = measureBlocks(dom, doc, base, inRegion);
    // Walk the blocks and decide where pages start. `shift` is how far the
    // blocks already pushed have moved everything below them.
    const starts = new Map(); // block index -> extra px pushed (applied as margin-top)
    const carried = new Set(); // blocks after a frame moved to the next page: their push is the frame's
    const tails = new Map(); // block index -> the height of the line before it that Word starts its page with
    const splits = new Map(); // table block index -> { rows: Map(row -> extra px), heads: [page top] }
    let crossAgain = false; // a float set down found reaching past its page's foot (crossFoot): the pass runs again
    const paraSplits = new Map(); // paragraph block index -> { pushes: [{ pos, delta }], total }
    const colSplits = new Map(); // (columns) paragraph block index -> { cut, height }: its lines from cut go on in the next column
    // (Columns: how often a paragraph moved on to a column it's taller than has been walked again, to be cut there. A
    // bound: one that can't be cut there either moved on and was walked again for good, and the page never drew.)
    const recut = new Map();
    const RECUT_MAX = 8;
    // Tables set down at a place on their page (styles.js: --odx-fy) and the blocks moved below them: block index ->
    // [[block index, px]...] added to those blocks' margin-top; taken back off when the walk comes past the block again.
    const pulls = new Map();
    let band = null; // { p, y0, y1 }: where such a table sits, which text after it goes above or below
    let shift = 0;
    // What the page's blocks so far, this one included, take in the browser more than they do exactly: its lines'
    // heights and the room between paragraphs are whole 64ths of a pixel there, to the nearest, and 50 lines of one
    // height add up to a third of a pixel either way. Word fits a page's last line exactly: in its PDFs of 189
    // generated documents (Arial, Times New Roman and Calibri, single, 1.15 and exact spacing), a page of one-line
    // paragraphs holds its last line when that ends on the margin and loses it one twip (0.07px) past.
    let roundDrift = 0;
    let roundOwn = 0;
    roundOver = 0;
    let breakPending = false;
    let breakKind = null;
    if (sectCache.doc !== doc) sectCache = { doc, ...sectionBreaks(doc, session, page, storyHeight) };
    const sectEnds = sectCache.ends;
    margins = [[0, sectCache.first, true]];
    sections = [[0, 0]];
    notesOn = new Map();
    notesH = new Map();
    noteFlow = new Map();
    sumSteps = [];
    sumFrom = -Infinity;
    cuts = new Map();
    // A one-section document in columns with endnotes: its last sheet's columns cut to an even height (worked out a pass
    // behind, balanceEndCols), as Word balances them before the endnotes.
    if (cols && colsBal) for (let c = 0; c < cols.num - 1; c++) { const v = colsBal.sheet * cols.num + c; cuts.set(v, contentTop(v) + colsBal.h); }
    // (And a sheet whose footnotes are too tall for it: its columns cut where its text up to the reference evens out.)
    for (const hp of heldPrev.values()) for (const [v, y] of hp.cuts) cuts.set(v, Math.min(cuts.get(v) ?? Infinity, y));
    colTops = new Map();
    for (const r of colRegs) {
      r.k0 = blocks.findIndex((x) => x.i >= r.lo);
      r.k1 = blocks.findIndex((x) => x.i > r.hi);
      if (r.k1 < 0) r.k1 = blocks.length;
    }
    const regs = colRegs.filter((r) => r.k0 >= 0 && r.k0 < r.k1);
    curRegs = regs;
    if (balancedDoc !== doc) { balancedDoc = doc; balancePasses = 0; }
    if (crossDoc !== doc) { crossDoc = doc; crossPasses = 0; }
    let balanceAgain = false;
    const rec = new Map(); // (columns sections) block k -> { top, bottom, carry } where it was last placed
    const frameDy = new Map(); // block index -> px a frame set down on its page is moved below where it comes
    // Block index -> px of its space before left above it at the top of the page it was moved to (the space kept: none at
    // a page Word starts by itself): a drawing placed from its paragraph's top above that space is placed from the page's
    // top there, not above it (reader.js fromParagraphTop, --odx-sbtop).
    const sbTop = new Map();
    floatUp.clear();
    shapeUp.clear();
    tbUp.clear();
    if (tbDoc !== doc) { tbDoc = doc; tbPasses = 0; }
    placed.clear();
    assignNotes(doc);
    // The steps of the walk below that have names of their own (closures: they move `shift` and the walk's maps).
    // What follows a block split into the next column moves down by the split's delta.
    const carryOn = (k, cs) => { const nb = blocks[k + 1]; if (nb) { addTo(starts, nb.i, cs.delta); shift += cs.delta; } };
    // A paragraph in a frame (frames.js), out of the flow.
    const placeFrame = (k, b, top, p, frame) => {
      if (frame === 'page' || frame === 'margin') {
        const now = parseFloat(getComputedStyle(b.el).getPropertyValue('--odx-fdy')) || 0;
        const natural = top - now;
        const pg = pageOf(natural);
        const base = frame === 'page' ? pageTop(pg) : contentTop(pg);
        const room = frame === 'page' ? pageH(pg) : contentBottom(pg) - contentTop(pg);
        const h = b.bottom - b.top;
        const al = b.el.dataset.fyalign;
        const y = alignedY(al, base, base + room, h, Number(b.el.dataset.fy || 0));
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
          addTo(starts, b.i, d);
          const nb = blocks[k2 + 1];
          if (nb) { addTo(starts, nb.i, d); carried.add(nb.i); } // (the text after it: a float's margin doesn't move it)
          shift += d;
        }
      }
    };
    // Whether a set-down table leaves room beside it in its column (ROOM_BESIDE, as roomBeside), where Word sets text: the rules
    // for text before such a table (placeSetDown, belowAhead) are for one leaving none. (Against page p's own text
    // width: the editor's text is as wide as the widest page, a landscape section's.)
    const setDownRoom = (blk, p) => {
      const t = tableOf(blk.el);
      if (!t) return false;
      const m = marginsAt(p);
      const inner = pageW(p) - (m.left ?? page.left) - (m.right ?? page.right);
      return inner - rectOf(t).width >= ROOM_BESIDE;
    };
    // Whether set-down table k, coming at `top` on page p, goes on to the next page (placeSetDown).
    const movesOn = (k, top, p, fy) => {
      const b = blocks[k];
      const y0 = (fy.rel === 'page' ? pageTop(p) : contentTop(p)) + fy.y;
      return !taken.has(k) && !setDownRoom(b, p) && top > y0 + (b.bottom - b.top) + fy.below + 0.5 && y0 + (b.bottom - b.top) > contentTop(p) + 1 && !cols && !regs.some((r) => k >= r.k0 && k < r.k1);
    };
    // Whether paragraph k, on page p, keeps with (keep with next, through other such paragraphs) a set-down table that
    // goes on to the next page, and with the paragraph after that table: it goes with them, under the table there
    // (belowAhead), as Word sets it (a sample's cost table past its place at a page's foot: its caption, kept with it,
    // under it on the next page; the paragraph before that, not kept with it, stayed).
    const keepsWithMoving = (k, p) => {
      let g = 0;
      for (let m = k; m < blocks.length - 1 && g < 6; m++, g++) {
        if (!keep(blocks[m].node)) return false;
        const nx = blocks[m + 1];
        const fy = nx.rows ? fyOf(nx) : null;
        if (fy) {
          if (!movesOn(m + 1, nx.top + shift, p, fy)) return false;
          // (Only when the paragraph after the table, its anchor, doesn't fit on the page either, where it comes:
          // Word's print, fltrow-word b1, keeps a caption with room under it on its page, the table and the text after it
          // on the next; a sample's caption at the page's foot went on with them.)
          const an = blocks[m + 2];
          if (!an) return true;
          const al = an.node.type === schema.nodes.paragraph ? paragraphLines(an) : null;
          return nx.top + shift + ((al ? al[0].bottom : an.top) - nx.bottom) > contentBottom(p) + 0.5;
        }
        if (nx.node.type !== schema.nodes.paragraph) return false;
      }
      return false;
    };
    // A floating table longer than a page, laid out in the flow and split from page p, with room beside it: the empty
    // paragraphs after it go beside its last part, from that page's top, as a float's text does, and what follows them
    // from there (Word's prints, why-tests flt0898: 2 empty paragraphs beside a 70-row table's last 25 rows, the text
    // after them at its third line; a sample's five, a floating table placed down from the last of them 117px low
    // there, and every page after it out).
    const besideLast = (k, p) => {
      const b = blocks[k], nb = blocks[k + 1];
      const last = pageOf(b.bottom + shift - 1);
      if (!nb || last <= p || !floating(b.node) || !unfloated.has(b.node) || !roomBeside(b.el)) return;
      if (nb.node.type !== schema.nodes.paragraph || nb.node.content.size) return;
      const d = contentTop(last) + spaceBefore(nb.node) - (nb.top + shift);
      if (d < 0) { addTo(starts, nb.i, d); shift += d; }
    };
    // A table set down at its place on its page (floatAt).
    const placeSetDown = (k, b, top, p, fy) => {
      let y0 = (fy.rel === 'page' ? pageTop(p) : contentTop(p)) + fy.y;
      // Its place on its page above where it comes, the text before it filling that place (below the page's top): on
      // the next page, at its place there, and the text after it from that page's top, as Word sets it (fltrow-word
      // b1: a table set down 8905tw from the top after a page of text reaching past that, on page 2, the paragraph after
      // it starting page 2; a sample's deliverables table after text at its page's foot, on the next page). Not when the
      // text before it went under it (belowAhead), nor for one with room beside it (setDownRoom).
      let from = top;
      if (movesOn(k, top, p, fy)) {
        p += 1;
        y0 = (fy.rel === 'page' ? pageTop(p) : contentTop(p)) + fy.y;
        from = contentTop(p);
      }
      const nb = blocks[k + 1];
      const d1 = y0 - top;
      const d2 = nb ? from + (nb.top - b.bottom) - (nb.top + shift + d1) : 0;
      pulls.set(b.i, nb ? [[b.i, d1], [nb.i, d2]] : [[b.i, d1]]);
      shift += d1 + d2;
      band = { k, p, y0: y0 - fy.above, y1: y0 + (b.bottom - b.top) + fy.below };
      // With less than two lines' room under it, the page's text ends above it: a table reaching into it splits there
      // and goes on on the next page (the contingency plan's page 4), rather than all of it moving under the band.
      if (band.y1 > contentBottom(p) - MIN_ROOM) { cuts.set(p, band.y0); band.cut = true; }
    };
    // The band of a table set down on page p that comes after block k with nothing between taking the text to another
    // page: { y0, y1 }, or null. (Outside newspaper columns; looked for within a page's worth of blocks.)
    const fyMemo = new Map();
    const refused = new Set(); // set-down tables (block index) the text before them on their page stays above
    const taken = new Set(); // and those it went under
    const fyOf = (blk) => { if (!fyMemo.has(blk.i)) fyMemo.set(blk.i, floatAt(blk)); return fyMemo.get(blk.i); };
    const bandAhead = (k, p) => {
      if (cols || regs.some((r) => k >= r.k0 && k < r.k1)) return null;
      for (let m = k + 1; m < blocks.length && m <= k + 60; m++) {
        const bm = blocks[m];
        if (pageOf(bm.top + shift) !== p || breakBefore(bm.node) || bm.el.querySelector?.('.odx-page-break')) return null;
        const fy = fyOf(bm);
        if (fy) {
          // (Only one leaving no room beside it, as Word's: with room the text goes beside it. A sample's set-down table
          // half the column wide, the text before it beside it in Word's print.)
          if (setDownRoom(bm, p)) return null;
          const y0 = (fy.rel === 'page' ? pageTop(p) : contentTop(p)) + fy.y;
          return { m, y0: y0 - fy.above, y1: y0 + (bm.bottom - bm.top) + fy.below };
        }
      }
      return null;
    };
    // How far paragraph k, at `top` on page p, moves to go below such a table it reaches into (the move noted in pulls);
    // 0 when it doesn't.
    const belowAhead = (k, top, p) => {
      const b = blocks[k];
      if (b.node.type !== schema.nodes.paragraph) return 0;
      const ab = bandAhead(k, p);
      if (!ab || top >= ab.y1 || b.bottom + (top - b.top) <= ab.y0) return 0;
      if (refused.has(ab.m)) return 0;
      const d = ab.y1 + (k > 0 ? Math.max(0, spaceBefore(b.node) - spaceAfter(blocks[k - 1].node)) : 0) - top;
      if (d <= 0) return 0;
      // (Only while the table and the first line of the paragraph after it, which it is anchored to, still come on this
      // page with the text moved: else the text stays, and the table goes on to the next page, placeSetDown. Word's
      // print, fltrow-word b1: a page of text under the table's place kept there, the table on page 2.)
      const tb = blocks[ab.m], nb = blocks[ab.m + 1];
      const nl = nb && nb.node.type === schema.nodes.paragraph ? paragraphLines(nb) : null;
      // (The paragraph after it starts where the table comes in the text, placeSetDown.)
      const anchorEnd = nb ? tb.top + ((nl ? nl[0].bottom : nb.top) - tb.bottom) : tb.top;
      if (anchorEnd + shift + d > contentBottom(p) + 0.5) { refused.add(ab.m); return 0; }
      pulls.set(b.i, [...(pulls.get(b.i) || []), [b.i, d]]);
      taken.add(ab.m);
      return d;
    };
    // A picture or text box floating from the paragraph (wrapped square) past the page's foot (floatsPast): taller than
    // the page, it goes to the next one with its paragraph (true); else drawn up to end at the page's edge (floatUp).
    const floatsMoveOn = (b, p) => {
      const f = floatsPast(b, shift, p);
      if (f.up?.size) floatUp.set(b.i, { el: b.el, up: f.up });
      return f.past;
    };
    // A page or column break inside a paragraph in newspaper columns: its lines from the break go on at the top of the
    // next page's first column, or of the next column, as in Word (its text after a page break had stayed in the
    // column, and what followed went on in the next column of the same page). True when it was split there.
    const splitAtBreak = (k, b, p, top, inReg) => {
      const lines = paragraphLines(b);
      const at = lines ? lines.findIndex((l, n) => n > 0 && l.forced) : -1;
      if (at <= 0) return false;
      const g = cols ? { P: 0, num: cols.num } : inReg;
      const to = breakKindOf(b.node) === 'page' ? g.num - ((p - (g.P ?? 0)) % g.num) : 1;
      const cs = { cut: lines[at].top - b.top, height: b.bottom - b.top, delta: contentTop(p + to) - (lines[at].top + shift), to };
      colSplits.set(b.i, cs);
      notesByCol(b, p, cs);
      if (inReg) rec.set(k, { top, bottom: top + cs.cut, carry: cs.height - cs.cut });
      carryOn(k, cs);
      breakPending = false;
      return true;
    };
    // Whether the blocks before block k that keep with it (keep with next) reach back to page p's top: moving them on
    // with the table's first rows gains nothing, and Word breaks the table where it runs out of room instead (generated
    // documents: a heading kept with a table of 70 rows each keeping with the next, on the next page with its first 55
    // rows; it had been left there alone, the table starting the page after).
    // Whether a block keeps with the one after it: a paragraph set to; a table whose rows all do (a paragraph in each
    // first cell keeps with the next), whole on page p, as Word moves such a table on with the paragraph after it that
    // doesn't fit (generated documents: five rows ending 5px above the foot, every paragraph in them kept with the next,
    // on page 2 under the empty paragraph after it; without the keeps, on page 1).
    const keepsOn = (blk, p) => {
      if (blk.node.type !== schema.nodes.table) return keep(blk.node);
      if (!blk.node.childCount || splits.has(blk.i) || floating(blk.node)) return false;
      for (let j = 0; j < blk.node.childCount; j++) if (!rowKeeps(blk.node.child(j))) return false;
      return pageOf(blk.top + shift) === p;
    };
    const keptFromTop = (k, p) => {
      for (let m = k - 1; m >= 0 && keepsOn(blocks[m], p); m--) {
        const t = blocks[m].top + shift;
        if (pageOf(t) !== p) return false;
        if (t <= contentTop(p) + 1) return true;
      }
      return false;
    };
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
        placeFrame(k, b, top, p, frame);
        continue;
      }
      // A table set down on its page: there, out of the flow. What follows it starts where it would have (the page above
      // it fills first), and a block reaching into its band goes below it, as in Word.
      const fy = floatAt(b);
      if (fy) {
        placeSetDown(k, b, top, p, fy);
        continue;
      }
      if (band && !band.cut && band.p === p && top < band.y1 && b.bottom + shift > band.y0) {
        // (Under the table, its space before less the space after the paragraph above it, as in Word: 24pt kept under
        // none, none under a larger one.)
        const d = band.y1 + (k > 0 ? Math.max(0, spaceBefore(b.node) - spaceAfter(blocks[k - 1].node)) : 0) - top;
        pulls.set(b.i, [[b.i, d]]);
        shift += d;
        top += d;
        p = pageOf(top);
      } else if (!band || band.p !== p) {
        // Text before such a table on its page goes below it too where it reaches into its band, as in Word (a sample's
        // cost table set down 8905tw from the page's top, after a paragraph and its caption in the file: both under it in
        // Word's print; ours ran them under the table, the page held more, and the next set-down table came a page early).
        const d = belowAhead(k, top, p);
        if (d) { shift += d; top += d; p = pageOf(top); }
      }
      // The end of a columns section (endColumns).
      const re = regs.find((r) => r.k1 === k && r.P != null);
      if (re) {
        const e = endColumns(re, k, b, top, blocks, rec, sectEnds, starts);
        if (e.moved != null) { shift += e.moved; top += e.moved; p = pageOf(top); }
        if (e.again) balanceAgain = true;
      }
      // (Below it, its space after and the note separator's line as Word draws it, when a note has to start under it: its first
      // line 37.6pt under the text's last baseline in Word's print, longfn-word fnhuge.)
      reserveNotes(b.i, p, b.bottom + shift + spaceAfter(b.node) + 11.5);
      // A new section's margins: from this page if it starts one, else from the next.
      while (margins.length > 1 && margins[margins.length - 1][0] > p + 1) margins.pop(); // left by an earlier pass over these blocks
      // (The section that ends between the block before and this one: its paragraph may have no box, collapsed into the
      // columns that follow it, editor.js.)
      let sect = null;
      if (k > 0) for (let j = b.i - 1; j >= blocks[k - 1].i && !sect; j--) sect = sectEnds.get(j) || null;
      if (sect) {
        // (A continuous section whose text starts a page all the same, a page break before it: that page is the
        // section's own from its top, its first page, with the first-page header and footer where it has them. Word's
        // prints of 18 generated documents: a section with a different first page starting part way down page 1
        // shows the ordinary header on page 2, whatever its margins, columns or the sections before it; with its first
        // paragraph set to break before, or a page break last before the section's end mark, page 2 has the
        // first-page header. A sample's page 2 stood 27px low under the ordinary header, and ran a page over.)
        // (At a page's top already, or about to be taken to the next page's by the break, below.)
        const atTop = top <= contentTop(p) + 1;
        const broken = !atTop && k > 0 && ((breakPending && breakKind !== 'column') || breakBefore(b.node));
        const from = atTop ? p : p + 1;
        setMargins(from, sect.next, sect.page || ((atTop || broken) && from > 0));
        while (sections.length > 1 && sections[sections.length - 1][0] >= from) sections.pop();
        sections.push([from, sect.index]);
      }
      const firstOnPage = top <= contentTop(p) + 1;
      roundDrift = firstOnPage ? 0 : roundDrift + roundOwn;
      roundOwn = roundOff(b, firstOnPage || k === 0 ? null : blocks[k - 1]);
      roundOver = Math.round((roundDrift + roundOwn) * 1e6) / 1e6; // (a tie on the margin isn't lost to floating-point dust)
      // A columns section: the first column starts where the section does (on a new page, at its top), the others beside
      // it at that height (flow pages P+1.. of the same sheet), then whole sheets of columns. Balanced, the columns of
      // its last sheet (as the last pass found it) end at the height worked out for them, all but the last.
      const rs = regs.find((r) => r.k0 === k);
      if (rs) startColumns(rs, regs, sect, firstOnPage, p, top);
      const inReg = regs.find((r) => k >= r.k0 && k < r.k1);
      const inCols = cols || !!inReg;
      if (inReg) rec.set(k, { top, bottom: b.bottom + shift });
      // A table placed up from its text (upFromText), in the flow (no room beside it), on its page under a paragraph of
      // one line: drawn up, it reaches into that line, which goes under it, and the paragraph after it under that, as
      // Word sets them (why-tests flt0898 op-*, th-*: tables 1 to 200tw up under a line of text or an empty one, that line
      // under the table and the anchor's next; a sample's table 182tw up, the text after it 12px high).
      if (b.rows && !firstOnPage && !inCols && k > 0 && upFromText(b.node) && !outOfFlow(b.el) && blocks[k - 1].node.type === schema.nodes.paragraph) {
        const pv = blocks[k - 1], nb = blocks[k + 1];
        const pl = paragraphLines(pv);
        const h = b.bottom - b.top, l = pv.bottom - pv.top;
        const g = nb ? spaceAfter(pv.node) + spaceBefore(nb.node) : 0; // (both, sp-a160-b*-ym200)
        // (Only when it reaches past that line's space after: 39tw up under a line with 8pt after, nothing moves, sp-a160-*.)
        if (-upLift(b.node) > spaceAfter(pv.node) + 0.01 && (pl ? pl.length === 1 : !pv.node.content.size) && pageOf(pv.top + shift) === p && top + h + l + g <= contentBottom(p) && (!nb || nb.node.type === schema.nodes.paragraph)) {
          const d = top + h - (pv.top + shift);
          addTo(starts, pv.i, d);
          addTo(starts, b.i, -d);
          if (nb) { addTo(starts, nb.i, l + g - (nb.top - b.bottom)); carried.add(nb.i); } // (kept if it then moves on)
          shift += nb ? l + g - (nb.top - b.bottom) : 0;
        }
      }
      // An empty paragraph that only carries a section break may hang into the bottom margin: Word doesn't start a page for it.
      // Nor after a manual page break: the page break and the section break then make one new page, not two.
      // (Less the space after the paragraph above: Word drops it at the page's foot, so a page filled exactly keeps its
      // section's empty last paragraph.)
      // (Not after a column break: that paragraph goes to the next column, and a next-page section starts the page after.)
      // (From above its own space before too, the larger of the two: Word's prints of a 14pt one, 5.9px before, after a
      // table ending 20px into its room, kept on the page; mksect s-p3..s-p20.)
      if (!firstOnPage && lineless(b.node) && sectEnds.get(b.i) && !(breakPending && breakKind === 'column') && top - Math.max(spaceBefore(b.node), k > 0 ? spaceAfter(blocks[k - 1].node) : 0) < contentBottom(p)) { if (sectEnds.get(b.i).page) breakPending = false; continue; }
      // (In columns of their own widths, a paragraph's bottom at the width of the column it's in: atColumnWidth.)
      const wide = inCols ? atColumnWidth(b, cols || inReg, p, shift, top <= contentTop(p) + 1) : null;
      const bottom = wide ? wide.bottom : b.bottom;
      // (Its shapes with text on both sides past the foot, drawn up: shapesPast. Not in columns.)
      let shapePast = false;
      if (!inCols && b.node.type === schema.nodes.paragraph) {
        const sp = shapesPast(b, shift, p);
        if (sp.up.size) shapeUp.set(b.i, { el: b.el, up: sp.up });
        shapePast = sp.past;
        // (Its picture wrapped top and bottom drawn up to the foot, out of the flow: the page's text ends at its top,
        // tbPast. Set out of the flow only after the walk; until then its paragraph is as tall as it, and the pages are
        // walked again.)
        const tb = tbPast(b, shift, p);
        if (tb) {
          tbUp.set(b.i, tb);
          if (tbOut.has(tb.img) && bottom + shift <= tb.top + 0.5) cuts.set(p, Math.min(cuts.get(p) ?? Infinity, tb.top));
        }
      }
      let target = null;
      let overflow = false;
      let forced = false;
      let sectStart = false;
      if (k > 0 && (breakPending || breakBefore(b.node) || (sect && sect.page) || (!inCols && keepsWithMoving(k, p))) && !firstOnPage) {
        const g = cols ? { P: 0, num: cols.num } : inReg;
        target = g && (breakPending ? breakKind === 'page' : !!cols && !lineSplits) ? p + g.num - ((p - (g.P ?? 0)) % g.num) : p + 1;
        if (sect && sect.parity) target = parityPage(target, sect.parity);
        forced = !breakPending; // after a manual page break Word drops the space before, as at a natural break
        sectStart = forced && !!(sect && sect.page) && !breakBefore(b.node);
      }
      else if (k > 0 && sect && sect.parity && firstOnPage && parityPage(p, sect.parity) !== p) {
        // (Already at a page's top, but of the wrong parity: a blank page first.)
        target = parityPage(p, sect.parity);
        forced = true;
        sectStart = true;
      }
      else if (p === 0 && k === 0 && top < contentTop(p) + spaceBefore(b.node) - 0.5) {
        // Under a header that reaches past the top margin, the first paragraph keeps its space before: Word moves the body down whole.
        // (Checked first: under a header reaching further than the space before, the paragraph would otherwise lose it.)
        target = p;
        forced = true;
      }
      else if (k > 0 && !firstOnPage && b.node.type === schema.nodes.paragraph && liftedPage(b.node) > p) target = liftedPage(b.node); // (its float's page: liftFloats)
      else if (top < contentTop(p)) target = p; // landed in a gap or a top margin
      else if (!firstOnPage && b.node.type === schema.nodes.paragraph && floatsMoveOn(b, p)) {
        target = p + 1;
        overflow = true;
      }
      else if (!firstOnPage && shapePast) { // (its shape drawn up over the text above it: shapesPast)
        target = p + 1;
        overflow = true;
      }
      else if (((cols && !lineSplits) || inReg) && b.node.type === schema.nodes.paragraph && hasPageBreak(b.node) && splitAtBreak(k, b, p, top, inReg)) continue;
      else if (inCols && b.node.type === schema.nodes.paragraph && !hasPageBreak(b.node) && bottom + shift + bandAfter(p, b.node) > contentBottom(p) + roundOver + lineSlack(b.el) - 1) {
        // In columns a paragraph goes on in the next column line by line, as in Word: it stays put, cut at the column's
        // bottom, and its lines from there are drawn again at the next column's top (drawHeads); what follows it moves
        // down below them.
        // (Laid out at another column's width: split on the next pass, at this one's, when it can be; else it moves on.)
        if (wide?.splits) continue;
        const cs = wide ? null : colSplit(b, shift, firstOnPage, p);
        if (cs) {
          // Notes whose references go on in the next column leave this one's band: cut again in the room they leave (unless
          // that brings another reference back), as in Word.
          if (notesByCol(b, p, cs)) {
            const cs2 = colSplit(b, shift, firstOnPage, p);
            if (cs2 && cs2.cut > cs.cut && !refsBetween(b, cs.cut, cs2.cut)) Object.assign(cs, cs2);
          }
          if ((cols || inReg).unequal) rewrapCarry(b, cs, cols || inReg, p);
          // (Rewrapped at the next column's width, a paragraph laid out in a narrower one can leave a gap under its box far
          // taller than its space after; what follows goes under its carried lines at its own spacing, as in Word: the
          // generated documents' indented paragraphs in a 1.4in first column, the next one 53px low otherwise.)
          const nx = blocks[k + 1];
          if (cs.rewrap && nx && nx.node.type === schema.nodes.paragraph) {
            const want = Math.max(spaceAfter(b.node), spaceBefore(nx.node));
            const gap = nx.top - b.bottom;
            if (gap > want + 2) cs.delta -= gap;
          }
          colSplits.set(b.i, cs);
          notesByCol(b, p, cs);
          if (inReg) rec.set(k, { top, bottom: top + cs.cut - (cs.lift || 0), carry: (cs.rewrap ? cs.rewrap.h : cs.height - cs.cut + (cs.lift || 0)) + (cs.head || 0), parts: cs.parts ? carriedOf(cs) : null });
          carryOn(k, cs);
          continue;
        }
        // (Taller than the next column too, it still goes on there, and is cut at its foot: kept here, it ran down past the
        // column's foot, over a band of notes taking most of the page, fncols-word tall3a.)
        if (!firstOnPage && !wide?.splits) {
          target = p + 1;
          overflow = true;
        }
      }
      else if (inCols && b.node.type === schema.nodes.table && b.bottom + shift > contentBottom(p) + 1) {
        const cs = colSplitTable(b, shift, p);
        if (cs) {
          colSplits.set(b.i, cs);
          notesByCol(b, p, cs);
          if (inReg) rec.set(k, { top, bottom: top + cs.cut, carry: cs.height - cs.cut + (cs.head || 0), parts: cs.parts ? carriedOf(cs) : null });
          carryOn(k, cs);
          continue;
        }
        if (!firstOnPage) {
          target = p + 1;
          overflow = true;
        }
      }
      else if (lineSplits && !inCols && b.node.type === schema.nodes.paragraph && (hasPageBreak(b.node) || b.bottom + shift + bandAfter(p, b.node) > contentBottom(p) + roundOver + (footBand(p) ? 0 : lineSlack(b.el) - 1) - tie(b.el))) {
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
          if (split.pushes.length) { roundDrift = 0; roundOwn = 0; } // (its last lines start a page: nothing above them there)
          continue;
        }
        // (One taller than a page moves too, and splits on its new page.)
        if (split === null && !firstOnPage) {
          target = p + 1;
          overflow = true;
        }
      } else if (!(lineSplits && b.node.type === schema.nodes.paragraph && b.node.textContent.trim()) && (bottom + shift > contentBottom(p) + roundOver + (b.node.type === schema.nodes.paragraph && !b.node.content.size ? lineSlack(b.el) : 1) || (b.rows && rowBreakIn(b.node)))) {
        // (A paragraph of text that passed the line test above stays: only the extra below its text passes the margin.
        // One holding only a picture has no such extra, and moves. An empty one's line is its mark's, with that extra
        // under it as a line of text has: a sample's empty 10pt paragraph at 1.15 lines, reaching 1.05px past the
        // margin, stays at its page's foot in Word and had gone to the next page's top, everything after it 32px low.)
        // Tables continue on the next page row by row, as in Word; other blocks move whole.
        if (b.rows && bottom + shift > contentBottom(p) + 1 && setDown(b.node) && !crossFoot.has(b.node) && crossPasses < CROSS_MAX && outOfFlow(b.el) && !roomBeside(b.el)) {
          crossFoot.add(b.node); // (a float set down that reaches past the foot: in the flow, next pass)
          crossAgain = true;
        }
        // (Not one placed up from its text, movesWhole, below a page's top: it goes on to the next page, below.)
        const split = b.rows && !(!firstOnPage && bottom + shift > contentBottom(p) + 1 && movesWhole(b)) && splitTable(b, shift, firstOnPage, !firstOnPage && keptFromTop(k, p));
        if (split) {
          splits.set(b.i, split);
          shift += split.total;
          breakPending = false;
          besideLast(k, p);
          continue;
        }
        // A table whose first row can't start here moves to the next page even when it's longer than a page (it splits there).
        if (!firstOnPage && bottom + shift > contentBottom(p) + 1 && (b.rows || b.bottom - b.top <= pageHeight(p + 1))) {
          target = p + 1;
          overflow = true;
        }
      }
      breakPending = (!lineSplits || !!inReg) && hasPageBreak(b.node);
      breakKind = breakKindOf(b.node);
      if (target == null) continue;
      roundDrift = 0; // (first on its new page: only its own part counts there)
      roundOver = roundOwn;
      // Keep with next: headings at the bottom of a page move with the block they introduce.
      let m = k;
      if (overflow) {
        // (In newspaper columns, also from the top of a column that isn't its page's first, moved there already: Word
        // takes a paragraph on with the one it keeps with column by column, to the next page's first column at most.
        // The generated documents' headings in two and three columns, tools/doc-generator.)
        const midSheet = !!cols && p % cols.num !== 0;
        let cut = null;
        while (m > 0 && keepsOn(blocks[m - 1], p) && (midSheet ? blocks[m - 1].top + shift >= contentTop(p) - 1 : !starts.has(blocks[m - 1].i) && blocks[m - 1].top + shift > contentTop(p) + 1)) {
          // A paragraph of several lines that keeps with the block moving on, and may split: only its last lines go
          // with that block (two under widow control, one without), the rest staying on this page, as Word breaks it
          // (120 generated documents: 2 to 6-line paragraphs each kept with the next, widow control on and off; with
          // it a paragraph under four lines can't split and moves whole, as one whose lines keep together does).
          cut = !midSheet && lineSplits && !cols ? keepCut(blocks[m - 1]) : null;
          if (cut) break;
          m--;
        }
        // (A continuous section whose first block goes on to the next page, by itself or with the blocks it is kept
        // with: the section's text starts that page, which is its first, as above. A sample's section opening with
        // two headings kept with their text, no room for them under the section before.)
        if (!cut && m > 0) {
          let s0 = null;
          for (let j = blocks[m].i - 1; j >= blocks[m - 1].i && !s0; j--) s0 = sectEnds.get(j) || null;
          const e = s0 && !s0.page ? margins.find((x) => x[1] === s0.next) : null;
          if (e && e[0] === target) e[2] = true;
        }
        if (cut) {
          const prev = blocks[m - 1];
          const split = splitParagraph(prev, shift, false, false, cut);
          if (split && split.pushes.length) {
            paraSplits.set(prev.i, split);
            notesByLine(prev, shift, split);
            shift += split.total;
            for (let n = m; n <= k; n++) {
              sbTop.delete(blocks[n].i);
              shapeUp.delete(blocks[n].i);
              moveNotes(blocks[n].i, target);
              splits.delete(blocks[n].i);
              paraSplits.delete(blocks[n].i);
              colSplits.delete(blocks[n].i);
            }
            k = m - 1; // (everything after the split paragraph again, from its new place)
            continue;
          }
        }
        // A run of blocks each kept with the next that reaches back to the page's top is longer than the page: moving
        // part of it on gains nothing, and Word breaks it where the page ends (generated documents: 55 to 120 paragraphs
        // each kept with the next, 51 lines a page in Word; each had gone to a page of its own, 113 pages for 14).
        if (!midSheet && m < k && m > 0 && keepsOn(blocks[m - 1], p)) m = k;
      }
      const moved = blocks[m];
      // At a new section's page, or moved on with what it keeps with, Word keeps the paragraph's space before; with
      // "page break before", as breakSpace has it (mode 14: less the space after above; from mode 15: none).
      // At a new section's page, less the space after of the paragraph ending the section before (sample S097:
      // a Title spaced 9pt before, under a section ended by a 3pt-after paragraph, 6pt down in Word's own layout; under
      // a 6pt one, 3pt; varied in copies of it, a space before up to that space after leaves none).
      const kept = forced && sectStart && m === k && k > 0 ? Math.max(0, spaceBefore(moved.node) - spaceAfter(blocks[k - 1].node))
        : forced && !sectStart && m === k && paraBreakBefore(moved.node) ? breakSpace(moved.node, k > 0 ? blocks[k - 1].node : null) : spaceBefore(moved.node);
      // (A floating table placed down from its paragraph keeps that drop on the page it moves to, the text after it starting
      // at the page's top. Word's prints of 12 generated tables, 120 to 300 twips down: all of the drop kept; a sample's
      // tables had lost it, 12px high and every page after them out.)
      // (And one placed up from it, upFromText, that far above the page's top, floating or laid out in the flow: Word's
      // prints, why-tests flt0898, 1 to 400tw up.)
      const floatDrop = moved.rows && outOfFlow(moved.el) && setDown(moved.node) ? ownTop(moved) : moved.rows && upFromText(moved.node) ? upLift(moved.node) : 0;
      const delta = contentTop(target) + floatDrop + (forced ? kept : tails.get(moved.i) || 0) - (moved.top + shift);
      starts.set(moved.i, (cols || carried.has(moved.i) ? starts.get(moved.i) || 0 : 0) + delta); // (in columns it may already carry a split's push; after a frame moved on, that push)
      // (The text after a floating table goes with it, a float's margin not moving it: it stayed at the old page's foot and
      // ran past it. Word's prints of the same tables: the next paragraph starts the new page, beside the table or above it.)
      if (moved.rows && outOfFlow(moved.el) && blocks[m + 1]) { addTo(starts, blocks[m + 1].i, delta); carried.add(blocks[m + 1].i); }
      shift += delta;
      if (rec.has(m)) rec.set(m, { top: moved.top + shift, bottom: moved.bottom + shift }); // (the walk goes on after it, not from it)
      sbTop.set(moved.i, forced ? kept : 0);
      for (let n = m; n <= k; n++) {
        if (n > m) sbTop.delete(blocks[n].i);
        shapeUp.delete(blocks[n].i);
        moveNotes(blocks[n].i, target);
        splits.delete(blocks[n].i);
        paraSplits.delete(blocks[n].i);
        colSplits.delete(blocks[n].i);
      }
      k = m; // re-check everything after the moved block
      // (Moved on to a page where a table is set down after it, into whose band it reaches: below that table, as
      // above. Word's print of a sample: a caption kept with the set-down table after it, both on to the next page,
      // the caption under the table there.)
      if (!inCols && (!band || band.p !== target)) shift += belowAhead(m, moved.top + shift, target);
      // (In columns, a paragraph taller than the column it moved to: itself too, so it's cut at that column's foot.)
      if (inCols && moved.node.type === schema.nodes.paragraph && moved.bottom + shift > contentBottom(target) + 1 && (recut.get(moved.i) || 0) < RECUT_MAX) {
        recut.set(moved.i, (recut.get(moved.i) || 0) + 1);
        k = m - 1;
      }
      breakPending = (!lineSplits || regs.some((r) => m >= r.k0 && m < r.k1)) && hasPageBreak(moved.node);
      breakKind = breakKindOf(moved.node);
      if (!inCols && moved.node.type === schema.nodes.paragraph) {
        const { up } = shapesPast(moved, shift, target);
        if (up.size) shapeUp.set(moved.i, { el: moved.el, up });
      }
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
          besideLast(m, target);
        }
      }
    }
    if (!cols && !regs.length) verticalAlign(doc, blocks, starts, pulls, splits, paraSplits);
    colVa = new Map();
    if (cols && !regs.length) verticalAlignCols(doc, blocks, starts, colSplits);
    if (cols && !regs.length) balanceEndCols(blocks, starts, colSplits);
    heldBands(blocks, starts, colSplits);
    const last = blocks.length ? blocks[blocks.length - 1].bottom + shift : 0;
    pages = Math.max(1, pageOf(Math.max(last - 1, 0)) + 1);
    // A band in columns grown (or shrunk) since the last pass, by a note further on in its sheet: its columns laid out
    // again to its top.
    // (A band going back to its height of two passes ago keeps the larger of the two: the columns before it were cut to
    // that, and it would flip until the passes ran out.)
    const nextH = new Map(notesH);
    for (const [b, f] of noteFlow) if (colGeo(b) && f.items.every((r) => r.kind === 'endnote')) nextH.delete(b); // (bandH)
    for (const b of new Set([...notesH.keys(), ...notesHPrev.keys()])) {
      if (!colGeo(b)) continue;
      const now = notesH.get(b) || 0, prev = notesHPrev.get(b) || 0;
      if (Math.abs(now - prev) > 0.5 && Math.abs(now - (notesHPrev2.get(b) || 0)) < 0.5) nextH.set(b, Math.max(now, prev));
      if (Math.abs((nextH.get(b) || 0) - prev) > 0.5) colsAgain = true;
    }
    notesHPrev2 = notesHPrev;
    notesHPrev = nextH;
    // (Notes continued past the body's last page: pages of their own, as Word adds.)
    if (noteFlow.size) pages = Math.max(pages, Math.max(...noteFlow.keys()) + 1);
    const { colCss, contHeads, blockAttrs } = placeColumns(blocks, regs, starts, colSplits, paraSplits, splits);
    for (const [i, dy] of frameDy) { if (!blockAttrs.has(i)) blockAttrs.set(i, { cls: [], style: [] }); blockAttrs.get(i).style.push(`--odx-fdy:${dy.toFixed(2)}px`); }
    for (const b of blocks) {
      if (!sbTop.has(b.i) || !b.el.querySelector?.('img[style*="--odx-sb"]')) continue;
      if (!blockAttrs.has(b.i)) blockAttrs.set(b.i, { cls: [], style: [] });
      blockAttrs.get(b.i).style.push(`--odx-sbtop:${Math.max(0, sbTop.get(b.i)).toFixed(2)}px`);
    }

    // The top-level paragraphs' print-only baseline shifts, on the blocks themselves (baseline.js).
    for (const [i, st] of blockShifts(blocks, sheet)) { if (!blockAttrs.has(i)) blockAttrs.set(i, { cls: [], style: [] }); blockAttrs.get(i).style.push(st); }
    const { css, heads } = applyPushes(doc, dom, blocks, base, colRegs, starts, pulls, splits, paraSplits, blockAttrs);
    css.push(`.${scope} > .odx-content{min-height:${pageTop(pages) - gap}px}`);
    if (cols || colRegs.length) dom.classList.remove('odx-measure');
    sheet.classList.remove('odx-bl-off');
    // Floats drawn up to end at their page's edge (floatsPast); those no longer drawn up, put back.
    const upNow = new Map();
    for (const [, { el, up }] of [...floatUp, ...shapeUp]) {
      const imgs = [...el.querySelectorAll('img')];
      for (const [n, dy] of up) if (imgs[n]) upNow.set(imgs[n], dy);
    }
    // (Only when there's one: quiet() holds the editor's watch on its DOM, which would miss what's typed meanwhile.)
    if (drawnUp.size || upNow.size) {
      quiet(() => {
        for (const img of drawnUp) if (!upNow.has(img)) { img.style.translate = ''; liftBox(img, 0); }
        for (const [img, dy] of upNow) {
          if (img.style.translate !== `0 ${(-dy).toFixed(2)}px`) img.style.translate = `0 ${(-dy).toFixed(2)}px`;
          liftBox(img, dy);
        }
      });
      drawnUp = new Set(upNow.keys());
    }
    // Pictures wrapped top and bottom drawn up to their page's foot (tbPast): out of the flow, so the text after them goes
    // on above them, and drawn up; those no longer, put back. Changed, the pages are walked again (a few times at most).
    const tbNow = new Map([...tbUp.values()].map((t) => [t.img, t.up]));
    let tbMoved = false;
    if (tbOut.size || tbNow.size) {
      quiet(() => {
        for (const img of tbOut) {
          if (tbNow.has(img)) continue;
          img.style.position = '';
          if (tbOwnMt.has(img)) img.style.marginTop = tbOwnMt.get(img);
          tbUpNow.delete(img);
          tbMoved = true;
        }
        for (const [img, dy] of tbNow) {
          if (!tbOwnMt.has(img)) tbOwnMt.set(img, img.style.marginTop);
          const mt = `calc(${tbOwnMt.get(img) || '0px'} - ${dy.toFixed(2)}px)`;
          if (img.style.position !== 'absolute') { img.style.position = 'absolute'; tbMoved = true; }
          // (Out of the flow: nothing else moves with it.)
          if (img.style.marginTop !== mt) img.style.marginTop = mt;
          tbUpNow.set(img, dy);
        }
      });
      tbOut = new Set(tbNow.keys());
    }
    let text = css.concat(colCss, baselineCss(scope)).join('\n');
    text += `\n@page { size: ${step(page.width)}px ${step(page.height)}px; margin: 0; }\n@media print { .${scope} { width: ${step(page.width)}px !important; height: ${pages * step(page.height) - 1}px; overflow: hidden; } .${scope} > .odx-content { width: ${page.width}px; } }`; // the browser's own print margins would shrink and shift each sheet
    if (rules.textContent !== text) rules.textContent = text; // unchanged rules: no style recalculation
    drawSheets();
    // Shapes with text on both sides moved by this walk (their paragraph moved on, or the shape drawn up or down): the
    // lines beside them were split where they were, so split them again where they are now, and walk again. Otherwise a
    // walk over lines not yet settled stayed: wrap-word4 vt-none, Text 20 moved on whole as its rectangle, drawn up past
    // the lines above as they were then, reached over them, and nothing laid the pages out again once they had moved
    // (unless a font came in later, so the pages came out one way or the other from one load to the next).
    let shapesNow = '';
    for (const b of blocks) {
      if (b.node.type !== schema.nodes.paragraph || !b.el.querySelector('img[style*="--odx-both"]')) continue;
      shapesNow += `${b.i}:${(applied.get(b.i)?.extra || 0).toFixed(1)}:${[...b.el.querySelectorAll('img[style*="--odx-both"]')].map((im) => im.style.translate).join('/')};`;
    }
    const shapesMoved = shapesNow !== shapesAt;
    shapesAt = shapesNow;
    // (Headers and footers drawn afresh with a shape the body's text wraps round: their lines split beside it, again.)
    const hfWrapAgain = !!layer.querySelector('img[style*="--odx-hfwrap"]') && wrapStep();
    drawHeads(heads.concat(contHeads));
    drawLineNumbers(view, lineNums, pageOf, (el) => carriedParts.get(el), (el) => blockCols.get(el));
    drawBalloons(view, session, sheet, pageTop(pages) - gap);
    lastMs = performance.now() - t0;
    passes++;
    // A balanced columns section's cut changed: lay out again with it (a few times at most, should it not settle).
    if (crossAgain) crossPasses++;
    if (hfWrapAgain || crossAgain) schedule(); // (crossAgain: once a table at most, crossFoot, and CROSS_MAX passes)
    else if (balanceAgain && balancePasses < 6) { balancePasses++; schedule(); }
    // Blocks moved into columns of other widths: again at those widths (a few times at most, should it not settle).
    else if (colsAgain && colsPasses < 8) { colsPasses++; colsRepass = true; schedule(); }
    else if (shapesMoved && shapesPasses < 6) { shapesPasses++; schedule(); }
    else if (tbMoved && tbPasses < 4) { tbPasses++; schedule(); }
    colsAgain = false;
    if (settleDoc !== doc) { settleDoc = doc; settleRuns = 0; }
    const rk = roundKey();
    if (!pushRound || pushRound.doc !== doc || pushRound.key !== rk) pushRound = { doc, key: rk, seen: [], target: null, done: false };
    if (pushesMoved) {
      if (pushRound.target === null) {
        const at = pushRound.seen.indexOf(pushState);
        if (at >= 0) pushRound.target = pushRound.seen.slice(at).sort()[0];
        else pushRound.seen.push(pushState);
      }
      if (pushRound.target === pushState) { pushRound.done = true; pushRound.seen = []; pushesMoved = false; }
    }
    if (!pushesMoved) settleRuns = 0;
    else if (settleRuns < SETTLE_MAX) { settleRuns++; passed = false; schedule(); }
    pushesMoved = false;
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
  // Where endnotes start, as drawSheets places them (under the text's end, or the columns balanced before them), for its
  // key: without it, columns cut a pass after the notes were first drawn, or the text above them grown or shrunk on the
  // same pages, left the endnotes where they were drawn first (columns-endnotes.docx: at the page's foot, under nothing).
  const endnotesAt = () => {
    if (![...notesOn.values()].some((l) => l.some((r) => r.kind === 'endnote'))) return '';
    const last = lastShown(view.dom);
    const y = cols && colsBal ? colsBal.bottom : last ? rectOf(last).bottom - rectOf(view.dom).top : 0;
    return y.toFixed(1);
  };
  function drawSheets() {
    if (seqDoc !== view.state.doc) { seqDoc = view.state.doc; docSeq++; }
    pageBorder = readPageBorder();
    // (Headers and footers with reference fields are drawn again as the document changes: a running head follows
    // its heading's text.)
    if (refsInStories == null) refsInStories = /\b(?:STYLEREF|PAGEREF|REF)\b/.test(Object.keys(session.pkg.files).filter((k) => /word\/(header|footer)\d*\.xml$/.test(k)).map((k) => session.pkg.text(k)).join(''));
    const key = `${refsInStories ? view.state.doc.content.size + ':' + docSeq + ':' : ''}${(view.state.doc.attrs.sectPr || '').match(/<w:pgBorders\b[\s\S]*?<\/w:pgBorders>/)?.[0] || ''}:${pages}:${margins.map((m) => m[0] + (m[2] ? 'f' : '')).join(',')}:${pitch}:${pageTop(pages)}:${[...notesOn].map(([p, l]) => p + '=' + l.map((r) => r.kind[0] + r.id + '.' + r.n).join('+')).join(',')}:${endnotesAt()}:${[...noteFlow].map(([p, f]) => `${p}/${Math.round(notesH.get(p) || 0)}/${Math.round(f.h)}/${Math.round(f.skip || 0)}`).join(',')}`;
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
          h.style.cssText = `top:${hf.headerDist}px;left:${at[1].left ?? page.left}px;right:${at[1].right ?? page.right}px;--odx-hf-left:${at[1].left ?? page.left}px`;
          // (The page's geometry, for its shapes placed against the page or margins: hfstory.js placeStoryShapes.)
          Object.assign(h.dataset, { odxPh: pageH(p), odxMtop: marginsAt(p).top, odxMbot: marginsAt(p).bottom });
          // A header reaching most of the way down the page (past 70% of it) is cut off 60px above the middle of the page,
          // as Word draws it. (Word's PDFs of the generated documents: a 0.8-page header cut there, a 0.6-page one whole.)
          if (hf.headerDist + storyHeight(header, false, hf.headerDist) > pageH(p) * (1 - HF_FAR)) h.style.cssText += `;max-height:${Math.max(0, pageH(p) / 2 - HF_GAP - hf.headerDist)}px;overflow:hidden`;
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
          f.style.cssText = `bottom:${hf.footerDist}px;left:${at[1].left ?? page.left}px;right:${at[1].right ?? page.right}px;--odx-hf-left:${at[1].left ?? page.left}px`;
          Object.assign(f.dataset, { odxPh: pageH(p), odxMtop: marginsAt(p).top, odxMbot: marginsAt(p).bottom });
          // A footer reaching most of the way up the page (above 30% of it) starts 46pt below the middle of the page and is
          // cut off at its distance from the foot, as Word draws it (not running up off the page).
          const fTop = pageH(p) / 2 + HF_GAP + 4 / 3;
          if (pageH(p) - hf.footerDist - storyHeight(footer, true) < pageH(p) * HF_FAR) f.style.cssText = `top:${fTop}px;left:${at[1].left ?? page.left}px;right:${at[1].right ?? page.right}px;--odx-hf-left:${at[1].left ?? page.left}px;max-height:${Math.max(0, pageH(p) - hf.footerDist - fTop)}px;overflow:hidden`;
          s.append(f);
        }
      }
      // (In columns, the sheet's band: kept by the flow page of its first column; in the per-column model, each column's
      // own, under it.)
      const nk0 = sheetBand ? sheetBand.get(p) ?? p : p;
      const g0 = perCol() && colGeo(nk0);
      const areas = g0 ? Array.from({ length: g0.num }, (_, c) => nk0 + c) : [nk0];
      for (const nk of areas) {
      const flow = noteFlow.get(nk);
      if (!flow) continue;
      const at = g0 ? colAt(nk) : null;
      {
        const box = document.createElement('div');
        box.className = 'odx-notes';
        let top = pageH(p) - marginsAt(p).bottom - (notesH.get(nk) ?? flow.h + NOTE_SEP);
        // Endnotes start where the text ends, as in Word (its separator line under the last paragraph's space after),
        // not at the page's foot as footnotes do: the generated documents' endnotes, tools/doc-generator.
        const last = lastShown(view.dom);
        const endOnly = flow.items.every((r) => r.kind === 'endnote');
        // (In columns, a page of endnotes alone, after the text's last: from the top, its columns left to right whatever
        // the section's direction, as Word prints them; going on from the page before, under a continuation separator.)
        const alone = endOnly && !!flow.cols && !flow.own;
        if (alone) top = marginsAt(p).top + (flow.cont ? 0 : ENDNOTE_TOP);
        else if (last && !flow.skip && endOnly) {
          // (In columns balanced before them, under the lowest column: balanceEndCols.)
          const y = cols && colsBal ? marginsAt(p).top + colsBal.bottom + ENDNOTE_LEAD : rectOf(last).bottom + (parseFloat(getComputedStyle(last).marginBottom) || 0) - rectOf(view.dom).top - pageTop(p) + ENDNOTE_LEAD;
          if (y > 0 && y < top) top = y;
        }
        box.style.cssText = at ? `top:${top}px;left:${page.left + at.x}px;width:${at.w}px` : `top:${top}px;left:${page.left}px;right:${page.right}px`;
        const sep = document.createElement('div');
        sep.className = noteSepRtl(endOnly ? 'endnote' : 'footnote') ? 'odx-note-sep odx-note-sep-rtl' : 'odx-note-sep';
        // (Notes carried on from the page before: the part not yet shown, cut where this page's room ends.)
        const clip = document.createElement('div');
        clip.style.cssText = `overflow:hidden;height:${flow.h}px`;
        const inner = document.createElement('div');
        inner.style.marginTop = `${-flow.skip}px`;
        inner.append(...flow.items.map((r) => noteEl(r.note, r.kind, r.n)));
        clip.append(inner);
        if (flow.cols) { // in the band's columns (endnotes too, in a one-section document of columns)
          clip.style.height = 'auto';
          const cont = alone && flow.cont;
          clip.replaceChildren(bandBox(flow.items, alone ? { ...flow.cols, rtl: false } : flow.cols, flow.lay, flow.fit, cont ? sep : null));
          if (cont) box.append(clip); else box.append(sep, clip);
        } else box.append(sep, clip);
        s.append(box);
      }
      }
      frag.append(s);
    }
    layer.replaceChildren(frag);
    for (const hf of layer.querySelectorAll('.odx-hf')) { storyLines(hf); settleStoryShapes(hf); } // (drawn now, so they can be measured)
    clipExactRows(layer); // (drawn now, so they can be measured)
    fitRowspans(layer);
    // Header and footer tabs to their stops (center, right...), as the body's: they had stayed plain tab characters,
    // which collapse to a space (sample S078's centred "Company Proprietary").
    // (Their text at Word's widths first, as the body's: widths.js fixStoryWidths.)
    for (const hf of layer.querySelectorAll('.odx-hf')) { fixStoryWidths(hf); layoutTabsIn(hf, session, (p) => { const n = hfNodes.get(p); if (!n) throw new Error('no node'); return n; }); layoutPtabs(hf); }
    hfShifts(layer, sheet); // the print-only baseline shifts of header and footer text, as the body's
    const gf = document.createDocumentFragment();
    if (gap) for (let p = 0; p + 1 < pages; p++) {
      const g = document.createElement('div');
      g.style.cssText = `top:${pageTop(p) + step(pageH(p))}px;height:${gap}px`;
      gf.append(g);
    }
    gaps.replaceChildren(gf);
  }

  // A drawing out of the flow (in front of the text or behind it) shows only on its paragraph's page, as in Word: what
  // reaches past that page's edges is cut off, not drawn over the gap, beside the page or on the next page (sample 2466's
  // cover: a page-wide box behind the text 78pt left of the column and 31pt past the right edge, and a text box 176px
  // under the page's last paragraph, which Word doesn't print). Turned drawings are left as they are (a clip is in their
  // own turned box).
  function clipToSheets() {
    const sheets = [...layer.children].filter((s) => s.classList.contains('odx-sheet')).map((s) => s.getBoundingClientRect());
    if (!sheets.length) return;
    for (const img of view.dom.querySelectorAll('img:not(.ProseMirror-separator)')) {
      const had = img.dataset.odxSheetclip;
      let want = '';
      const cs = getComputedStyle(img);
      if (cs.position === 'absolute' && cs.transform === 'none' && cs.rotate === 'none' && !/--odx-picclip/.test(img.getAttribute('style') || '')) {
        const host = img.closest('p, td') || img.parentElement;
        const y = host.getBoundingClientRect().top;
        const s = sheets.find((r) => y >= r.top - 1 && y < r.bottom) || sheets.filter((r) => r.top <= y).pop();
        if (s) {
          const r = img.getBoundingClientRect();
          const cut = [s.top - r.top, r.right - s.right, r.bottom - s.bottom, s.left - r.left].map((v, k) => (v > 0.5 ? +Math.min(v, k % 2 ? r.width : r.height).toFixed(2) : 0));
          if (cut.some((v) => v > 0)) want = `inset(${cut.map((v) => `${v}px`).join(' ')})`;
        }
      }
      if (want === (had || '')) continue;
      if (want) { img.style.clipPath = want; img.dataset.odxSheetclip = want; } else { img.style.removeProperty('clip-path'); delete img.dataset.odxSheetclip; }
    }
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
  const onLoad = (e) => { if (e.target.tagName === 'IMG') { pushRound = null; scheduleIdle(); } };
  view.dom.addEventListener('load', onLoad, true);
  // Long documents skip laying out off-screen blocks (content-visibility) and
  // paginate on estimated heights; when a block is first drawn its real height
  // replaces the estimate, so paginate again.
  const onShown = (e) => { if (!e.skipped) { pushRound = null; scheduleIdle(); } };
  view.dom.addEventListener('contentvisibilityautostatechange', onShown, true);
  window.addEventListener('resize', scheduleIdle);
  if (document.fonts) fontsSettled().then(() => !destroyed && schedule());
  // A long document opens fast, laid out from estimates (odx-large: off-screen blocks skipped, paragraphs moved to the next
  // page whole). Once it is drawn and the browser is idle, it is laid out in full as any document, so its pages are Word's
  // on screen as in print: a 122-page document's heading and the paragraph after it had gone to the next page whole, a
  // third of a page left empty, 125 pages. (About 1.6s more for that document, after it is shown, not before.)
  if (sheet.classList.contains('odx-large')) {
    const full = () => {
      if (destroyed || printing || !sheet.classList.contains('odx-large')) return;
      // (Only once the pages it opened with are drawn and have stayed so for a second: started while they were still
      // being laid out, as Firefox's take longer, it held them back 3s more.)
      if (frame || performance.now() - lastPass < 1000) { setTimeout(() => idle(full), 500); return; }
      sheet.classList.remove('odx-large');
      pushRound = null;
      schedule();
    };
    const idle = (fn) => (typeof requestIdleCallback === 'function' ? requestIdleCallback(fn, { timeout: 4000 }) : setTimeout(fn, 500));
    (document.fonts ? fontsSettled() : Promise.resolve()).then(() => setTimeout(() => idle(full), 1500));
  }
  // A web font that arrives later (a stand-in such as Carlito, requested only when the text using it is drawn) changes
  // what was measured in the fallback font: tabs, widths and pages are worked out again.
  // (From the text's own flow: a spacer put between two lines in the fallback font holds that line break in the real one,
  // so the pages came out by which arrived first, the font or the first layout.)
  const fromNaturalFlow = () => {
    const set = pushKey.getState(view.state);
    const all = set.find();
    const keep = all.filter((d) => !/^push(ed)?:/.test(String(d.spec.key || '')));
    if (keep.length === all.length) return;
    for (const a of applied.values()) a.inline = 0;
    pushState = ''; // (no longer the pushes the lines stand under: applyWrapBoth tells its rounds by them)
    pushesKnown = false;
    view.dispatch(view.state.tr.setMeta(pushKey, keep).setMeta('addToHistory', false));
  };
  const onFonts = () => { if (destroyed) return; fontArrivals++; fromNaturalFlow(); scheduleIdle(); };
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
    printing = true;
    // (Printed at full size, whatever the zoom on screen.)
    printZoom = sheet.style.zoom;
    sheet.style.zoom = '';
    // (At full size a line can break a little differently: the words the hyphenation zone and limit leave whole, again.)
    settleHyphens(view);
    // Word updates the body's fields as it prints (page references, cross-references): so do we, from the pages.
    updateFields(view, { all: true });
    // A very long document is laid out on screen from estimates for what's off screen, so where its pushes stand depends on
    // where it has been scrolled; and its passes may not settle (the browser rounds some lines' heights by where they
    // land, 1/32px, and a thousand tables keep moving each other). Its print starts from no pushes at all, so the same
    // passes from the same place give the same pages, whatever was on screen before.
    if (sheet.classList.contains('odx-large')) {
      applied = new Map();
      pushRound = null;
      pushState = '';
      pushesKnown = false;
      settleRuns = 0;
      rules.textContent = '';
      view.dispatch(view.state.tr.setMeta(pushKey, []).setMeta('addToHistory', false));
    }
    setGap(0);
    // The print is taken right after this handler: finish the layout now rather than on later frames, which a page
    // printed in the background may not get (its pages kept their on-screen gaps).
    // (At least once: the fields' update can draw tables again, at their grid's widths, and the layout before had passed.)
    passed = false;
    // (Up to more passes than any one step may take, wrapboth.js's 81: stopped at 8, a layout still settling printed with
    // no pages drawn.)
    for (let k = 0; k < 120 && !passed; k++) {
      if (frame) cancelAnimationFrame(frame);
      layout();
    }
    // Lifted floats read where pages are from the layout before (liftFloats): settle them on the gapless pages.
    if (view.dom.querySelector('img[style*="--odx-lrel"], img[style*="--odx-arel"]')) for (let k = 0; k < 2; k++) { if (frame) cancelAnimationFrame(frame); layout(); }
    if (engineOn()) buildPrintLines(sheet, view.dom, (el) => shiftFor(el, sheet), scope, quiet);
    printBox?.remove();
    printBox = printCopies(sheet, pages, { pageW, pageH, pageTop });
    if (printBox) document.body.append(printBox);
  };
  const afterPrint = () => { printing = false; removePrintLines(sheet, quiet); printBox?.remove(); printBox = null; setGap(GAP); if (printZoom) sheet.style.zoom = printZoom; printZoom = ''; settleHyphens(view); };
  window.addEventListener('beforeprint', beforePrint);
  window.addEventListener('afterprint', afterPrint);

  return {
    update: scheduleIdle,
    balloons: () => unzoomed(() => drawBalloons(view, session, sheet, pageTop(pages) - gap)),
    /** Headers or footers changed: measure and draw them again. */
    restory() {
      sectCache = { doc: null, ends: new Map(), first: page };
      heights.clear();
      notes.clear();
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
      buildPrintLines(sheet, view.dom, (el) => shiftFor(el, sheet), scope, quiet);
      return () => removePrintLines(sheet, quiet);
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
