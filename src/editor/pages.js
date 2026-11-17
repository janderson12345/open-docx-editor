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
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { schema } from './schema.js';
import { layoutTabs, layoutTabsIn } from './tabs.js';
import { fitRowspans, clipExactRows } from './rowfit.js';
import { storyDrawer, hfNodes, settleStoryShapes } from './hfstory.js';
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
import { applyLineFix } from './lhfix.js';
import { applyWidths } from './widths.js';
import { applyFitText } from './fittext.js';
import { applyWrapBoth, drawnUpTo } from './wrapboth.js';
import { settleHyphens } from './hyphenate.js';
import { fitAutofitColumns } from './tablefit.js';
import { baselineCss, cellShifts, hfShifts, shiftFor } from './baseline.js';
import { buildPrintLines, removePrintLines, engineOn } from '../engine/printlines.js';
import { twipsToPx } from '../ooxml/units.js';
import { placeVar, padOf } from '../ooxml/place.js';

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
const bothShape = (e) => e.tagName === 'IMG' && !!e.style.getPropertyValue('--odx-both') && getComputedStyle(e).position === 'absolute';

const GAP = 24; // px between sheets
const TIE = 0.33; // px an exactly spaced line must end above the page's bottom margin (splitParagraph)
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
    return Math.min(pageTop(p) + pageH(p) - marginsAt(p).bottom - (h ? h + NOTE_BODY_GAP : 0), cuts.get(p) ?? Infinity);
  };
  const NOTE_BODY_GAP = 4;
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
  const notes = createNoteBoxes({ page, sheet, session });
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
    const base = view.dom.getBoundingClientRect().top + b.undo;
    const pushes = [...b.el.querySelectorAll('.odx-push')].map((e) => e.getBoundingClientRect());
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
      const y0 = el.getBoundingClientRect().top;
      const nat = natural(y0);
      moveNote(r, bandOf(pageOf(nat + shift0 + split.pushes.reduce((sum, x) => sum + (x.pos <= r.pos ? x.delta : 0), 0))));
    }
  };
  // In columns, a footnote whose reference is in the lines a paragraph carries into the next column goes in that
  // column's band (its sheet's, when the column is on the next page), as in Word.
  const notesByCol = (b, p, cs) => {
    const list = blockNotes.get(b.i);
    if (!list) return false;
    const top = b.el.getBoundingClientRect().top;
    const parts = carriedOf(cs);
    let moved = false;
    for (const r of list) {
      // (A footnote by its reference's line, in the column of the part that holds it; endnotes, after the text, with
      // its last lines.)
      let to;
      if (r.kind === 'footnote') {
        const el = view.nodeDOM(r.pos);
        if (!el || el.nodeType !== 1) continue;
        const y = el.getBoundingClientRect().top - top;
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
    const y = el.getBoundingClientRect().top - b.el.getBoundingClientRect().top;
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
    const er = b.el.getBoundingClientRect();
    const up = new Map();
    [...b.el.querySelectorAll('img')].forEach((img, n) => {
      if (getComputedStyle(img).float === 'none') return;
      const r = img.getBoundingClientRect();
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
  const shapesPast = (b, shift, p) => {
    const er = b.el.getBoundingClientRect();
    const up = new Map();
    let past = false;
    [...b.el.querySelectorAll('img')].forEach((img, n) => {
      if (!bothShape(img)) return;
      const r = img.getBoundingClientRect();
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
      if (over <= PAST_MIN) return;
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
        const sr = sib.getBoundingClientRect();
        if (sr.bottom <= top) break;
        const walker = document.createTreeWalker(sib, NodeFilter.SHOW_TEXT);
        for (let t = walker.nextNode(); t && !past; t = walker.nextNode()) {
          if (!t.data.trim()) continue;
          range.selectNodeContents(t);
          for (const q of range.getClientRects()) if (q.width > 0 && q.bottom > top + 0.25 && q.top < own && q.right > left && q.left < right) { past = true; break; }
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
  const tablesIn = (td) => [...td.children].map(cellTable).filter(Boolean);
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
        const t = cellTable(el);
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
    const natural = naturalOf(b);
    const frags = [];
    const walker = document.createTreeWalker(b.el, NodeFilter.SHOW_TEXT);
    const range = document.createRange();
    for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      if (t.parentElement.closest('.odx-push')) continue;
      range.selectNodeContents(t);
      for (const r of range.getClientRects()) if (r.height > 0) frags.push(r);
    }
    // (A turned picture by its box, not its turned outline: the room it takes on its line.)
    // (Not a shape with text on both sides: bothShape.)
    b.el.querySelectorAll('img:not(.ProseMirror-separator), br:not(.ProseMirror-trailingBreak)').forEach((e) => {
      if (bothShape(e)) return;
      const r = layoutRect(e);
      if (r.height > 0) frags.push(r);
    });
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
        if ((n.tagName === 'IMG' || n.tagName === 'BR') && !bothShape(n) && onLine(n.getBoundingClientRect())) return view.posAtDOM(n.parentNode, [...n.parentNode.childNodes].indexOf(n));
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
  /**
   * How far a paragraph's last line may hang below the page's text area: Word places the extra space of
   * "multiple" line spacing below the text and doesn't count it when fitting a line (1px for rounding).
   */
  function lineSlack(el) {
    const cs = getComputedStyle(el);
    const mult = parseFloat(cs.getPropertyValue('--odx-mult')) || 1;
    const lh = parseFloat(lineHeightOf(cs));
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
    const lblUp = labelPad && b.el.classList.contains('odx-li') ? parseFloat(getComputedStyle(b.el, '::before').paddingTop) || 0 : 0;
    const lead0 = lines[0].screenTop - (b.el.getBoundingClientRect().top + parseFloat(pcs.paddingTop) + parseFloat(pcs.borderTopWidth)) - lblUp;
    const atLeast = lh1 && parseFloat(pcs.getPropertyValue('--odx-runmult')) !== 0;
    const lead = (mult > 1 || atLeast) && Math.abs(lead0) <= 20 ? lead0 : 0;
    return { lead0, lead };
  }

  /** Where a paragraph breaks at its column's bottom: { cut (px below its top), height, delta (for what follows) } or null. */
  function colSplit(b, shift, firstOnPage, p) {
    const lines = paragraphLines(b);
    if (!lines || lines.length < 2 || keepLines(b.node)) return null;
    const slack = lineSlack(b.el) - 1;
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
    cs.delta += h - (cs.height - cs.cut) - (cs.lift || 0); // (a fresh copy: its first line has its own lead)
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
    const h = headerRowsOf(b.node);
    const hb = h && h < trs.length ? trs[h].getBoundingClientRect().top - top0 : 0;
    const head = hb && cut >= hb - 0.5 ? hb : 0; // (the header rows drawn again above what goes on, once the cut is past them)
    // A table longer than the next column too goes on in the one after, and so on: cut again at a row's top at each
    // column's foot (columns of equal widths), its header rows again above each part.
    const parts = [cut];
    const cg = cols || colGeo(p);
    for (let c = 1; cg && !cg.unequal && c < MAX_CARRY; c++) {
      const from = parts[parts.length - 1];
      const fit = contentBottom(p + c) - contentTop(p + c) - head + 1;
      const j = trs.findIndex((tr) => tr.getBoundingClientRect().bottom - top0 - from > fit);
      if (j < 0) break;
      const next = trs[j].getBoundingClientRect().top - top0;
      if (next <= from + 0.5) break; // (a row taller than a column: it stays whole)
      parts.push(next);
    }
    const lastCut = parts[parts.length - 1];
    return { cut, head, parts: parts.length > 1 ? parts : undefined, height: b.bottom - b.top, delta: contentTop(p + parts.length) + head - (b.top + shift + lastCut) };
  }

  /**
   * Lines that must start a new page: after a page break, or where the page
   * runs out (with Word's widow/orphan control and keep-lines-together).
   * Returns null when the paragraph should move to the next page whole.
   */
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
      const q = e && e.nodeType === 1 ? e.getBoundingClientRect() : null;
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
      else if (fitBottom > room(n) + slack && n > pageStart && !(ownNotes.length && partialFits(pg, n, bottom))) {
        if (keepLines(b.node) && pushes.length === 0) return blockFirstOnPage ? { pushes, total: 0 } : null;
        at = n;
        if (widow) {
          at = widowCut(at, lines.length, blockFirstOnPage, pageStart);
          if (at < 0) return null;
        }
      } else if (fitBottom > room(n) + slack && n === 0 && !blockFirstOnPage && !lines[1]?.forced) return null;
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
  // A paragraph's part carried to the next column, as drawn there: block element -> { copy, cut, rewrap } (for its
  // line numbers, which go beside the lines where they're seen).
  let carriedParts = new WeakMap();
  let blockCols = new WeakMap(); // block element -> its column ({ x from the text's left, w }), in columns
  const nodeIds = new WeakMap(); // block node -> a number (a block whose content changed is a new node)
  let idSeq = 0;
  function drawHeads(list) {
    // (Unchanged repeats are left as they are: redrawing a long document's every pass took a third of its layout time.)
    const key = list.map((h) => { if (!nodeIds.has(h.b.node)) nodeIds.set(h.b.node, ++idSeq); return `${nodeIds.get(h.b.node)}:${h.b.i}:${h.b.top.toFixed(1)}:${h.top}:${h.rule ? 1 : 0}:${h.dy ?? ''}:${h.dx ?? ''}:${h.cut ?? ''}:${h.stop ?? ''}:${h.end ?? ''}:${h.rewrap ? `${h.rewrap.width.toFixed(1)}/${h.rewrap.h.toFixed(1)}` : ''}`; }).join(',');
    if (!list.length && !headKey) return;
    if (key === headKey && repeats.childElementCount) return;
    headKey = key;
    carriedParts = new WeakMap();
    const origin = layer.getBoundingClientRect();
    const frag = document.createDocumentFragment();
    for (const { b, top, rule, cont, cut, stop, end, dx, dy, rewrap, col } of list) {
      if (rewrap) {
        // Its lines laid out again at the next column's width (rewrapCarry): from the top of that copy.
        const r = b.el.getBoundingClientRect();
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
  // How many rows a table starts with that repeat at each page's top (w:tblHeader).
  const headerRowsOf = (table) => { let h = 0; while (h < table.childCount && isHeader(table.child(h))) h++; return h; };
  const isHeader = (row) => /<w:tblHeader(?![^>]*w:val="(?:0|false|off)")/.test(row.attrs.trPr || '');
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

  const heights = new Map();
  const storyHeight = (doc) => {
    if (!doc) return 0;
    if (!heights.has(doc)) {
      const el = story(doc, 'odx-header', 1, 1);
      el.style.cssText = `visibility:hidden;top:0;left:${page.left}px;right:${page.right}px`;
      sheet.append(el);
      settleStoryShapes(el); // (a shape the body wraps round, below its paragraph: not in its height)
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
  // Where a block's first line is (px, screen): its box's top, unless pagination moved its lines to the next page by a
  // spacer inside it (a paragraph whose box still starts at the foot of the page before: pagefloat-word pf-early's anchor,
  // its picture lifted onto the wrong page).
  const firstLineTop = (el) => {
    const top = el.getBoundingClientRect().top;
    const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.data.trim() && !n.parentElement.closest('.odx-lift, .odx-push') ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP) });
    const t = w.nextNode();
    if (!t) return top;
    const rg = document.createRange();
    rg.setStart(t, 0); rg.setEnd(t, Math.min(1, t.data.length));
    const q = rg.getClientRects()[0];
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
      const drop = Math.round((y - pad - Math.max(contentTop(pg), first.getBoundingClientRect().top - base)) * 10) / 10;
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
      const top = Math.round((y - (cb.getBoundingClientRect().top - base)) * 10) / 10;
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
      const r = el.getBoundingClientRect();
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
          const h = headerRowsOf(node);
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
        blockNotes.get(i).push({ ...r, h: noteHeight(note, r.kind, r.n), note, bi: i });
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
      // (Its parts carried on, each at a column's top: the last ends with its space after.)
      const parts = x.parts || (x.carry ? [{ off: 1, h: x.carry, stop: null }] : []);
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
        if (up || x) { const a = attrsOf(b.i); a.cls.push('odx-colmv'); a.style.push(`--odx-colt:translate(${x.toFixed(1)}px,${-(up * pitch)}px)`); }
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
            contHeads.push({ b, cont: true, cut: part.cut, stop: part.stop, rewrap: cs.rewrap, col: q.cw ? { x: q.x, w: q.cw } : null, dx: q.x - x, dy: (contentTop(v2) + va + head + (cs.rewrap ? 0 : cs.lift || 0) - up2 * pitch) - (b.top + sh + part.cut - up * pitch) });
            if (head) contHeads.push({ b, cont: true, cut: 0, end: head, dx: q.x - x, dy: (contentTop(v2) + va - up2 * pitch) - (b.top + sh - up * pitch) });
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
        const bySheet = [];
        for (const [f, v, fresh] of margins) {
          const at = sheetOfFlow(f);
          if (bySheet.length && bySheet[bySheet.length - 1][0] === at) bySheet.pop(); // (the later one wins a sheet)
          bySheet.push([at, v, fresh]);
        }
        margins = bySheet;
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
      const er = b.el.getBoundingClientRect();
      for (const f of b.el.querySelectorAll('img')) {
        if (getComputedStyle(f).position !== 'absolute' && getComputedStyle(f).float === 'none') continue;
        r.reach = Math.max(r.reach, b.top + acc + (f.getBoundingClientRect().bottom - er.top));
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
    doc.forEach((node, pos, i) => { if (unfloated.has(node)) onBlock(i, { style: 'float:none !important' }, 'nf'); });
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
      const cut = bands.map(([a, z]) => `V${a.toFixed(1)} H${-CLIP_TOP} V${z.toFixed(1)} H${CLIP_FAR}`).join(' ');
      css.push(`${sel} table{clip-path:path('M${-CLIP_TOP} ${-CLIP_TOP} H${CLIP_FAR} ${cut} V${CLIP_DEEP} H${-CLIP_TOP} Z')}`);
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
    // (And round the shapes wrapped round in the headers and footers drawn on the sheets: wrapboth.js storyZones.)
    if (applyWrapBoth(view, layer)) return schedule(); // a line split beside a shape with text on both sides: measure again
    // (Its own pass: a view update only lays out again when the document changed, and a squeeze that settled after
    // everything else (lines split beside a shape, wrapboth.js) had left the pages undrawn.)
    if (applySqueeze(view, session)) return schedule(); // lines changed: measure again once they're drawn
    if (quiet(() => fitRowspans(view.dom))) return schedule(); // merged cells' extra height moved to their last row: measure again
    passed = true;
    const dom = view.dom;
    const doc = view.state.doc;
    // Word lets a floating table longer than a page run on in the flow: nothing wraps beside it.
    // Kept as a rule too: redrawing the table's decorations would reset an inline style.
    // (Each block's own element, as measureBlocks finds it: the DOM's children needn't line up with the document's.)
    const topEls = topElements(doc);
    doc.forEach((node, pos, i) => {
      const t = topEls ? topEls[i] : view.nodeDOM(pos);
      if (!t || t.nodeType !== 1) return;
      if (getComputedStyle(t).float !== 'none' && t.getBoundingClientRect().height > pageHeight(0)) { quiet(() => { t.style.float = 'none'; }); unfloated.add(node); }
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
    let balanceAgain = false;
    const rec = new Map(); // (columns sections) block k -> { top, bottom, carry } where it was last placed
    const frameDy = new Map(); // block index -> px a frame set down on its page is moved below where it comes
    floatUp.clear();
    shapeUp.clear();
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
    // A table set down at its place on its page (floatAt).
    const placeSetDown = (k, b, top, p, fy) => {
      const y0 = (fy.rel === 'page' ? pageTop(p) : contentTop(p)) + fy.y;
      const nb = blocks[k + 1];
      const d1 = y0 - top;
      const d2 = nb ? top + (nb.top - b.bottom) - (nb.top + shift + d1) : 0;
      pulls.set(b.i, nb ? [[b.i, d1], [nb.i, d2]] : [[b.i, d1]]);
      shift += d1 + d2;
      band = { k, p, y0: y0 - fy.above, y1: y0 + (b.bottom - b.top) + fy.below };
      // With less than two lines' room under it, the page's text ends above it: a table reaching into it splits there
      // and goes on on the next page (the contingency plan's page 4), rather than all of it moving under the band.
      if (band.y1 > contentBottom(p) - MIN_ROOM) { cuts.set(p, band.y0); band.cut = true; }
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
      // (Less the space after the paragraph above: Word drops it at the page's foot, so a page filled exactly keeps its
      // section's empty last paragraph.)
      // (Not after a column break: that paragraph goes to the next column, and a next-page section starts the page after.)
      if (!firstOnPage && !b.node.childCount && sectEnds.get(b.i) && !(breakPending && breakKind === 'column') && top - (k > 0 ? spaceAfter(blocks[k - 1].node) : 0) < contentBottom(p)) { if (sectEnds.get(b.i).page) breakPending = false; continue; }
      // (In columns of their own widths, a paragraph's bottom at the width of the column it's in: atColumnWidth.)
      const wide = inCols ? atColumnWidth(b, cols || inReg, p, shift, top <= contentTop(p) + 1) : null;
      const bottom = wide ? wide.bottom : b.bottom;
      // (Its shapes with text on both sides past the foot, drawn up: shapesPast. Not in columns.)
      let shapePast = false;
      if (!inCols && b.node.type === schema.nodes.paragraph) {
        const sp = shapesPast(b, shift, p);
        if (sp.up.size) shapeUp.set(b.i, { el: b.el, up: sp.up });
        shapePast = sp.past;
      }
      let target = null;
      let overflow = false;
      let forced = false;
      let sectStart = false;
      if (k > 0 && (breakPending || breakBefore(b.node) || (sect && sect.page)) && !firstOnPage) {
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
      else if (inCols && b.node.type === schema.nodes.paragraph && !hasPageBreak(b.node) && bottom + shift + bandAfter(p, b.node) > contentBottom(p) + lineSlack(b.el) - 1) {
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
      else if (lineSplits && !inCols && b.node.type === schema.nodes.paragraph && (hasPageBreak(b.node) || b.bottom + shift + bandAfter(p, b.node) > contentBottom(p) + (footBand(p) ? 0 : lineSlack(b.el) - 1) - tie(b.el))) {
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
      breakPending = (!lineSplits || !!inReg) && hasPageBreak(b.node);
      breakKind = breakKindOf(b.node);
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
        shapeUp.delete(blocks[n].i);
        moveNotes(blocks[n].i, target);
        splits.delete(blocks[n].i);
        paraSplits.delete(blocks[n].i);
        colSplits.delete(blocks[n].i);
      }
      k = m; // re-check everything after the moved block
      // (In columns, a paragraph taller than the column it moved to: itself too, so it's cut at that column's foot.)
      if (inCols && moved.node.type === schema.nodes.paragraph && moved.bottom + shift > contentBottom(target) + 1) k = m - 1;
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
    let text = css.concat(colCss, baselineCss(blocks, scope, sheet)).join('\n');
    text += `\n@page { size: ${page.width}px ${page.height}px; margin: 0; }\n@media print { .${scope} { height: ${pages * step(page.height) - 1}px; overflow: hidden; } }`; // the browser's own print margins would shrink and shift each sheet
    if (rules.textContent !== text) rules.textContent = text; // unchanged rules: no style recalculation
    drawSheets();
    // (Headers and footers drawn afresh with a shape the body's text wraps round: their lines split beside it, again.)
    const hfWrapAgain = !!layer.querySelector('img[style*="--odx-hfwrap"]') && applyWrapBoth(view, layer);
    drawHeads(heads.concat(contHeads));
    drawLineNumbers(view, lineNums, pageOf, (el) => carriedParts.get(el), (el) => blockCols.get(el));
    drawBalloons(view, session, sheet, pageTop(pages) - gap);
    lastMs = performance.now() - t0;
    passes++;
    // A balanced columns section's cut changed: lay out again with it (a few times at most, should it not settle).
    if (hfWrapAgain) schedule();
    else if (balanceAgain && balancePasses < 6) { balancePasses++; schedule(); }
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
  // Where endnotes start, as drawSheets places them (under the text's end, or the columns balanced before them), for its
  // key: without it, columns cut a pass after the notes were first drawn, or the text above them grown or shrunk on the
  // same pages, left the endnotes where they were drawn first (columns-endnotes.docx: at the page's foot, under nothing).
  const endnotesAt = () => {
    if (![...notesOn.values()].some((l) => l.some((r) => r.kind === 'endnote'))) return '';
    const last = view.dom.lastElementChild;
    const y = cols && colsBal ? colsBal.bottom : last ? last.getBoundingClientRect().bottom - view.dom.getBoundingClientRect().top : 0;
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
          h.style.cssText = `top:${hf.headerDist}px;left:${at[1].left ?? page.left}px;right:${at[1].right ?? page.right}px`;
          // A header reaching most of the way down the page (past 70% of it) is cut off 60px above the middle of the page,
          // as Word draws it. (Word's PDFs of the generated documents: a 0.8-page header cut there, a 0.6-page one whole.)
          if (hf.headerDist + storyHeight(header) > pageH(p) * (1 - HF_FAR)) h.style.cssText += `;max-height:${Math.max(0, pageH(p) / 2 - HF_GAP - hf.headerDist)}px;overflow:hidden`;
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
          // A footer reaching most of the way up the page (above 30% of it) starts 46pt below the middle of the page and is
          // cut off at its distance from the foot, as Word draws it (not running up off the page).
          const fTop = pageH(p) / 2 + HF_GAP + 4 / 3;
          if (pageH(p) - hf.footerDist - storyHeight(footer) < pageH(p) * HF_FAR) f.style.cssText = `top:${fTop}px;left:${at[1].left ?? page.left}px;right:${at[1].right ?? page.right}px;max-height:${Math.max(0, pageH(p) - hf.footerDist - fTop)}px;overflow:hidden`;
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
        const last = view.dom.lastElementChild;
        const endOnly = flow.items.every((r) => r.kind === 'endnote');
        // (In columns, a page of endnotes alone, after the text's last: from the top, its columns left to right whatever
        // the section's direction, as Word prints them; going on from the page before, under a continuation separator.)
        const alone = endOnly && !!flow.cols && !flow.own;
        if (alone) top = marginsAt(p).top + (flow.cont ? 0 : ENDNOTE_TOP);
        else if (last && !flow.skip && endOnly) {
          // (In columns balanced before them, under the lowest column: balanceEndCols.)
          const y = cols && colsBal ? marginsAt(p).top + colsBal.bottom + ENDNOTE_LEAD : last.getBoundingClientRect().bottom + (parseFloat(getComputedStyle(last).marginBottom) || 0) - view.dom.getBoundingClientRect().top - pageTop(p) + ENDNOTE_LEAD;
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
    for (const hf of layer.querySelectorAll('.odx-hf')) settleStoryShapes(hf); // (drawn now, so they can be measured)
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
    // (At full size a line can break a little differently: the words the hyphenation zone and limit leave whole, again.)
    settleHyphens(view);
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
  const afterPrint = () => { removePrintLines(sheet); printBox?.remove(); printBox = null; setGap(GAP); if (printZoom) sheet.style.zoom = printZoom; printZoom = ''; settleHyphens(view); };
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
