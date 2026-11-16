// Footnotes and endnotes as pagination measures and draws them (pages.js): each note's element, its height at a text
// width and where it may be cut (its paragraphs' lines, its table's rows), a band of notes laid out in columns and
// balanced as Word balances them, and the side the separator line goes. Measured off-screen on the sheet, and kept
// until the headers, footers or notes change (clear()).
import { DOMSerializer } from 'prosemirror-model';
import { schema } from './schema.js';
import { styleStoryTables } from './hfstory.js';
import { wrapText } from './textspans.js';
import { noteLabel } from './notes.js';
import { startTags, tagAttr } from '../ooxml/xml.js';

const serializer = DOMSerializer.fromSchema(schema);
// A note paragraph's table: the element itself or its direct child.
const ownTable = (el) => (el.tagName === 'TABLE' ? el : el.querySelector(':scope > table'));

/** The note boxes for one pagination: `page` its first section's page, `sheet` the element they're measured in. */
export function createNoteBoxes({ page, sheet, session }) {
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
  // The footnote separator goes at the text's right edge when its paragraph is right to left (its own w:bidi, or the
  // document's default paragraph's), as in Word; at the left otherwise, whatever the section's direction.
  let sepRtl = null;
  const noteSepRtl = () => {
    if (sepRtl == null) {
      const bidi = (x) => /<w:bidi(?=[\s/>])(?![^>]*w:val="(?:0|false|off)")/.test(x || '');
      const part = session.pkg?.partByType?.('http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes');
      const sep = part && /<w:footnote\b[^>]*w:type="separator"[\s\S]*?<\/w:footnote>/.exec(session.pkg.text(part) || '')?.[0];
      const ppr = sep && /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(sep)?.[0];
      const stylesPart = session.pkg?.partByType?.('http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles');
      const styles = stylesPart ? session.pkg.text(stylesPart) || '' : '';
      const defaults = /<w:pPrDefault>[\s\S]*?<\/w:pPrDefault>/.exec(styles)?.[0];
      // (The default paragraph style's element, read tag by tag: xml.js startTag.)
      let normal = null;
      for (const t of startTags(styles, 'w:style')) {
        if (tagAttr(t.tag, 'w:type') !== 'paragraph' || tagAttr(t.tag, 'w:default') !== '1') continue;
        const close = styles.indexOf('</w:style>', t.end);
        if (close >= 0) normal = styles.slice(t.at, close + 10);
        break;
      }
      const normalPpr = normal && /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(normal)?.[0];
      sepRtl = !!sep && (ppr && /<w:bidi\b/.test(ppr) ? bidi(ppr) : normalPpr && /<w:bidi\b/.test(normalPpr) ? bidi(normalPpr) : bidi(defaults));
    }
    return sepRtl;
  };
  const noteBounds = new Map(); // note doc -> where each of its paragraphs ends, px from its top
  const noteInks = new Map(); // note doc -> where the text of each of those lines ends, px from its top
  // (At `w`, a page's text width when it isn't the first section's: a landscape page's notes are wider, so shorter,
  // longfn-word2 fnland.)
  const noteHeight = (note, kind, n, w = null) => {
    const key = w == null ? '' : String(Math.round(w));
    if (!noteHeights.has(note.doc)) { noteHeights.set(note.doc, new Map()); noteBounds.set(note.doc, new Map()); noteInks.set(note.doc, new Map()); }
    if (!noteHeights.get(note.doc).has(key)) {
      const box = document.createElement('div');
      box.className = 'odx-notes';
      box.style.cssText = `visibility:hidden;top:0;left:${page.left}px;${w == null ? `right:${page.right}px` : `width:${w}px`}`;
      const el = noteEl(note, kind, n);
      box.append(el);
      sheet.append(box);
      const top = box.getBoundingClientRect().top;
      noteHeights.get(note.doc).set(key, box.getBoundingClientRect().height);
      // (A table's rows each end a place it can be cut too: a note holding a long table goes on on the next page from
      // the row that doesn't fit, as in Word.)
      // (Each with where its line's text ends, `ink`: the browser puts half a line's leading under its text, Word all of it
      // above, so Word fits a cut whose last line's text ends on the margin, longfn-word2 fnland page 4.)
      const cuts = [...el.children].flatMap((c) => {
        const t = ownTable(c);
        const rows = t && t.tBodies[0] ? [...t.tBodies[0].rows] : [];
        if (rows.length) return rows.map((e) => { const y = e.getBoundingClientRect().bottom - top; return { y, ink: y }; });
        // (And a paragraph's every line: a note longer than its page goes on on the next from the first line that doesn't
        // fit, as in Word, not from mid-line (longfn-word fnhuge).)
        const rg = document.createRange();
        rg.selectNodeContents(c);
        const rects = [...rg.getClientRects()].filter((q) => q.height > 0);
        // (Swept in order once, line ends and rects together: a note paragraph of thousands of lines costs its length,
        // not lines times rects.)
        const tops = [];
        for (const t of rects.map((q) => q.top).sort((x, y) => x - y)) if (!tops.length || t - tops[tops.length - 1] >= 2) tops.push(t);
        const ends = [...tops.slice(1), c.getBoundingClientRect().bottom];
        const byMid = rects.map((q) => ({ mid: (q.top + q.bottom) / 2, bottom: q.bottom })).sort((x, y) => x.mid - y.mid);
        let k = 0;
        return ends.map((e, i) => {
          const from = i ? ends[i - 1] : -Infinity;
          while (k < byMid.length && byMid[k].mid <= from) k++;
          let ink = -Infinity;
          for (; k < byMid.length && byMid[k].mid < e; k++) if (byMid[k].mid > from) ink = Math.max(ink, byMid[k].bottom);
          return { y: e - top, ink: Math.min(e, ink) - top };
        });
      }).filter((b) => b.y > 0);
      noteBounds.get(note.doc).set(key, cuts.map((b) => b.y));
      noteInks.get(note.doc).set(key, cuts.map((b) => (Number.isFinite(b.ink) ? b.ink : b.y)));
      box.remove();
    }
    return noteHeights.get(note.doc).get(key);
  };
  // A band's notes in its columns: laid out once at the column's width, one under another, with where a column may end
  // (bandFill takes the columns' stretches of that). (The browser's own balanced columns ignore widows and orphans.)
  // { total, w, breaks: [px], inks: [px], notes: [[top, bottom] px] }.
  const bandLays = new Map(); // key -> layout
  const noteDocIds = new WeakMap();
  let noteDocSeq = 0;
  const bandWidth = (bc) => bc.w ?? (page.width - page.left - page.right - bc.gap * (bc.m - 1)) / bc.m;
  const bandLayout = (items, bc) => {
    const key = `${bc.m}/${bc.gap}/${bc.w ?? ''}/${page.width - page.left - page.right}:${items.map((r) => { if (!noteDocIds.has(r.note.doc)) noteDocIds.set(r.note.doc, ++noteDocSeq); return `${noteDocIds.get(r.note.doc)}.${r.n}`; }).join(',')}`;
    if (bandLays.has(key)) return bandLays.get(key);
    const w = bandWidth(bc);
    const box = document.createElement('div');
    box.className = 'odx-notes';
    box.style.cssText = `visibility:hidden;top:0;left:${page.left}px;width:${w}px`;
    const els = items.map((r) => noteEl(r.note, r.kind, r.n));
    box.append(...els);
    sheet.append(box);
    const top0 = box.getBoundingClientRect().top;
    const total = box.getBoundingClientRect().height;
    const range = document.createRange();
    // Where a column may end: between notes and between a note's paragraphs, at a table's rows, and inside a paragraph
    // of n lines after its line i when 2 <= i <= n - 2.
    const breaks = [];
    const spans = []; // each note's [top, bottom]
    const lines = []; // every line's text, [mid, bottom]: where the text above a break ends (its ink)
    els.forEach((el, j) => {
      const r = el.getBoundingClientRect();
      spans.push([r.top - top0, r.bottom - top0]);
      if (j) breaks.push(r.top - top0);
      [...el.children].forEach((c, ci) => {
        if (ci) breaks.push(c.getBoundingClientRect().top - top0);
        const t = ownTable(c);
        if (t && t.tBodies[0]) {
          [...t.tBodies[0].rows].forEach((row, ri) => { const q = row.getBoundingClientRect(); if (ri) breaks.push(q.top - top0); lines.push([(q.top + q.bottom) / 2 - top0, q.bottom - top0]); });
          return;
        }
        const tops = [];
        const walk = document.createTreeWalker(c, NodeFilter.SHOW_TEXT);
        for (let tn = walk.nextNode(); tn; tn = walk.nextNode()) {
          if (!tn.data.length) continue;
          range.selectNodeContents(tn);
          for (const q of range.getClientRects()) {
            if (!q.height) continue;
            lines.push([(q.top + q.bottom) / 2 - top0, q.bottom - top0]);
            if (!tops.some((y) => Math.abs(y - q.top) < 2)) tops.push(q.top);
          }
        }
        tops.sort((x, y) => x - y);
        for (let i = 2; i <= tops.length - 2; i++) breaks.push(tops[i] - top0);
      });
    });
    box.remove();
    breaks.sort((x, y) => x - y);
    // (Each break's ink: the lowest text above it. Swept once, lines and breaks in order.)
    lines.sort((x, y) => x[0] - y[0]);
    const ink = (ys) => { let k = 0, hi = 0; return ys.map((y) => { for (; k < lines.length && lines[k][0] < y; k++) hi = Math.max(hi, lines[k][1]); return Math.min(y, hi || y); }); };
    const [inks, totalInk] = [ink(breaks), ink([total])[0]];
    const lay = { total, totalInk, w, breaks, inks, notes: spans, fills: new Map() };
    bandLays.set(key, lay);
    return lay;
  };
  // The band's columns from `from` px into its notes (where a note carried on from the page before goes on), as Word
  // fills them: balanced, the lowest band that holds them all, the first column taking the extra, a note split between
  // columns only where each part keeps two lines (so a note of three lines never is); or, when that is taller than
  // `room`, each column filled to the room and the rest left for the next page's band (a column fits when its text ends
  // in the room: the browser puts half a line's leading under it, Word above). { h, spans: [[from, to] px], end }.
  const bandFill = (lay, m, from = 0, room = Infinity) => {
    const fkey = `${m}:${from.toFixed(2)}:${room === Infinity ? '' : room.toFixed(2)}`;
    if (lay.fills.has(fkey)) return lay.fills.get(fkey);
    const { total, breaks, inks } = lay;
    const fill = (H, part) => { // the columns' stretches at height H; part: as far as they go, else null when they don't hold it all
      // (Whether break k, -1 the end, ends a column from x: held to the room, by its text's end.)
      const fits = (x, k) => (k < 0 ? (part ? lay.totalInk : total) : part ? inks[k] : breaks[k]) - x <= H + 0.5;
      const spans = [];
      let at = from;
      for (let c = 0; c < m; c++) {
        if (fits(at, -1)) { spans.push([at, total]); return spans; }
        let to = null;
        for (let k = 0; k < breaks.length; k++) if (breaks[k] > at + 0.5 && fits(at, k)) to = breaks[k];
        // (Nowhere to end in the room, the first column: cut at the room, as one column cuts a paragraph taller than it.)
        if (to == null && part && !c) to = Math.min(total, at + H);
        if (to == null) return part ? spans : null;
        spans.push([at, to]);
        at = to;
      }
      return part ? spans : null;
    };
    // The heights a column can come to: from any place it can start (`from`, a break) to any place it can end.
    const starts = [from, ...breaks.filter((y) => y > from + 0.5)], ends = [...breaks, total].filter((y) => y > from + 0.5);
    const cands = [...new Set(starts.flatMap((x) => ends.filter((y) => y > x).map((y) => +(y - x).toFixed(2))))].filter((h) => h >= (total - from) / m - 1 && h <= room + 0.5).sort((x, y) => x - y);
    let cols = null;
    for (const H of cands) if ((cols = fill(H))) break;
    if (!cols) cols = room === Infinity ? [[from, total]] : fill(room, true);
    if (!cols.length) cols = [[from, from]];
    const out = { h: Math.max(...cols.map(([x, y]) => y - x)), spans: cols, end: cols[cols.length - 1][1] };
    lay.fills.set(fkey, out);
    return out;
  };
  // The band drawn: each column its stretch of the notes (bandFill; right to left in a right-to-left section).
  const bandBox = (items, bc, lay, fit) => {
    const out = document.createElement('div');
    out.className = 'odx-note-band';
    out.style.cssText = `position:relative;height:${fit.h}px`;
    const inner = page.width - page.left - page.right;
    fit.spans.forEach(([from, to], c) => {
      const x = c * (lay.w + bc.gap);
      const col = document.createElement('div');
      col.style.cssText = `position:absolute;top:0;left:${bc.rtl ? inner - x - lay.w : x}px;width:${lay.w}px;height:${to - from}px;overflow:hidden`;
      // Only the notes it shows (a note split between columns in both): copies hidden above or below had taken clicks
      // on the body text and been found by Find and screen readers.
      const shown = items.map((r, j) => [r, j]).filter(([, j]) => lay.notes[j][0] < to - 0.5 && lay.notes[j][1] > from + 0.5);
      const run = document.createElement('div');
      run.style.marginTop = `${-(from - (shown.length ? lay.notes[shown[0][1]][0] : from))}px`;
      run.append(...shown.map(([r]) => noteEl(r.note, r.kind, r.n)));
      col.append(run);
      out.append(col);
    });
    return out;
  };

  return {
    noteEl, noteSepRtl, noteHeight, bandLayout, bandFill, bandBox,
    /** Where a note measured at width key `key` may be cut (px from its top), or undefined before it's measured. */
    boundsOf: (doc, key) => noteBounds.get(doc)?.get(key),
    /** Where the text of the line ending at each of those cuts ends (px from the note's top). */
    inksOf: (doc, key) => noteInks.get(doc)?.get(key),
    /** Everything measured forgotten: the notes or the styles they're drawn in changed. */
    clear() { noteHeights.clear(); bandLays.clear(); sepRtl = null; },
  };
}
