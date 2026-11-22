// Where the text sits in its line. Word puts a line's baseline at its top plus the font's ascender and line gap
// (hhea), unrounded; the browser rounds the font's ascent and descent to whole pixels and centres the leading,
// so its text sits up to a pixel and a half higher. For printing, each paragraph's text is moved down by the
// difference, measured once per font, size and line height. (Only the paint moves: pagination is unaffected.)
import { lineHeightOf } from './util.js';

// hhea ascender + line gap, per em, of the fonts Word documents use most.
export const ASCENT = {
  arial: 0.938, 'liberation sans': 0.938, helvetica: 0.938, 'times new roman': 0.9336, 'liberation serif': 0.9336,
  calibri: 0.9707, 'calibri light': 0.9707, carlito: 0.9707, cambria: 0.9502, caladea: 0.9502, verdana: 1.0054,
  'dejavu sans': 1.0054, tahoma: 1.0005, 'segoe ui': 1.0791, georgia: 0.917, 'courier new': 0.8325, 'liberation mono': 0.8325,
  'trebuchet ms': 0.939, 'century schoolbook': 0.9858, century: 0.9858,
  aptos: 0.939, 'aptos narrow': 0.939, 'aptos display': 0.939, consolas: 0.9136,
  ubuntu: 0.96, 'ubuntu mono': 0.83, // (embedded in sample S125: its own hhea tables)
  'microsoft sans serif': 0.9219, // (sample S113. Not Comic Sans MS or Century Gothic: their hhea ascents put sample S038's lines low)
};
// What remains, per font, in px whatever the size: the average error of printed lines against Word's own PDFs, by font
// and size (tools/word-compare). Times New Roman sits 0.33-0.36px high at every size from 8 to 14pt, so it's a
// fixed amount, not a share of the em. Edge prints each line at a whole pixel, rounded to the nearest.
const OFFSET = {
  'times new roman': 0.34, 'liberation serif': 0.34, arial: 0.08, 'liberation sans': 0.08, helvetica: 0.08,
  calibri: 0, carlito: 0, cambria: 0.24, caladea: 0.24, 'trebuchet ms': -0.7, verdana: -0.6,
  aptos: 0.26, 'aptos narrow': 0.26, 'aptos display': 0.26, consolas: 0.58,
  // (Aptos refitted from 0.56: after later fixes its lines sat 0.25-0.31px low at 10-12pt. DejaVu Sans, set
  // in Verdana, apart from Verdana itself: its 10.5pt lines sat 0.81px high with Verdana's offset, Verdana's own at 0.)
  'dejavu sans': 0.2,
  // (Segoe UI, fitted on the "Open Sans;Arial" lists Word sets in it, 24 / file-sample's 10.5pt body: 1.3 and
  // 0.6px high in turn without it. List paragraphs only: shiftFor.)
  'segoe ui': 1.0,
  // (Arial and Calibri refitted: after later fixes their lines sat 0.15-0.28px low on average.)
};

// Word's own ascent where it isn't ASCENT's (the font's Windows ascent plus its external leading, GDI's): what a line of
// runs in several sizes or fonts is sized from (mixlines.js). Calibri's line is its Windows ascent and descent, 0.9521em
// above the baseline (Word's PDFs: 12pt Calibri raised 3pt in 10pt text, 17.11pt lines), where ASCENT has hhea's
// ascender plus line gap, which the print offsets below were fitted on.
const WORD_ASCENT = { calibri: 0.9521, 'calibri light': 0.9521, carlito: 0.9521, consolas: 0.9199 };
// (Looked up by the font a document names: no prototype, so "constructor" finds nothing.)
Object.setPrototypeOf(ASCENT, null);
Object.setPrototypeOf(OFFSET, null);
/** Word's ascent per em of a font ASCENT knows (its lowercase family name). */
export const wordAscent = (family) => (Object.hasOwn(WORD_ASCENT, family) ? WORD_ASCENT[family] : ASCENT[family]);
/** The first family of a computed style's stack that ASCENT knows, lowercase, or ''. */
export const knownFamily = (cs) => cs.fontFamily.split(',').map((f) => f.replace(/["']/g, '').trim().toLowerCase()).find((f) => Object.hasOwn(ASCENT, f)) || '';
/** How far (px) a span's text is raised in its paragraph `el`: its wrappers' vertical-align (not its own), or null when one isn't a length. */
export function raiseOf(span, el) {
  let r = 0;
  for (let n = span.parentElement; n && n !== el; n = n.parentElement) {
    const va = getComputedStyle(n).verticalAlign;
    if (va === 'baseline') continue;
    const v = parseFloat(va);
    if (!Number.isFinite(v) || !va.endsWith('px')) return null;
    r += v;
  }
  return r;
}
/** A font's print offset (OFFSET), with its exceptions: Trebuchet's small text, Segoe UI's outside lists. */
const offsetFor = (family, fs, el) => ((family === 'trebuchet ms' && fs < 12) || (family === 'segoe ui' && !el.classList.contains('odx-li')) ? 0 : OFFSET[family] || 0);

/**
 * The print shift (px) of a paragraph whose text spans have Word's boxes (mixlines.js, data-odx-mix): its baseline is
 * already Word's, so only the difference between the ascent the offsets were fitted from (ASCENT) and Word's own, and
 * the font's offset, of the run with the highest top on its first line. `first`: its first text node.
 */
function mixedShift(el, first) {
  const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  walk.currentNode = first;
  let best = null, most = -Infinity, line = null;
  const r = document.createRange();
  for (let t = first, n = 0; t && n < 200; t = walk.nextNode(), n++) {
    if (!t.data.trim()) continue;
    const span = t.parentElement;
    if (span.closest('sup, sub')) continue;
    const cs = getComputedStyle(span);
    const family = knownFamily(cs);
    if (!family) continue;
    const fs = parseFloat(cs.fontSize);
    const top = wordAscent(family) * fs + (raiseOf(span, el) || 0);
    if (!line) { r.selectNodeContents(first); line = r.getClientRects()[0]; if (!line) return 0; }
    if (t !== first) {
      r.selectNodeContents(t);
      const at = r.getClientRects()[0];
      if (!at || at.top >= line.bottom) break; // (on a later line: so is every run after it)
    }
    if (top > most) { most = top; best = { family, fs }; }
  }
  if (!best) return 0;
  const d = Math.round(((ASCENT[best.family] - wordAscent(best.family)) * best.fs + offsetFor(best.family, best.fs, el)) * 100) / 100;
  return Math.abs(d) < 0.05 ? 0 : d;
}

const cache = new Map();
let probe = null;

/** The browser's baseline offset (px) from a line's top, for a paragraph style and its first run's font. */
function browserOffset(pcs, rcs, host) {
  const key = [pcs.fontSize, pcs.fontFamily, lineHeightOf(pcs), rcs.font].join('|');
  if (cache.has(key)) return cache.get(key);
  if (!probe) {
    probe = document.createElement('div');
    probe.style.cssText = 'position:absolute;left:-9999px;top:0;visibility:hidden;white-space:nowrap';
  }
  host.append(probe);
  probe.style.font = `${pcs.fontSize} ${pcs.fontFamily}`;
  probe.style.lineHeight = '0'; // (only the run's line counts, as a paragraph's: textspans.js)
  probe.innerHTML = '<span>X<i style="display:inline-block;width:0;height:0"></i></span>';
  probe.firstChild.style.font = rcs.font;
  probe.firstChild.style.lineHeight = lineHeightOf(pcs);
  const v = probe.querySelector('i').getBoundingClientRect().top - probe.getBoundingClientRect().top;
  probe.remove();
  cache.set(key, v);
  return v;
}

/**
 * CSS for the print-only shift of each top-level paragraph whose first run is in a font we know.
 * `blocks`: [{ el, i }] in document order; `scope`: the editor's class.
 */
/**
 * The run on a paragraph's first line whose ascent (Word's: hhea, plus any raise) passes `have` (the first run's) by
 * more than half a pixel, the tallest: { cs, family }, or null. Only runs that could pass it are looked for on the line
 * (reading where text sits lays the paragraph out: a long document skips that off screen).
 */
function tallest(el, first, have, known) {
  const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.data.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP) });
  walk.currentNode = first;
  let best = null;
  let most = have + 0.5;
  let line = null;
  // (The first 200 runs at most: a line holds far fewer.)
  for (let t = walk.nextNode(), n = 0; t && n < 200; t = walk.nextNode(), n++) {
    const host = t.parentElement;
    if (host.closest('sup, sub, math')) continue;
    const cs = getComputedStyle(host);
    const family = known(cs);
    if (!family) continue;
    const va = parseFloat(cs.verticalAlign);
    const a = ASCENT[family] * parseFloat(cs.fontSize) + (Number.isFinite(va) ? va : 0);
    if (!(a > most)) continue;
    const r = document.createRange();
    if (!line) { r.selectNodeContents(first); line = r.getClientRects()[0]; if (!line) return null; }
    r.selectNodeContents(t);
    const at = r.getClientRects()[0];
    if (!at || at.top >= line.bottom) break; // (on a later line: so is every run after it)
    most = a;
    best = { cs, family };
  }
  return best;
}

/** How far (px) a paragraph's printed text should move down to sit on Word's baseline, or 0. */
export function shiftFor(el, host, anchored = false) {
  // (A paragraph holding anchored pictures stays put, unless the caller moves them back: moving it would move them too.)
  if (el.tagName !== 'P' || (!anchored && el.querySelector('[style*="position: absolute"], [style*="position:absolute"]'))) return 0;
  const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.data.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP) });
  const text = walk.nextNode();
  if (!text) return 0;
  if (el.dataset.odxMix === '1') return mixedShift(el, text);
  let rcs = getComputedStyle(text.parentElement);
  // After a leading bare break (its line in the smaller strut, doc 07), the text's line is its runs' own: measured as theirs.
  // (A leading break, not a leading run in bold or italic: their <strong>/<em> are runs like a span. Doc 08 edited to
  // open with a bold word printed 0.7px low.)
  const own = el.classList.contains('odx-markers') && !/^(SPAN|STRONG|EM|U|S|B|I|SUB|SUP|A)$/.test(el.firstElementChild?.tagName || 'SPAN');
  let pcs = own ? rcs : getComputedStyle(el);
  // (And where none of its text is in the paragraph's own size, a run among them leaving its size unset: its lines are
  // its runs' own, as tall as their sizes make them, the paragraph's size in none of them. A list paragraph of 12pt
  // runs and one 10pt character in an 11pt style, measured as an 11pt line holding a 12pt run, printed its four lines
  // 1px low. Runs that need Word's boxes are sized and shifted together: mixlines.js, mixedShift above.)
  // (Larger than the paragraph's size, that is: smaller runs share their line with what is left in the paragraph's
  // own size, an unsized space or its mark, and Word's line is that size's. Cells of 10pt runs and one unsized space in
  // an 11pt style, measured as 10pt lines, printed 0.7px lower, 1.3px under Word's.)
  if (!own && rcs.fontSize !== pcs.fontSize && parseFloat(rcs.fontSize) > parseFloat(pcs.fontSize) && el.dataset.odxMix !== '1') {
    // (On its first line, that is, the line this shift is worked out for: text in the paragraph's size further down
    // takes no part in it. A list paragraph of 12pt runs with five 11pt characters on its third line was measured as
    // an 11pt line again, and its lines after the first printed 1px low.)
    let none = true;
    const first = document.createRange();
    first.selectNodeContents(text);
    const line = first.getClientRects()[0];
    for (const s of el.querySelectorAll('.odx-t')) {
      if (!s.textContent.trim()) continue;
      if (line) { const q = s.getClientRects()[0]; if (q && q.top >= line.bottom) break; }
      if (getComputedStyle(s).fontSize === pcs.fontSize) { none = false; break; }
    }
    if (none) pcs = rcs;
  }
  // The first family of the stack we know: "DejaVu Sans", Verdana prints (in Word too) in Verdana.
  const known = knownFamily;
  let family = known(rcs);
  let asc = ASCENT[family];
  if (!asc || lineHeightOf(pcs) === 'normal') return 0;
  let fs = parseFloat(rcs.fontSize);
  // Text in the paragraph's own font with a taller ascent than the first run's sets the line's baseline (sample S099's 8pt
  // Calibri labels before their 10pt Times values, printed 2.7px high measured from the labels).
  const pfam = !own && pcs.fontFamily !== rcs.fontFamily ? known(pcs) : '';
  if (pfam && ASCENT[pfam] * parseFloat(pcs.fontSize) > asc * fs + 0.5) {
    const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    for (let t = w.nextNode(); t; t = w.nextNode()) {
      const cs = t.data.trim() && getComputedStyle(t.parentElement);
      if (cs && cs.fontFamily === pcs.fontFamily && cs.fontSize === pcs.fontSize) { family = pfam; asc = ASCENT[pfam]; fs = parseFloat(pcs.fontSize); break; }
    }
  }
  // Exact line spacing: Word puts the baseline four fifths of the way down the line, whatever the font, its size or the
  // line's height, shorter than the text or taller (Word's PDFs of exactly 10 to 60pt lines in Times New Roman, Calibri,
  // Aptos and Arial at 11 and 12pt: 0.800-0.806 of the line, within Word's 0.12pt grid). The browser centres the text's
  // box in the line instead: 12pt Aptos at exactly 20pt printed 4px high, at 60pt 19px.
  if (parseFloat(pcs.getPropertyValue('--odx-runmult')) === 0) {
    const h = parseFloat(lineHeightOf(pcs));
    if (!(h > 0) || family !== known(rcs)) return 0;
    // (Plus 0.1px for every font, not the fonts' offsets below: those were fitted where Word sets the text on its ascent,
    // and left exact lines in Times New Roman 0.25px low, Aptos 0.13px; with 0.1px all four sit within 0.05px on average.)
    const e = Math.round((0.8 * h - browserOffset(pcs, rcs, host) + 0.1) * 100) / 100;
    // (A shift beyond where the browser's centred text can be, half the line's height off either way: left alone.)
    return Math.abs(e) >= 0.05 && Math.abs(e) < h / 2 ? e : 0;
  }
  // A larger run later in the first line sets its baseline, in Word as in the browser: the shift is that run's (sample
  // S078's footer, 10pt "ITP Template" before 12pt bold "Company Proprietary" raised 3pt, printed 0.76px high; Word-made
  // footers and body lines of 10pt Times New Roman with a 12pt run, raised, lowered or not, 1.8-2.8px high).
  const top = tallest(el, text, asc * fs, known);
  if (top) { rcs = pcs = top.cs; family = top.family; asc = ASCENT[family]; fs = parseFloat(rcs.fontSize); }
  // A line taller than its text: the browser puts half the extra height above the text. Word puts all of it above
  // for at-least spacing, and none for multiple spacing (double-spaced text sits at its line's top).
  // A picture in the first line taller than its text: the line's top is the picture's, its baseline the picture's foot,
  // in the browser as in Word, with no extra spacing above the text to move (a 0.33in picture in a double-spaced line
  // printed 8px high: the generated documents, tools/doc-generator).
  // (Measured only where there's a picture: reading a paragraph's text box lays it out, and a long document skips laying
  // out paragraphs far off screen, content-visibility: 6000 tracked changes' paragraphs took 8s to open.)
  // (Trebuchet's offset was fitted on body text; its small header labels, 7.5pt in the sample S108/sample S109, sit right
  // without it.)
  // (Segoe UI's only in list paragraphs, where it was fitted: corpus 2's plain 9pt Segoe UI sat right without it,
  // 25 lines 1.4px low with it.)
  const off = offsetFor(family, fs, el);
  const imgs = el.querySelectorAll('img');
  const tr = imgs.length ? text.parentElement.getBoundingClientRect() : null;
  for (const img of imgs) {
    if (/position:\s*absolute|float:/.test(img.getAttribute('style') || '')) continue;
    const r = img.getBoundingClientRect();
    // (Its text still drawn the font's offset high, as every line's: picmatrix-word's picture lines 0.2pt high in Aptos.)
    if (r.height > tr.height + 1 && r.bottom > tr.top && r.top < tr.bottom) return el.classList.contains('odx-pictext') && Math.abs(off) >= 0.05 ? off : 0;
  }
  const ncs = fs === parseFloat(rcs.fontSize) && family === known(rcs) ? rcs : pcs;
  const natural = (parseFloat(ncs.getPropertyValue('--odx-lh')) || 1.17) * fs;
  const room = Math.max(0, parseFloat(lineHeightOf(pcs)) - natural);
  const atLeast = pcs.getPropertyValue('--odx-lhlen').trim() === '1';
  // (A line on a document grid has its text centred, as the browser sets it: lhfix.js.)
  const grid = el.dataset?.odxGrid === '1';
  const extra = atLeast ? room : grid ? room / 2 : 0;
  const d = Math.round((asc * fs + extra - browserOffset(pcs, rcs, host) + off) * 100) / 100;
  // (A shift far from the expected half of the extra means something else is going on: left alone.)
  return Math.abs(d) < 0.05 || (!grid && Math.abs(d - (atLeast ? room / 2 : -room / 2)) > 3) || (grid && Math.abs(d) > 3) ? 0 : d;
}

/**
 * The print-only shift of each top-level paragraph whose first run is in a font we know, as a custom property for the
 * block itself (block index -> its style; pages.js puts it on the block, a node decoration): --odx-pshift, or
 * --odx-pshifta for one holding anchored pictures (its text moved, not the paragraph: position:relative would make it
 * their containing block, and move them; rotpic-word's text beside a floating picture at 1.16 printed 0.7px low,
 * unshifted). (Not a rule each on its place among its siblings, :nth-child: those are tried on every element at each
 * restyle of the print, and of each page's copy of the sheet (printcopies.js); a 535-page document's 2,500 of them kept
 * its print from finishing in 10 minutes.) `blocks`: [{ el, i, node }] in document order.
 */
export function blockShifts(blocks, host) {
  const out = new Map();
  // (An unchanged paragraph keeps its shift, as cellShifts' do: reading every paragraph's formatting again each pass
  // took 80ms a keystroke in a 750-page document.)
  if (blockEpoch !== fontLoads) { blockMemo = new WeakMap(); blockEpoch = fontLoads; }
  for (const { el, i, node } of blocks) {
    const anchored = el.tagName === 'P' && !!el.querySelector(':scope > [style*="position: absolute"], :scope > [style*="position:absolute"]');
    let d = node ? blockMemo.get(node) : undefined;
    if (d === undefined) { d = shiftFor(el, host, anchored); if (node) blockMemo.set(node, d); }
    if (d) out.set(i, `${anchored ? '--odx-pshifta' : '--odx-pshift'}:${d}px`);
  }
  return out;
}

/** The print-only baseline shifts' rules (the top-level paragraphs' from blockShifts, cells', headers' and footers'). */
export function baselineCss(scope) {
  const rules = [];
  const on = `.${scope}:not(.odx-bl-off) > .odx-content > `;
  rules.push(`${on}[style*="--odx-pshifta:"] > :not([style*="position: absolute"], [style*="position:absolute"], img[style*="float:"], img.odx-wrap-block){position:relative;top:var(--odx-pshifta)}`);
  rules.push(`${on}[style*="--odx-pshift:"]{position:relative;top:var(--odx-pshift)}`);
  // Its floating pictures (wrapped round on one side, or top and bottom) drawn back where they were: the shift is the
  // text's, Word draws them at their offsets (wrap-word5 sb-sqright, sb-tb: rectangles 0.8pt lower than Word's, as
  // much as the text's shift, where those not in the flow were at their place).
  rules.push(`${on}p[style*="--odx-pshift:"] > :is(img[style*="float:"], img.odx-wrap-block){translate:0 calc(-1 * var(--odx-pshift))}`);
  // Paragraphs in table cells carry their shift as --odx-bl (see cellShifts), headers' and footers' too (hfShifts).
  rules.push(`.${scope}:not(.odx-bl-off) td > p[style*="--odx-bl"]{position:relative;top:var(--odx-bl)}`);
  rules.push(`.odx-hf p[style*="--odx-bl:"]{position:relative;top:var(--odx-bl)}`);
  // One holding anchored pictures is moved down by padding instead (position:relative would make it their containing
  // block), and they are moved back up: sample S062's footer "Page 3 of 9" beside its page art, 1.3px high.
  rules.push(`.odx-hf p[style*="--odx-bla"]{padding-top:var(--odx-bla);margin-bottom:calc(-1 * var(--odx-bla))}`);
  // (Not one set on the page from its top or its margins, hfstory.js placeStoryShapes: its place is the page's, which
  // the padding doesn't move. Word's prints of six generated footers, a text box 9.5in down the page beside the
  // footer's line: its text 1.7px high with this taken off it too.)
  rules.push(`.odx-hf p[style*="--odx-bla"] [style*="position: absolute"]:not([style*="--odx-arel"], [style*="--odx-lrel"]){translate:0 calc(-1 * var(--odx-bla))}`);
  return `@media print{${rules.join('')}}`;
}

/** Sets --odx-bl on the header and footer paragraphs under `root` (drawn, so they can be measured). */
export function hfShifts(root, host) {
  for (const p of root.querySelectorAll('.odx-hf p')) {
    const anchored = !!p.querySelector('[style*="position: absolute"], [style*="position:absolute"]');
    // (Padding of its own, a border's, would be lost: left alone.)
    const d = shiftFor(p, host, anchored && !parseFloat(getComputedStyle(p).paddingTop));
    if (d) p.style.setProperty(anchored ? '--odx-bla' : '--odx-bl', `${d}px`);
  }
}

/** [{ pos, size, d }] for the paragraphs in table cells: positions in `doc` and their shifts. */
export function cellShifts(view, host) {
  const out = [];
  if (cellEpoch !== fontLoads) { cellMemo = new WeakMap(); cellEpoch = fontLoads; }
  view.state.doc.descendants((node, pos, parent) => {
    if (node.type.name !== 'paragraph') return true;
    if (parent && /cell|header/.test(parent.type.name)) {
      // An unchanged paragraph in an unchanged cell (its computed formatting is in the cell's attrs) keeps its shift.
      let d = cellMemo.get(parent)?.get(node);
      if (d === undefined) {
        const el = view.nodeDOM(pos);
        if (!el || el.nodeType !== 1) return false;
        d = shiftFor(el, host);
        if (!cellMemo.has(parent)) cellMemo.set(parent, new WeakMap());
        cellMemo.get(parent).set(node, d);
      }
      if (d) out.push({ pos, size: node.nodeSize, d });
    }
    return false;
  });
  return out;
}
// Cell paragraphs' shifts by cell and paragraph node, until fonts load (measuring each one again on every layout pass
// was a long document's largest cost after the pages' restyling).
let cellMemo = new WeakMap();
let blockMemo = new WeakMap(); // top-level paragraph node -> its shift
let blockEpoch = -1;
let cellEpoch = -1;
let fontLoads = 0;
if (typeof document !== 'undefined' && document.fonts) document.fonts.addEventListener('loadingdone', () => { fontLoads++; });
