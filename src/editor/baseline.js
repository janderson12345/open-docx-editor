// Where the text sits in its line. Word puts a line's baseline at its top plus the font's ascender and line gap
// (hhea), unrounded; the browser rounds the font's ascent and descent to whole pixels and centres the leading,
// so its text sits up to a pixel and a half higher. For printing, each paragraph's text is moved down by the
// difference, measured once per font, size and line height. (Only the paint moves: pagination is unaffected.)
import { lineHeightOf } from './util.js';

// hhea ascender + line gap, per em, of the fonts Word documents use most.
const ASCENT = {
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

/** The browser's text box in a line: its top (px) from the line's top, and its height. */
function textBox(pcs, rcs, host) {
  const key = ['box', pcs.fontSize, pcs.fontFamily, lineHeightOf(pcs), rcs.font].join('|');
  if (cache.has(key)) return cache.get(key);
  if (!probe) {
    probe = document.createElement('div');
    probe.style.cssText = 'position:absolute;left:-9999px;top:0;visibility:hidden;white-space:nowrap';
  }
  host.append(probe);
  probe.style.font = `${pcs.fontSize} ${pcs.fontFamily}`;
  probe.style.lineHeight = '0';
  probe.innerHTML = '<span>X</span>';
  probe.firstChild.style.font = rcs.font;
  probe.firstChild.style.lineHeight = lineHeightOf(pcs);
  const r = probe.firstChild.getBoundingClientRect();
  const v = [r.top - probe.getBoundingClientRect().top, r.height];
  probe.remove();
  cache.set(key, v);
  return v;
}

/**
 * CSS for the print-only shift of each top-level paragraph whose first run is in a font we know.
 * `blocks`: [{ el, i }] in document order; `scope`: the editor's class.
 */
/** How far (px) a paragraph's printed text should move down to sit on Word's baseline, or 0. */
export function shiftFor(el, host, anchored = false) {
  // (A paragraph holding anchored pictures stays put, unless the caller moves them back: moving it would move them too.)
  if (el.tagName !== 'P' || (!anchored && el.querySelector('[style*="position: absolute"], [style*="position:absolute"]'))) return 0;
  const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.data.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP) });
  const text = walk.nextNode();
  if (!text) return 0;
  const rcs = getComputedStyle(text.parentElement);
  // After a leading bare break (its line in the smaller strut, doc 07), the text's line is its runs' own: measured as theirs.
  // (A leading break, not a leading run in bold or italic: their <strong>/<em> are runs like a span. Doc 08 edited to
  // open with a bold word printed 0.7px low.)
  const own = el.classList.contains('odx-markers') && !/^(SPAN|STRONG|EM|U|S|B|I|SUB|SUP|A)$/.test(el.firstElementChild?.tagName || 'SPAN');
  const pcs = own ? rcs : getComputedStyle(el);
  // The first family of the stack we know: "DejaVu Sans", Verdana prints (in Word too) in Verdana.
  const known = (cs) => cs.fontFamily.split(',').map((f) => f.replace(/["']/g, '').trim().toLowerCase()).find((f) => ASCENT[f]) || '';
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
  // Exact line spacing places text its own way in Word: leave those lines alone, but for one thing. A line a little
  // shorter than its text box has a negative half-leading, which the browser rounds down to a whole pixel: the text
  // jumps up by nearly 1px where Word moves it by the half (sample S085's 11.2pt-exact headings in 10pt Times, 1.3px high).
  if (parseFloat(pcs.getPropertyValue('--odx-runmult')) === 0) {
    const [top, box] = textBox(pcs, rcs, host);
    const h = parseFloat(lineHeightOf(pcs));
    const d = Math.round(((h - box) / 2 - top) * 100) / 100;
    return h < box && d > 0.1 && d < 1 ? d : 0;
  }
  // A line taller than its text: the browser puts half the extra height above the text. Word puts all of it above
  // for at-least spacing, and none for multiple spacing (double-spaced text sits at its line's top).
  const ncs = fs === parseFloat(rcs.fontSize) && family === known(rcs) ? rcs : pcs;
  const natural = (parseFloat(ncs.getPropertyValue('--odx-lh')) || 1.17) * fs;
  const room = Math.max(0, parseFloat(lineHeightOf(pcs)) - natural);
  const atLeast = pcs.getPropertyValue('--odx-lhlen').trim() === '1';
  const extra = atLeast ? room : 0;
  // (Trebuchet's offset was fitted on body text; its small header labels, 7.5pt in the sample S108/sample S109, sit right
  // without it.)
  // (Segoe UI's only in list paragraphs, where it was fitted: corpus 2's plain 9pt Segoe UI sat right without it,
  // 25 lines 1.4px low with it.)
  const off = (family === 'trebuchet ms' && fs < 12) || (family === 'segoe ui' && !el.classList.contains('odx-li')) ? 0 : OFFSET[family] || 0;
  const d = Math.round((asc * fs + extra - browserOffset(pcs, rcs, host) + off) * 100) / 100;
  // (A shift far from the expected half of the extra means something else is going on: left alone.)
  return Math.abs(d) < 0.05 || Math.abs(d - (atLeast ? room / 2 : -room / 2)) > 3 ? 0 : d;
}

/**
 * CSS for the print-only shift of each top-level paragraph whose first run is in a font we know.
 * `blocks`: [{ el, i }] in document order; `scope`: the editor's class.
 */
export function baselineCss(blocks, scope, host) {
  const rules = [];
  for (const { el, i } of blocks) {
    const d = shiftFor(el, host);
    if (d) rules.push(`.${scope}:not(.odx-bl-off) > .odx-content > :nth-child(${i + 1}){position:relative;top:${d}px}`);
  }
  // Paragraphs in table cells carry their shift as --odx-bl (see cellShifts), headers' and footers' too (hfShifts).
  rules.push(`.${scope}:not(.odx-bl-off) td > p[style*="--odx-bl"]{position:relative;top:var(--odx-bl)}`);
  rules.push(`.odx-hf p[style*="--odx-bl:"]{position:relative;top:var(--odx-bl)}`);
  // One holding anchored pictures is moved down by padding instead (position:relative would make it their containing
  // block), and they are moved back up: sample S062's footer "Page 3 of 9" beside its page art, 1.3px high.
  rules.push(`.odx-hf p[style*="--odx-bla"]{padding-top:var(--odx-bla);margin-bottom:calc(-1 * var(--odx-bla))}`);
  rules.push(`.odx-hf p[style*="--odx-bla"] [style*="position: absolute"]{translate:0 calc(-1 * var(--odx-bla))}`);
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
let cellEpoch = -1;
let fontLoads = 0;
if (typeof document !== 'undefined' && document.fonts) document.fonts.addEventListener('loadingdone', () => { fontLoads++; });
