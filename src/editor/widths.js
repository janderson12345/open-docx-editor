// Text widths for fonts Edge sets with hinted advances. Edge lays some fonts out at small sizes with their hinted
// (device) advances, not the scaled design ones Word uses: Cambria 10pt comes out 3.6% narrower than its design widths,
// 12pt 1.9% wider, 11pt right, and Word's lines take fewer or more words (corpus 2's justified
// Cambria paragraph a line shorter). Each run in such a font and size gets the difference back as letter spacing: its
// text's width at 100px, scaled, against its width at its size, shared out per character (on top of the run's own
// character spacing, --odx-csp).
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { fitOf } from './fittext.js';
import { ASCENT, wordAscent } from './baseline.js';
import { lineFactor } from '../ooxml/props.js';

const key = new PluginKey('odx-widths');
export const widthsPlugin = new Plugin({
  key,
  state: {
    init: () => DecorationSet.empty,
    apply(tr, set) {
      const next = tr.getMeta(key);
      return next ? DecorationSet.create(tr.doc, next) : set.map(tr.mapping, tr.doc);
    },
  },
  props: { decorations: (state) => key.getState(state) },
});

const REF = 'The quick brown fox jumps over the lazy dog while seven stored procedures were rewritten';
const hinted = new Map(); // font (style weight family) + size -> true when Edge's widths there aren't the scaled ones
// Text measured as the page lays it out, in a span of its own. (Not a canvas: Firefox's canvas measures text with
// hinted widths where its page uses the design ones, which made runs it lays out right look wrong.)
let probe = null;
function width(cs, px, text, caps = 'normal') {
  if (!probe || !probe.isConnected) {
    probe = document.createElement('span');
    probe.setAttribute('aria-hidden', 'true');
    probe.style.cssText = 'position:absolute;left:-99999px;top:0;visibility:hidden;white-space:pre;letter-spacing:0;word-spacing:0;font-kerning:none;font-variant-ligatures:none;line-height:normal';
    document.body.append(probe);
  }
  probe.style.font = `${cs.fontStyle} ${cs.fontWeight} ${px}px ${cs.fontFamily}`;
  probe.style.fontVariantCaps = caps;
  probe.textContent = text;
  return probe.getBoundingClientRect().width;
}
function isHinted(cs) {
  const k = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontFamily} ${cs.fontSize}`;
  if (!hinted.has(k)) {
    const px = parseFloat(cs.fontSize);
    const at = width(cs, px, REF);
    const scaled = (width(cs, 100, REF) * px) / 100;
    // (Edge's own cut of the size to 1/100px, sizeCss's business, is far under this.)
    hinted.set(k, Math.abs(at / scaled - 1) > 0.003);
  }
  return hinted.get(k);
}

// Bold over the Microsoft 365 cloud SemiBold and Demi families. They have no bold face (open-metric-fonts' stand-ins
// neither: omf.js gives each a 400 face only), so Word fakes the Bold, and its PDFs show the faked text at the Regular's
// advances: the same line breaks and line ends to 0.01pt (Word COM, 8 to 72pt, compatibility modes 12 and 15;
// docs/LAYOUT.md). Edge's faked bold keeps the widths too; Firefox's widens every character (0.42px at 10.67px, 2px at
// 96px), which broke each bold line a word or two early. Such a run is measured at weight 400 and given the difference.
const FAKE_BOLD = new Set(['aptos semibold', 'avenir next lt pro demi', 'georgia pro semibold', 'verdana pro semibold']);
const fakeBold = (cs) => (parseInt(cs.fontWeight, 10) || 400) >= 600
  && FAKE_BOLD.has(cs.fontFamily.split(',')[0].replace(/["']/g, '').trim().toLowerCase());

// A run in another font than its paragraph's (Cambria or Courier New words in a Calibri paragraph). Word makes a line
// as tall as the highest ascent and the deepest descent on it (hhea, with the line gap above), so a run whose font
// reaches no further than the paragraph's leaves the line as it is, and one with a deeper descent (Courier New) makes
// it taller below only. The browser centres each run's leading on its own font's box, so a run in a font with a
// smaller box but a taller ascent made the line taller (Cambria: 1.3px, Courier New twice Word's 0.7px). Such a run
// gets the line height that keeps its box inside the paragraph font's, moved out only as far as Word's grows.
const boxes = new Map(); // font, size, line height -> [above, below] the baseline of a text span's box, px
let lineProbe = null;
function spanBox(font, lh, where) {
  const k = `${font}|${lh}`;
  if (boxes.has(k)) return boxes.get(k);
  if (!lineProbe) {
    lineProbe = document.createElement('div');
    lineProbe.setAttribute('aria-hidden', 'true');
    lineProbe.style.cssText = 'position:absolute;left:-99999px;top:0;visibility:hidden;white-space:nowrap;line-height:0';
    lineProbe.innerHTML = '<span>X</span><i style="display:inline-block;width:0;height:0"></i>';
  }
  // (Beside the editor's content, not in its paragraph: the editor reads a change inside its content as an edit made in
  // the page, and drew the paragraph again from it. In a table of merged cells that remade the table, and with Track
  // Changes on, an untouched document saved its tables twice over, as tracked deletions and insertions. The probe sets
  // its own font and line height, so it measures the same there.)
  const content = where.closest('.ProseMirror, [contenteditable]');
  (content?.parentElement || document.body).append(lineProbe);
  const span = lineProbe.firstChild;
  span.style.font = font;
  span.style.lineHeight = `${lh}px`;
  const r = span.getBoundingClientRect();
  const base = lineProbe.lastChild.getBoundingClientRect().top;
  const half = (lh - r.height) / 2;
  const v = [base - r.top + half, r.bottom - base + half, base - r.top, r.bottom - base];
  lineProbe.remove();
  boxes.set(k, v);
  return v;
}
const famOf = (cs) => cs.fontFamily.split(',').map((f) => f.replace(/["']/g, '').trim().toLowerCase()).find((f) => Object.hasOwn(ASCENT, f)) || '';
const fontOfCs = (cs) => `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
/** The line height (px) for a text span in another font than its paragraph's, or 0 to leave it. */
function mixedLine(host, cs) {
  const p = host.closest('p');
  if (!p) return 0;
  const pcs = getComputedStyle(p);
  if (pcs.fontFamily === cs.fontFamily || parseFloat(pcs.getPropertyValue('--odx-runmult')) === 0) return 0;
  const fr = famOf(cs), fp = famOf(pcs);
  if (!fr || !fp || fr === fp) return 0;
  // The run's own line height, as the paragraph gives it (not one set here before).
  const v = getComputedStyle(host.parentElement).getPropertyValue('--odx-lhv').trim();
  const lhOf = (size) => (v.endsWith('px') ? parseFloat(v) : (Number(v) || 0) * size);
  const sr = parseFloat(cs.fontSize), sp = parseFloat(pcs.fontSize);
  const lr = lhOf(sr), lp = lhOf(sp);
  if (!lr || !lp) return 0;
  const [ap, bp] = spanBox(fontOfCs(pcs), lp, p);
  const [ar, br, Ab, Db] = spanBox(fontOfCs(cs), lr, p);
  const grow = (k) => Math.max(0, k(fr) * sr - k(fp) * sp);
  // (By Word's own ascent and descent, Calibri's Windows ones: 0.9521 and 0.2686em, not hhea's 0.9707 and 0.25. Word's
  // prints of 20 generated documents, a Calibri line holding a run in another font: Garamond's (0.2632em deep) and
  // Consolas's (0.251em) leave the line Calibri's own, Courier New's (0.3003em) makes it 0.47px taller at 11pt, Verdana's
  // (1.0054em high) 0.79px. Measured from hhea's 0.25em, Garamond and Consolas runs made 11pt lines 0.2 to 0.4px taller.)
  const up = grow((f) => wordAscent(f)), down = grow((f) => lineFactor(f) - wordAscent(f));
  if (ar <= ap + up + 0.05 && br <= bp + down + 0.05) return 0;
  const lh = Math.min(2 * (ap + up) - (Ab - Db), 2 * (bp + down) + (Ab - Db));
  return lh > 0 && lh < lr - 0.05 ? Math.round(lh * 100) / 100 : 0;
}

let last = '';
// Corrections by paragraph and run (parent node -> run node -> letter spacing, 0 for none), until fonts load.
let memo = new WeakMap();
let memoEpoch = -1;
let loads = 0;
// (The font measures go too: one taken in the fallback while a web font was on its way stayed for good otherwise.)
if (typeof document !== 'undefined' && document.fonts) document.fonts.addEventListener('loadingdone', () => { loads++; hinted.clear(); boxes.clear(); });
const fontEpoch = () => loads;
/** Recomputes the corrections; true when they changed. */
export function applyWidths(view) {
  const decos = [];
  if (memoEpoch !== fontEpoch()) { memo = new WeakMap(); memoEpoch = fontEpoch(); }
  view.state.doc.descendants((node, pos, parent) => {
    if (!node.isText || node.text.length < 2) return !node.isText;
    if (fitOf(node)) return false; // (fitted text has its own width: fittext.js)
    // An unchanged run in an unchanged paragraph keeps its correction (ProseMirror keeps both nodes): no style lookups.
    const known = memo.get(parent)?.get(node);
    if (known !== undefined) {
      if (known) decos.push(Decoration.inline(pos, pos + node.nodeSize, { class: 'odx-wfix', style: known }, { key: `w:${known}` }));
      return false;
    }
    if (!memo.has(parent)) memo.set(parent, new WeakMap());
    const remember = (v) => memo.get(parent).set(node, v);
    const at = view.domAtPos(pos + 1);
    let host = at.node.nodeType === 3 ? at.node.parentElement : at.node;
    if (host && host.classList.contains('odx-wfix')) host = host.parentElement;
    if (!host) return false;
    const cs = getComputedStyle(host);
    const style = [];
    // A character scale (w:w): the run's text as wide as Word makes it, its width times the scale (scaleCss's
    // average-advance estimate made exact for this text: 150% Calibri lines came out a word short or long).
    const scale = parseFloat(cs.getPropertyValue('--odx-w')) || 1;
    // Bold over a cloud SemiBold or Demi family (fakeBold): measured, and so set, at the Regular's widths.
    const fake = fakeBold(cs);
    const m = fake ? { fontStyle: cs.fontStyle, fontWeight: '400', fontFamily: cs.fontFamily, fontSize: cs.fontSize } : cs;
    const hinted = isHinted(m);
    // Small caps as Word sets them: lowercase letters as capitals at 80% of the size (to half a point: 9pt in 11pt
    // text, Word's PDF), the rest as they are; the browser's own small capitals are another width.
    const smallCaps = cs.fontVariantCaps === 'small-caps';
    if (hinted || scale !== 1 || smallCaps || fake) {
      const px = parseFloat(cs.fontSize);
      const actual = width(cs, px, node.text, smallCaps ? 'small-caps' : 'normal');
      let design = hinted ? (width(m, 100, node.text) * px) / 100 : fake ? width(m, px, node.text) : actual;
      if (smallCaps) {
        const sc = (Math.round(px * 0.75 * 0.8 * 2) / 2) / 0.75;
        design = node.text.match(/\p{Ll}+|\P{Ll}+/gu).reduce((t, part) => t + (/\p{Ll}/u.test(part) ? width(m, sc, part.toUpperCase()) : width(m, px, part)), 0);
      }
      const target = design * scale;
      const ls = Math.round(((target - actual) / node.text.length) * 1e4) / 1e4;
      if (ls) style.push(`letter-spacing:calc(var(--odx-csp, 0px) + ${ls}px)`);
    }
    const lh = mixedLine(host, cs);
    if (lh) style.push(`--odx-lhv:${lh}px`);
    const css = style.join(';');
    remember(css);
    if (css) decos.push(Decoration.inline(pos, pos + node.nodeSize, { class: 'odx-wfix', style: css }, { key: `w:${pos}:${node.nodeSize}:${css}` }));
    return false;
  });
  const sig = decos.map((d) => `${d.from}:${d.to}:${d.spec.key}`).join(','); // (positions here, not in the keys: pages.js applyPushes)
  if (sig === last) return false;
  last = sig;
  view.dispatch(view.state.tr.setMeta(key, decos).setMeta('addToHistory', false));
  return true;
}
