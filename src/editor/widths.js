// Text widths for fonts Edge sets with hinted advances. Edge lays some fonts out at small sizes with their hinted
// (device) advances, not the scaled design ones Word uses: Cambria 10pt comes out 3.6% narrower than its design widths,
// 12pt 1.9% wider, 11pt right, and Word's lines take fewer or more words (corpus 2's justified
// Cambria paragraph a line shorter). Each run in such a font and size gets the difference back as letter spacing: its
// text's width at 100px, scaled, against its width at its size, shared out per character (on top of the run's own
// character spacing, --odx-csp).
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { fitOf } from './fittext.js';
import { ASCENT } from './baseline.js';
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
function width(cs, px, text) {
  if (!probe || !probe.isConnected) {
    probe = document.createElement('span');
    probe.setAttribute('aria-hidden', 'true');
    probe.style.cssText = 'position:absolute;left:-99999px;top:0;visibility:hidden;white-space:pre;letter-spacing:0;word-spacing:0;font-kerning:none;font-variant-ligatures:none;line-height:normal';
    document.body.append(probe);
  }
  probe.style.font = `${cs.fontStyle} ${cs.fontWeight} ${px}px ${cs.fontFamily}`;
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
  where.append(lineProbe);
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
const famOf = (cs) => cs.fontFamily.split(',').map((f) => f.replace(/["']/g, '').trim().toLowerCase()).find((f) => ASCENT[f]) || '';
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
  const up = grow((f) => ASCENT[f]), down = grow((f) => lineFactor(f) - ASCENT[f]);
  if (ar <= ap + up + 0.05 && br <= bp + down + 0.05) return 0;
  const lh = Math.min(2 * (ap + up) - (Ab - Db), 2 * (bp + down) + (Ab - Db));
  return lh > 0 && lh < lr - 0.05 ? Math.round(lh * 100) / 100 : 0;
}

let last = '';
// Corrections by paragraph and run (parent node -> run node -> letter spacing, 0 for none), until fonts load.
let memo = new WeakMap();
let memoEpoch = -1;
let loads = 0;
if (typeof document !== 'undefined' && document.fonts) document.fonts.addEventListener('loadingdone', () => { loads++; });
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
    const hinted = isHinted(cs);
    if (hinted || scale !== 1) {
      const px = parseFloat(cs.fontSize);
      const actual = width(cs, px, node.text);
      const target = (hinted ? (width(cs, 100, node.text) * px) / 100 : actual) * scale;
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
