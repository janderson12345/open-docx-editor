// Text widths for fonts Edge sets with hinted advances. Edge lays some fonts out at small sizes with their hinted
// (device) advances, not the scaled design ones Word uses: Cambria 10pt comes out 3.6% narrower than its design widths,
// 12pt 1.9% wider, 11pt right, and Word's lines take fewer or more words (corpus 2's justified
// Cambria paragraph a line shorter). Each run in such a font and size gets the difference back as letter spacing: its
// text's width at 100px, scaled, against its width at its size, shared out per character.
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';

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
    // An unchanged run in an unchanged paragraph keeps its correction (ProseMirror keeps both nodes): no style lookups.
    const known = memo.get(parent)?.get(node);
    if (known !== undefined) {
      if (known) decos.push(Decoration.inline(pos, pos + node.nodeSize, { class: 'odx-wfix', style: `letter-spacing:${known}px` }, { key: `w:${pos}:${node.nodeSize}:${known}` }));
      return false;
    }
    if (!memo.has(parent)) memo.set(parent, new WeakMap());
    const remember = (v) => memo.get(parent).set(node, v);
    const at = view.domAtPos(pos + 1);
    let host = at.node.nodeType === 3 ? at.node.parentElement : at.node;
    if (host && host.classList.contains('odx-wfix')) host = host.parentElement;
    if (!host) return false;
    const cs = getComputedStyle(host);
    if (!isHinted(cs)) { remember(0); return false; }
    const px = parseFloat(cs.fontSize);
    const actual = width(cs, px, node.text);
    const target = (width(cs, 100, node.text) * px) / 100;
    const ls = Math.round(((target - actual) / node.text.length) * 1e4) / 1e4;
    remember(ls);
    if (ls) decos.push(Decoration.inline(pos, pos + node.nodeSize, { class: 'odx-wfix', style: `letter-spacing:${ls}px` }, { key: `w:${pos}:${node.nodeSize}:${ls}` }));
    return false;
  });
  const sig = decos.map((d) => d.spec.key).join(',');
  if (sig === last) return false;
  last = sig;
  view.dispatch(view.state.tr.setMeta(key, decos).setMeta('addToHistory', false));
  return true;
}
