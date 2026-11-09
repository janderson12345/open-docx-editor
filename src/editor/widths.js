// Text widths for fonts Edge sets with hinted advances. Edge lays some fonts out at small sizes with their hinted
// (device) advances, not the scaled design ones Word uses: Cambria 10pt comes out 3.6% narrower than its design widths,
// 12pt 1.9% wider, 11pt right (measured with canvas text metrics), and Word's lines take fewer or more words (corpus 2's justified
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
let ctx = null;
const fontAt = (cs, px) => `${cs.fontStyle} ${cs.fontWeight} ${px}px ${cs.fontFamily}`;
function isHinted(cs) {
  const k = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontFamily} ${cs.fontSize}`;
  if (!hinted.has(k)) {
    const px = parseFloat(cs.fontSize);
    ctx.font = fontAt(cs, px);
    const at = ctx.measureText(REF).width;
    ctx.font = fontAt(cs, 100);
    const scaled = (ctx.measureText(REF).width * px) / 100;
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
  ctx = ctx || document.createElement('canvas').getContext('2d');
  ctx.letterSpacing = '0px';
  ctx.fontKerning = 'none';
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
    ctx.font = fontAt(cs, px);
    const actual = ctx.measureText(node.text).width;
    ctx.font = fontAt(cs, 100);
    const target = (ctx.measureText(node.text).width * px) / 100;
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
