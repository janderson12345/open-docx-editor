// Line pitch as Word lays it. Word stacks lines at the font's line height, unrounded (Times New Roman 10pt:
// 11.5pt = 15.333px a line). Edge rounds the font's ascent and descent to whole px and truncates each half of
// the leading to 1/64px, so the same line comes out shorter (15.3125px), by an amount that depends on the font,
// size and line height: up to ~1px over a page, enough to fit a line Word moves to the next page.
// For each paragraph style in use, a hidden probe finds the line height (just above the intended one) whose laid-out
// pitch comes closest, and the paragraph gets that value as a node decoration: a number for "auto" spacing (runs
// in other sizes still scale from it), a length for exact and at-least spacing (as they were).
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { schema } from './schema.js';
import { lineHeightOf } from './util.js';

const key = new PluginKey('odx-lhfix');
export const lineFixPlugin = new Plugin({
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

const cache = new Map(); // paragraph and run styles, line height -> chosen line height in px (or null: leave it)
const OURS = /;?\s*--odx-lhv:\s*[\d.]+(px)?;?\s*$/; // the decoration's part of a paragraph's style attribute

/**
 * Edge's distance between two lines' tops (px) for a paragraph laid out with line height `lh` (px), measured on a
 * copy of it (same classes and style, inside a copy of the editor's content box) holding text in a copy of its
 * first run's elements: a line is as tall as the runs on it, each rounded on its own.
 */
function pitch(box, el, chain, lh, length) {
  const q = el.cloneNode(false);
  q.removeAttribute('data-odx-lh');
  q.setAttribute('style', (el.getAttribute('style') || '').replace(OURS, ''));
  q.style.setProperty('--odx-lhv', length ? `${lh}px` : String(lh / parseFloat(getComputedStyle(el).fontSize)));
  let inner = q;
  for (const c of chain) { const k = c.cloneNode(false); inner.append(k); inner = k; }
  box.append(q);
  inner.textContent = 'x';
  const one = q.getBoundingClientRect().height;
  inner.textContent = 'x\nx\nx\nx\nx';
  const five = q.getBoundingClientRect().height;
  q.remove();
  return (five - one) / 4;
}

/** The line height (px) whose pitch is nearest `want`, or null when Edge already lays it within 0.002px. */
function bestHeight(box, el, chain, want, length) {
  const k = [el.className, (el.getAttribute('style') || '').replace(OURS, ''), ...chain.map((c) => `${c.tagName}.${c.className}:${c.getAttribute('style') || ''}`), want, length].join('|');
  if (cache.has(k)) return cache.get(k);
  let best = null;
  let err = Math.abs(pitch(box, el, chain, want, length) - want);
  if (err > 0.002) {
    for (let i = 1; i <= 4; i++) {
      const c = want + i / 128;
      const e = Math.abs(pitch(box, el, chain, c, length) - want);
      if (e < err - 1e-6) { err = e; best = c; }
    }
  }
  cache.set(k, best);
  return best;
}

let last = '';
// Results by paragraph node ({ value, want }, null for none), until fonts load: probing every paragraph again on each
// layout pass was a long document's single largest cost after the table measuring.
let memo = new WeakMap();
let memoFonts = -1;
let fontLoads = 0;
if (typeof document !== 'undefined' && document.fonts) document.fonts.addEventListener('loadingdone', () => { fontLoads++; });
/** Recomputes the paragraphs' line heights; true when the decorations changed (measure again once drawn). */
export function applyLineFix(view) {
  const decos = [];
  // Probes go in a copy of the content box (its classes give the document's styles), off to the side.
  const box = document.createElement('div');
  box.className = view.dom.className;
  box.setAttribute('aria-hidden', 'true');
  box.style.cssText = `position:absolute;left:-99999px;top:0;width:${view.dom.clientWidth}px;visibility:hidden;contain:layout`;
  view.dom.parentElement.append(box);
  if (memoFonts !== fontLoads) { memo = new WeakMap(); memoFonts = fontLoads; }
  view.state.doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.paragraph) return true;
    // An unchanged paragraph keeps its result (ProseMirror keeps the node): no probing again.
    if (memo.has(node)) {
      const r = memo.get(node);
      if (r) decos.push(Decoration.node(pos, pos + node.nodeSize, { style: `--odx-lhv:${r.value}`, 'data-odx-lh': String(r.want) }, { key: `lh:${r.value}` }));
      return false;
    }
    const el = view.nodeDOM(pos);
    if (!el || el.nodeType !== 1) return false;
    memo.set(node, null); // (none, unless found below)
    const cs = getComputedStyle(el);
    const fs = parseFloat(cs.fontSize);
    // The value Edge used before any fix of ours (a decoration from an earlier pass shows in the data attribute).
    const want = el.dataset.odxLh ? Number(el.dataset.odxLh) : parseFloat(lineHeightOf(cs));
    if (!want || !fs) return false;
    // The text's own elements (Word documents often set the font and size on runs, not the paragraph). Not an equation's:
    // its parts (a script with one of its two children copied) aren't a line of text, and Firefox reports them as errors.
    const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.data.trim() && !n.parentElement.closest('math') ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP) });
    const text = walk.nextNode();
    const chain = [];
    for (let n = text?.parentElement; n && n !== el; n = n.parentElement) chain.unshift(n);
    const length = cs.getPropertyValue('--odx-lhlen').trim() === '1';
    const best = bestHeight(box, el, chain, want, length);
    if (best == null) return false;
    const value = length ? `${best.toFixed(4)}px` : `${(best / fs).toFixed(6)}`;
    memo.set(node, { value, want });
    decos.push(Decoration.node(pos, pos + node.nodeSize, { style: `--odx-lhv:${value}`, 'data-odx-lh': String(want) }, { key: `lh:${value}` }));
    return false;
  });
  box.remove();
  const sig = decos.map((d) => `${d.from}:${d.to}:${d.spec.key}`).join(','); // (positions here, not in the keys: pages.js applyPushes)
  if (sig === last) return false;
  last = sig;
  view.dispatch(view.state.tr.setMeta(key, decos).setMeta('addToHistory', false));
  return true;
}
