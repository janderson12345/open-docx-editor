// Show/Hide ¶ (formatting marks), as in Word: ¶ at the end of each paragraph, a dot in each space, an arrow in each
// tab, ↵ at each line break and a label on page breaks. Nothing is added to the text itself (even splitting it into
// spans moves justified text by fractions of a pixel): tabs, page breaks and empty paragraphs get their marks from CSS
// on elements that are there anyway, and the rest are drawn in a layer over the page, at the positions of the laid-out
// text, for the part of the document on screen. Lines, wrapping and pages are the same with the marks on or off, and
// the marks never print.
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';

const key = new PluginKey('odx-marks');

function compute(doc) {
  const decos = [];
  doc.descendants((node, pos) => {
    if (node.isTextblock) {
      if (!node.content.size) decos.push(Decoration.node(pos, pos + node.nodeSize, { class: 'odx-fm-empty' }));
      return false;
    }
    return true;
  });
  return DecorationSet.create(doc, decos);
}

/** Draws the space, paragraph and line-break marks of the paragraphs on screen into `layer`. */
function draw(view, layer) {
  const host = layer.parentElement;
  if (!host) return;
  const origin = host.getBoundingClientRect();
  const top = -200;
  const bottom = window.innerHeight + 200;
  const frag = document.createDocumentFragment();
  const add = (cls, x, y, h) => {
    const m = document.createElement('span');
    m.className = `odx-fm ${cls}`;
    m.style.cssText = `left:${x - origin.left}px;top:${y - origin.top}px;height:${h}px;line-height:${h}px;font-size:${Math.round(h * 0.7)}px`;
    frag.append(m);
  };
  const range = document.createRange();
  for (const p of view.dom.querySelectorAll('p')) {
    const box = p.getBoundingClientRect();
    if (box.bottom < top || box.top > bottom || !box.height) continue;
    if (p.closest('.odx-joined-away, .odx-sep-away, .odx-all-del')) continue;
    let last = null;
    const walk = document.createTreeWalker(p, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, {
      acceptNode: (n) => (n.nodeType === 1 ? (n.getAttribute('contenteditable') === 'false' ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_SKIP) : NodeFilter.FILTER_ACCEPT),
    });
    for (let n = walk.nextNode(); n; n = walk.nextNode()) {
      const text = n.data;
      for (let i = 0; i < text.length; i++) {
        if (text[i] !== ' ' && text[i] !== ' ') continue;
        range.setStart(n, i);
        range.setEnd(n, i + 1);
        const r = range.getClientRects()[0];
        if (r && r.width) add('odx-fm-sp', r.left + r.width / 2, r.top, r.height);
      }
      if (text.length) last = n;
    }
    for (const br of p.querySelectorAll('br:not(.ProseMirror-trailingBreak)')) {
      const r = br.getBoundingClientRect();
      if (r.height) add('odx-fm-br', r.left, r.top, r.height);
    }
    if (last) {
      range.setStart(last, last.data.length - 1);
      range.setEnd(last, last.data.length);
      const rects = range.getClientRects();
      const r = rects[rects.length - 1];
      if (r) add('odx-fm-para', r.right + 1, r.top, r.height);
    }
  }
  layer.replaceChildren(frag);
}

/** The formatting-marks plugin; `toggleMarks(view)` switches the marks on or off. */
export const marksPlugin = new Plugin({
  key,
  state: {
    init: () => ({ on: false, set: DecorationSet.empty }),
    apply(tr, value) {
      const meta = tr.getMeta(key);
      const on = meta != null ? meta : value.on;
      if (!on) return { on, set: DecorationSet.empty };
      if (meta != null || tr.docChanged || !value.on) return { on, set: compute(tr.doc) };
      return value;
    },
  },
  props: {
    decorations: (state) => key.getState(state).set,
    attributes: (state) => (key.getState(state).on ? { class: 'odx-show-marks' } : {}),
  },
  view(view) {
    const layer = document.createElement('div');
    layer.className = 'odx-fm-layer';
    layer.setAttribute('aria-hidden', 'true');
    let frame = 0;
    const redraw = () => {
      if (frame) return;
      frame = requestAnimationFrame(() => {
        frame = 0;
        if (!key.getState(view.state).on) { layer.remove(); return; }
        if (layer.parentElement !== view.dom.parentElement) view.dom.parentElement?.append(layer);
        draw(view, layer);
      });
    };
    // Pagination moves text after an edit, so draw again a moment later too; scrolling brings other paragraphs on screen.
    let later = 0;
    const settle = () => { redraw(); clearTimeout(later); later = setTimeout(redraw, 400); };
    document.addEventListener('scroll', redraw, true);
    window.addEventListener('resize', redraw);
    return {
      update: settle,
      destroy() { cancelAnimationFrame(frame); clearTimeout(later); document.removeEventListener('scroll', redraw, true); window.removeEventListener('resize', redraw); layer.remove(); },
    };
  },
});

export const marksOn = (state) => !!key.getState(state)?.on;

export function toggleMarks(view) {
  view.dispatch(view.state.tr.setMeta(key, !marksOn(view.state)).setMeta('addToHistory', false));
  // Redrawn tables get their grid widths back; lay the pages out again so fitted columns are reapplied.
  view.odx?.pages?.update?.();
  return true;
}
