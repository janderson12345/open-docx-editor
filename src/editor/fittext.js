// Fit text (w:fitText, Word's Asian Layout > Fit Text): a run, or the runs sharing its w:id, spread or squeezed to the
// width it gives, as Word lays it out: the difference shared out after each character but the last (Word's PDF of
// test/fixtures/fittext.docx: "Name" fitted to 1in starts the text after it exactly 1in on). Drawn as letter spacing,
// measured after the layout like widths.js.
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { schema } from './schema.js';

const key = new PluginKey('odx-fittext');
export const fitTextPlugin = new Plugin({
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

/** A text node's fit: { width (px), id }, or null. */
export function fitOf(node) {
  const m = node.marks.find((x) => x.type === schema.marks.rextra && x.attrs.items.some(([k]) => k === 'fitText'));
  const xml = m && m.attrs.items.find(([k]) => k === 'fitText')[1];
  const val = xml && Number(/w:val="(\d+)"/.exec(xml)?.[1]);
  return val ? { width: val / 15, id: /w:id="([^"]*)"/.exec(xml)?.[1] ?? '' } : null;
}

let probe = null;
/** The text's width in its element's font, spaced as plain text. */
function natural(el, text) {
  if (!probe || !probe.isConnected) {
    probe = document.createElement('span');
    probe.setAttribute('aria-hidden', 'true');
    probe.style.cssText = 'position:absolute;left:-99999px;top:0;visibility:hidden;white-space:pre;letter-spacing:0;word-spacing:0;line-height:normal';
    document.body.append(probe);
  }
  const cs = getComputedStyle(el);
  probe.style.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
  // (Kerning and ligatures as the text has them: the editor turns kerning off, and kerned text measured narrower.)
  probe.style.fontKerning = cs.fontKerning;
  probe.style.fontVariantLigatures = cs.fontVariantLigatures;
  probe.textContent = text;
  return probe.getBoundingClientRect().width;
}

let last = '';
/** Spaces each fitted run out to its width; true when the spacing changed (the page is measured again). */
export function applyFitText(view) {
  const decos = [];
  view.state.doc.descendants((para, ppos) => {
    if (!para.isTextblock) return true;
    // Runs in a row with the same fit (the same w:id) are one group.
    const groups = [];
    let cur = null;
    para.forEach((n, off) => {
      const f = n.isText && fitOf(n);
      if (f && cur && cur.id === f.id) cur.nodes.push([n, ppos + 1 + off]);
      else if (f) groups.push((cur = { ...f, nodes: [[n, ppos + 1 + off]] }));
      else if (n.isText || n.isInline) cur = null;
    });
    for (const g of groups) {
      let width = 0, chars = 0;
      for (const [n, pos] of g.nodes) {
        const at = view.domAtPos(pos + 1);
        const el = at.node.nodeType === 3 ? at.node.parentElement : at.node;
        if (!el) return false;
        width += natural(el, n.text);
        chars += n.text.length;
      }
      if (chars < 2) continue;
      const ls = Math.round(((g.width - width) / (chars - 1)) * 1e4) / 1e4;
      for (const [n, pos] of g.nodes) decos.push(Decoration.inline(pos, pos + n.nodeSize, { class: 'odx-fit', style: `letter-spacing:${ls}px` }, { key: `f:${ls}` }));
      // (None after the last character.)
      const [ln, lpos] = g.nodes[g.nodes.length - 1];
      decos.push(Decoration.inline(lpos + ln.nodeSize - 1, lpos + ln.nodeSize, { class: 'odx-fit', style: 'letter-spacing:0' }, { key: 'f:end' }));
    }
    return false;
  });
  const sig = decos.map((d) => `${d.from}:${d.to}:${d.spec.key}`).join(',');
  if (sig === last) return false;
  last = sig;
  view.dispatch(view.state.tr.setMeta(key, decos).setMeta('addToHistory', false));
  return true;
}
