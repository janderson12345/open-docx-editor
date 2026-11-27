// A list paragraph whose first word is longer than its line: Word breaks the word on the label's line, filling it to
// the margin (a bullet and the first part of a long slashed word, then the rest of it, in a narrow cell; Word's prints of
// why-tests/fix2711, bullets in a cell and in the body). The browser does too after the label alone, but the room
// for a label reaching above the text (pararules.js labelLine, --odx-lbl-own: an inline-block before the text) is a
// place it may break, so it put the word on a line of its own under the label. Such a word is given break-all, to be
// broken where it reaches the margin; measured after the layout like fittext.js.
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { natural } from './fittext.js';

const key = new PluginKey('odx-lblword');
export const labelWordPlugin = new Plugin({
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

/** The first word of a paragraph node: [[text node, its offset in the paragraph, length taken]...], or null. */
function firstWord(para) {
  const parts = [];
  let done = false;
  para.forEach((n, off) => {
    if (done) return;
    if (!n.isText) { done = true; return; }
    const ws = n.text.search(/\s/);
    if (ws === 0) { done = true; return; }
    parts.push([n, off, ws < 0 ? n.text.length : ws]);
    if (ws > 0) done = true;
  });
  return parts.length ? parts : null;
}

let last = '';
/** Breaks list paragraphs' overlong first words as Word does; true when that changed (the page is measured again). */
export function applyLabelWords(view) {
  const decos = [];
  view.state.doc.descendants((para, ppos) => {
    if (!para.isTextblock) return true;
    if (para.type.name !== 'paragraph' || !para.childCount) return false;
    const el = view.nodeDOM(ppos);
    if (!el || el.nodeType !== 1 || !el.classList.contains('odx-li') || !el.style.getPropertyValue('--odx-lbl-own')) return false;
    if (el.matches('.odx-lbl-right, .odx-contlines, .odx-hjoin-nolabel')) return false;
    const word = firstWord(para);
    if (!word) return false;
    const cs = getComputedStyle(el);
    const label = parseFloat(getComputedStyle(el, '::before').width) || 0;
    const room = el.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight) - parseFloat(cs.textIndent) - label;
    let width = 0;
    for (const [n, off, len] of word) {
      const at = view.domAtPos(ppos + 1 + off + 1);
      const tel = at.node.nodeType === 3 ? at.node.parentElement : at.node;
      if (!tel) return false;
      width += natural(tel, n.text.slice(0, len));
    }
    if (!(room > 0) || width <= room + 0.5) return false;
    for (const [, off, len] of word) decos.push(Decoration.inline(ppos + 1 + off, ppos + 1 + off + len, { style: 'word-break:break-all' }, { key: 'lw' }));
    return false;
  });
  const sig = decos.map((d) => `${d.from}:${d.to}`).join(',');
  if (sig === last) return false;
  last = sig;
  view.dispatch(view.state.tr.setMeta(key, decos).setMeta('addToHistory', false));
  return true;
}
