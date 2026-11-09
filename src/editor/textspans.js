// Lines as tall as the text on them, in any document mode. Word sizes a line from the runs on it; the browser in
// standards mode also counts each paragraph's own font on every line (its strut), so a line of smaller runs, a line
// holding only a picture, or a span wrapping another span's text made lines taller than Word's. (In quirks mode the
// browser leaves out boxes with no text of their own, which is what the layout was fitted on.)
//
// So: every text node is in a span of its own (odx-t), the paragraph and the inline elements around the text have
// line-height 0, and the text spans take the line height everything would have had (--odx-lhv, a registered property
// that inherits like line-height; editor.css). A line holding nothing but a line break is as tall as the paragraph's
// font, as in Word and in quirks mode: an empty paragraph by a rule in editor.css, a break right after another break by
// a zero-width text span put before it. (ProseMirror puts a zero-size image on a line that starts or ends with a break,
// and with it neither mode gives that line the strut: editor.css's odx-markers rules size it.)
import { Plugin } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';

// Inline content that takes no room on a line: markers (bookmarks, empty content controls) and floating pictures.
const marker = (c) => (c.type.name === 'opaque_inline' && !c.attrs.text && c.attrs.width == null) || (c.type.name === 'image' && /float:|position:\s*absolute/.test(c.attrs.place || ''));

const strut = () => {
  const s = document.createElement('span');
  s.className = 'odx-t odx-strut';
  s.contentEditable = 'false';
  s.textContent = '​';
  return s;
};

/** The decorations of one textblock at `pos`: each text node in an odx-t span, a strut before a break after a break. */
function blockDecos(node, pos, out) {
  let afterBreak = false; // only markers since the last line break
  node.forEach((c, off) => {
    const at = pos + 1 + off;
    if (c.isText) out.push(Decoration.inline(at, at + c.nodeSize, { class: 'odx-t' }));
    if (c.type.name === 'hard_break') {
      if (afterBreak) out.push(Decoration.widget(at, strut, { side: -1, key: 'odx-strut', ignoreSelection: true }));
      afterBreak = true;
    } else if (!marker(c)) afterBreak = false;
  });
  return out;
}

function build(doc) {
  const out = [];
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    blockDecos(node, pos, out);
    return false;
  });
  return out;
}

export const textSpansPlugin = new Plugin({
  state: {
    init: (_, { doc }) => DecorationSet.create(doc, build(doc)),
    apply(tr, set) {
      if (!tr.docChanged) return set;
      set = set.map(tr.mapping, tr.doc);
      // The textblocks a step touched, in the new document: their decorations made again. (A mark step changes no
      // positions but splits text nodes: its range is in its map too.)
      const maps = tr.mapping.maps;
      const touched = [];
      tr.steps.forEach((step, i) => {
        let { from, to } = step;
        if (from == null) step.getMap().forEach((_s, _e, a, b) => { from = Math.min(from ?? a, a); to = Math.max(to ?? b, b); });
        if (from == null) return;
        for (let j = i; j < maps.length; j++) { from = maps[j].map(from, -1); to = maps[j].map(to, 1); }
        touched.push([Math.max(0, from - 1), Math.min(tr.doc.content.size, to + 1)]);
      });
      for (const [a, b] of touched) {
        tr.doc.nodesBetween(a, b, (node, pos) => {
          if (!node.isTextblock) return true;
          set = set.remove(set.find(pos, pos + node.nodeSize));
          set = set.add(tr.doc, blockDecos(node, pos, []));
          return false;
        });
      }
      return set;
    },
  },
  props: { decorations(state) { return this.getState(state); } },
});

/** The same for content drawn outside an editor (headers, footers, notes as drawn on the sheets): text in odx-t spans,
 *  and a strut on a line holding only a break (no zero-size image there: a break at a paragraph's start too). */
export function wrapText(root) {
  const walk = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, { acceptNode: (t) => (t.data && !t.parentElement.closest('math, svg, .odx-t') ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT) });
  const texts = [];
  for (let t = walk.nextNode(); t; t = walk.nextNode()) texts.push(t);
  for (const t of texts) {
    const s = document.createElement('span');
    s.className = 'odx-t';
    t.replaceWith(s);
    s.append(t);
  }
  const isMarker = (c) => (c.classList.contains('odx-opaque-inline') && !c.textContent) || (c.tagName === 'IMG' && /float:|position:\s*absolute/.test(c.getAttribute('style') || ''));
  for (const p of root.querySelectorAll('p')) {
    let lineStart = true;
    for (const c of [...p.children]) {
      if (c.tagName === 'BR') { if (lineStart) c.before(strut()); lineStart = true; }
      else if (!isMarker(c)) lineStart = false;
    }
  }
}
