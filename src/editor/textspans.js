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
import { w14Color } from '../ooxml/texteffects.js';

// Inline content that takes no room on a line: markers (bookmarks, empty content controls) and floating pictures.
const marker = (c) => (c.type.name === 'opaque_inline' && !c.attrs.text && c.attrs.width == null) || (c.type.name === 'image' && /float:|position:\s*absolute/.test(c.attrs.place || ''));

const strut = () => {
  const s = document.createElement('span');
  s.className = 'odx-t odx-strut';
  s.contentEditable = 'false';
  s.textContent = '​';
  return s;
};

// A run with Word 2010's reflection (w14:reflection; texteffects.js gives it the room below the line and the custom
// properties) is drawn mirrored below its baseline: each word in an odx-rf span carrying its text, which editor.css
// draws again as the span's ::after, flipped and faded. Generated content isn't selected, copied, found or read out,
// and is in no saved document; a word at a time, so a run that wraps is mirrored on each of its lines.
//
// A run with an inner shadow (a w14:shadow scaled to nothing, sx and sy 0) the same way: each word in an odx-is span,
// which editor.css draws again over itself as its ::before, through the shadow's SVG filter (texteffects.js), leaving
// only the shadow inside the letters.
const effect = (c, name, test) => c.marks.some((m) => m.type.name === 'rextra' && m.attrs.items.some(([n, x]) => n === name && x.startsWith('<w14:') && test(x)));
const reflects = (c) => effect(c, 'reflection', () => true);
const pressed = (c) => effect(c, 'shadow', (x) => /\bw14:sx="0"/.test(x) && /\bw14:sy="0"/.test(x) && !!w14Color(x));
const WORD = /\S+/g;
/** A word's decoration attributes for a reflection and an inner shadow. */
const wordAttrs = (w, rf, is) => ({ class: [rf && 'odx-rf', is && 'odx-is'].filter(Boolean).join(' '), ...(rf ? { 'data-odx-rf': w } : {}), ...(is ? { 'data-odx-is': w } : {}) });

/** The decorations of one textblock at `pos`: each text node in an odx-t span, a strut before a break after a break. */
function blockDecos(node, pos, out) {
  let afterBreak = false; // only markers since the last line break
  let first = true; // no text yet
  node.forEach((c, off) => {
    const at = pos + 1 + off;
    // (Plain spaces alone also odx-blank: they don't size their line in Word, nor does a shadow on them: editor.css.)
    // (The block's first text is odx-first too: where a list label in a font of its own makes Word's line taller, the
    // room is added above it, editor.css.)
    if (c.isText) out.push(Decoration.inline(at, at + c.nodeSize, { class: (/^ +$/.test(c.text) ? 'odx-t odx-blank' : 'odx-t') + (first ? ' odx-first' : '') }));
    if (c.isText) first = false;
    const rf = c.isText && reflects(c), is = c.isText && pressed(c);
    if (rf || is) for (const m of c.text.matchAll(WORD)) out.push(Decoration.inline(at + m.index, at + m.index + m[0].length, wordAttrs(m[0], rf, is)));
    if (c.type.name === 'hard_break') {
      if (afterBreak) out.push(Decoration.widget(at, strut, { side: -1, key: 'odx-strut', ignoreSelection: true }));
      afterBreak = true;
    } else if (!marker(c)) afterBreak = false;
  });
  // The spaces that end a paragraph, after its last text: marked, so that under a centred or right-aligned paragraph they
  // take no room (editor.css). Word sets the line as if they weren't there; the browser counts them at a paragraph's end
  // (a centred "Unit Cost " header 1.6px left of Word's, a right-aligned amount 3.3px, one space's width).
  let last = null, lastAt = 0, some = false;
  node.forEach((c, off) => { if (c.isText) { if (last && last.text.trim()) some = true; last = c; lastAt = pos + 1 + off; } else if (!marker(c)) { if (last && last.text.trim()) some = true; last = null; } });
  if (last) {
    // (Counted back from the end, not by a pattern: " +$" tries again from every space of a long run that doesn't end
    // the text, time that grows with the square of its length in a document made to be slow.)
    let from = last.text.length;
    while (from > 0 && last.text.charCodeAt(from - 1) === 32) from--;
    if (from < last.text.length && (from > 0 || some)) out.push(Decoration.inline(lastAt + from, lastAt + last.nodeSize, { class: 'odx-trail' }));
  }
  return out;
}

/** Every text span's decoration in the document (spansSet arranges them in a set). */
export function spanDecos(doc) {
  const out = [];
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    blockDecos(node, pos, out);
    return false;
  });
  return out;
}

/**
 * DecorationSet.create(doc, spanDecos(doc)), made a top-level block at a time: each block's own set, the document's set
 * holding them as its children, as create arranges them. (create looks for each top-level block's decorations among
 * all that are left: a long document's spans took quadratic time, 6,400 paragraphs 0.35s and 4 times as many blocks 14
 * times as long; test/stress.test.js checks both give the same set.)
 */
export function spansSet(doc) {
  const children = [];
  doc.forEach((block, offset) => {
    // (Positions in the block, from its content's start, as a child set's are.)
    const out = [];
    if (block.isTextblock) blockDecos(block, -1, out);
    else block.descendants((node, pos) => { if (!node.isTextblock) return true; blockDecos(node, pos, out); return false; });
    const set = out.length ? DecorationSet.create(block, out) : DecorationSet.empty;
    if (set !== DecorationSet.empty) children.push(offset, offset + block.nodeSize, set);
  });
  return children.length ? new DecorationSet([], children) : DecorationSet.empty;
}

const MANY_STEPS = 100;
/** A step replacing more than half of the document (the one step of an Accept All on many changes). */
function wide(tr) {
  const half = tr.before.content.size / 2;
  return tr.steps.some((s) => s.from != null && s.to - s.from > half && (s.slice?.size ?? 0) > tr.doc.content.size / 2);
}

export const textSpansPlugin = new Plugin({
  state: {
    init: (_, { doc }) => spansSet(doc),
    apply(tr, set) {
      if (!tr.docChanged) return set;
      // Many steps (Accept All, a paste of a long document), or one taking in most of the document: made again whole.
      // (Mapping each step's range through the steps after it is quadratic in their number: 6000 steps took a second.)
      if (tr.steps.length > MANY_STEPS || wide(tr)) return spansSet(tr.doc);
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
    s.className = /^ +$/.test(t.data) ? 'odx-t odx-blank' : 'odx-t';
    t.replaceWith(s);
    s.append(t);
    // (A reflected or inner-shadowed run's words, as in the editor.)
    const rf = !!s.closest('[style*="--odx-rf-a"]'), is = !!s.closest('[style*="--odx-is:"]');
    if (rf || is) {
      s.replaceChildren(...t.data.split(/(\S+)/).filter(Boolean).map((w) => {
        if (!/\S/.test(w)) return document.createTextNode(w);
        const r = document.createElement('span');
        const a = wordAttrs(w, rf, is);
        for (const [k, v] of Object.entries(a)) r.setAttribute(k, v);
        r.textContent = w;
        return r;
      }));
    }
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
