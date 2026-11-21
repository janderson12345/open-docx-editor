// Word 2013+ (compatibility mode 15) squeezes the spaces of a justified line to fit one more word, when
// that narrows each space by at most a quarter (and 0.07em) and by no more than half of what each would be widened
// by with the word left for the next line (see squeezeFits). Measured from Word's own PDFs.
// Browsers only stretch, so for each line Word would squeeze we narrow that line's spaces with an inline
// decoration; the browser then fits the extra word, and justification spreads what is left.
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { schema } from './schema.js';
import { ZERO_WIDTH } from '../ooxml/reader.js';
import { splitOf } from './wrapboth.js';
import { mergeTabs } from '../ooxml/props.js';

const MAX_SQUEEZE = 0.25; // share of a space's width
const MAX_SQUEEZE_EM = 0.07; // and of the font size (Verdana's wide spaces squeeze less than a quarter)
const SHRINK_WEIGHT = 2; // a space narrowed counts as one widened by twice as much
/**
 * Whether Word squeezes a word of width `w` on to a line it sticks out of by `over` px, its `n` spaces (`spaceW` px
 * wide in all, of fonts `em` px in all) narrowed to take it. Also for each side of a line split beside a shape
 * (wrapboth.js, `side`; wrap-word2 wj-both: the words after the shape squeezed as a line's are): there by the spaces'
 * limits alone, however far the word sticks out (wrap-word3, Word's print of six documents with a rectangle beside
 * each paragraph: of 128 sides it squeezed, the most by 24.9% of their spaces' width, of those it left a word off the
 * least by 25.4%; "a" squeezed on to a side it stuck out of by 44% of its width).
 *
 * On a whole line Word weighs the two ways of ending it: the word taken on, each of the `n` spaces narrowed by
 * over / n, or left for the next line, each of the `gaps` spaces before it widened by what the line is then short of
 * (`space` px of space before the word, plus the word, less `over`). It takes the word on when the narrowing is no
 * more than half that widening. Word's prints of 66 generated justified documents, full lines and lines of three to
 * eight words, 4,516 words that stuck out of a line: every one in plain text is as the rule says, the nearest either
 * side at 2.005 and 1.993 times; in all, seven differ, three on the quarter's limit and four in text whose runs carry
 * letter-spacing, within 3% of the rule. (Before: the word's share past the margin against a fitted limit, 31 of them
 * wrong, 13 in the narrow lines.)
 */
export function squeezeFits(over, n, spaceW, em, w, side = false, space = spaceW / n, gaps = n - 1) {
  if (!n) return false;
  const per = Math.min(MAX_SQUEEZE * spaceW, MAX_SQUEEZE_EM * em) / n; // the most each space may lose
  if (over / n > per) return false;
  if (side || gaps < 1) return true; // (a line of one word can't be widened: the squeeze, when the spaces allow it)
  return (space + w - over) / gaps >= SHRINK_WEIGHT * over / n;
}

const key = new PluginKey('odx-justify');
export const justifyPlugin = new Plugin({
  key,
  state: {
    init: () => DecorationSet.empty,
    apply(tr, set) {
      const next = tr.getMeta(key);
      if (next) return DecorationSet.create(tr.doc, next);
      const drop = tr.getMeta(dropKey);
      const out = set.map(tr.mapping, tr.doc);
      if (!drop) return out;
      const gone = [];
      for (const [from, to] of drop) gone.push(...out.find(from, to).filter((d) => d.to > from && d.from < to));
      return gone.length ? out.remove(gone) : out;
    },
  },
  props: { decorations: (state) => key.getState(state) },
});

/**
 * Takes away the squeezes over the ranges `at` ([from, to]: wrapboth.js, lines split beside a shape, the column's
 * width no longer theirs); the next pass works them out again. Returns the transaction.
 */
const dropKey = new PluginKey('odx-justify-drop');
export function dropSqueeze(tr, at) {
  if (!at.length) return tr;
  last = '';
  return tr.setMeta(dropKey, at);
}

let ctx = null;
const fontOf = (cs) => `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;

/**
 * One paragraph's lines as Word would break them. Returns [{ from, to, squeeze }] for the lines Word squeezes
 * (PM positions of the line's start and the end of the word it pulls in, squeeze in px per space), or null
 * when the paragraph holds anything but plain text, or isn't laid out.
 */
function squeezedLines(view, node, pos, el, split, session) {
  const segs = []; // [text, pm start, font, space width]
  let plain = true;
  // A tab: the text after it is set from where the tab ends, and its line's spaces are counted from there (Word
  // justifies a line after its last tab only). Word's prints of 18 generated documents, justified paragraphs with a
  // label and a tab to their hanging indent, a tab that opens them, or one to a stop after their first words: Word
  // squeezes their lines as a plain paragraph's, the line with the tab too; left alone, 246 of their 1,046 lines broke
  // as Word's. (Left stops only: after a centre, right or decimal stop the text's place depends on its own width.)
  let tabsOk = null;
  const leftStops = () => {
    if (tabsOk == null) {
      try { tabsOk = mergeTabs(session?.styleProp(node.attrs.styleId, 'tabs'), node.attrs.view?.tabs).every((t) => !t.val || t.val === 'left' || t.val === 'start' || t.val === 'num' || t.val === 'bar'); } catch { tabsOk = false; }
    }
    return tabsOk;
  };
  node.descendants((child, off) => {
    if (!plain) return false;
    // Markers that take no room (bookmarks, comment anchors and references) don't stop it.
    if (child.type === schema.nodes.opaque_inline) {
      const dom = view.nodeDOM(pos + 1 + off);
      if (ZERO_WIDTH.has(child.attrs.kind) || child.attrs.kind === 'proofErr' || (dom && dom.nodeType === 1 && !dom.getBoundingClientRect().width)) return false;
    }
    if (!child.isLeaf) return true; // content controls and the like: their text counts
    // Nor a shape set out of the flow, wrapped with text on both sides (wrapboth.js splits the lines beside it).
    if (child.isAtom && !child.isText) {
      const dom = view.nodeDOM(pos + 1 + off);
      if (dom && dom.nodeType === 1 && getComputedStyle(dom).position === 'absolute' && dom.style.getPropertyValue('--odx-both')) return false;
    }
    // A line break (w:br, w:cr) ends a line where it is; the lines either side are fitted on their own.
    if (child.type === schema.nodes.hard_break) { segs.push({ brk: true, start: pos + 1 + off }); return false; }
    if (child.type === schema.nodes.tab && !split && leftStops()) { segs.push({ tab: true, start: pos + 1 + off }); return false; }
    if (!child.isText || /[\t\xad]/.test(child.text)) { plain = false; return false; }
    const at = view.domAtPos(pos + 1 + off + 1);
    const host = at.node.nodeType === 3 ? at.node.parentElement : at.node;
    const cs = getComputedStyle(host);
    // (A run's own letter-spacing is measured with its text, the space's too: Word squeezes such lines as any other.
    // Its prints of 27 justified documents whose runs each carry a few twips of spacing, as a document converted from a
    // PDF has: none of them broke as Word's while these paragraphs were left alone.)
    const ls = cs.letterSpacing === 'normal' ? 0 : parseFloat(cs.letterSpacing) || 0;
    ctx.font = fontOf(cs);
    ctx.letterSpacing = `${ls}px`;
    // Measured as the page sets it: kerning is off unless the run turns it on (w:kern), as in Word.
    const kern = cs.fontKerning === 'none' ? 'none' : 'normal';
    ctx.fontKerning = kern;
    segs.push({ text: child.text, start: pos + 1 + off, font: ctx.font, kern, ls: ctx.letterSpacing, space: ctx.measureText(' ').width, em: parseFloat(cs.fontSize) });
  });
  if (!plain || !segs.length) return null;
  // Line widths from the layout: the first line starts after any indent or list label.
  const cs = getComputedStyle(el);
  const box = el.getBoundingClientRect();
  const left = box.left + parseFloat(cs.borderLeftWidth) + parseFloat(cs.paddingLeft);
  const right = box.right - parseFloat(cs.borderRightWidth) - parseFloat(cs.paddingRight);
  const r = document.createRange();
  const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.textContent.length ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP) });
  const firstText = walk.nextNode();
  if (!firstText) return null;
  r.setStart(firstText, 0);
  r.setEnd(firstText, 1);
  const fr = r.getClientRects()[0];
  if (!fr || !fr.height || right - left < 20) return null;
  // Words: runs of non-space characters, measured piece by piece in each piece's font.
  const words = []; // { w, from, to, space }  (space: width of the space before it)
  let cur = null;
  let pendingSpace = 0;
  let pendingEm = 0;
  for (const s of segs) {
    if (s.brk || s.tab) {
      if (cur) { words.push(cur); cur = null; }
      words.push({ brk: !!s.brk, tab: !!s.tab, from: s.start, to: s.start + 1, space: pendingSpace });
      pendingSpace = 0;
      pendingEm = 0;
      continue;
    }
    ctx.font = s.font;
    ctx.fontKerning = s.kern;
    ctx.letterSpacing = s.ls;
    let i = 0;
    while (i < s.text.length) {
      if (s.text[i] === ' ') {
        if (cur) { words.push(cur); cur = null; }
        pendingSpace += s.space;
        pendingEm = Math.max(pendingEm, s.em);
        i++;
        continue;
      }
      let j = i;
      while (j < s.text.length && s.text[j] !== ' ' && s.text[j - 1] !== '-') j++;
      if (j === i) j++;
      const w = ctx.measureText(s.text.slice(i, j)).width;
      if (!cur) { cur = { w: 0, from: s.start + i, to: 0, space: pendingSpace, em: pendingEm }; pendingSpace = 0; pendingEm = 0; }
      cur.w += w;
      cur.to = s.start + j;
      cur.ls = parseFloat(s.ls) || 0; // (its last character's)
      i = j;
      // A line can also break after a hyphen.
      if (s.text[j - 1] === '-' && j < s.text.length && s.text[j] !== ' ') { words.push(cur); cur = null; }
    }
  }
  if (cur) words.push(cur);
  const out = [];
  // (Split beside a shape, wrapboth.js: the lines before the first split one, as from the paragraph's start, and those
  // after the last, as from a line's.)
  if (split) {
    run(words.filter((w) => w.to <= split.first), right - fr.left, segs[0].start);
    if (split.next != null) run(words.filter((w) => w.from >= split.next), right - left, split.next);
  } else run(words, right - fr.left, segs[0].start);
  return out;

  function run(words, avail, lineFrom) {
    let width = 0;
    let spaces = 0;
    let spaceW = 0;
    let spaceEm = 0;
    let lineX = right - avail; // where the line's text starts, or the text after its last tab
    for (let k = 0; k < words.length; k++) {
      const wd = words[k];
      if (wd.brk) { width = 0; lineFrom = wd.to; lineX = left; avail = right - left; continue; }
      if (wd.tab) {
        // (Where the text after it starts: at the stop the tab goes to, which tabs.js leaves on it, from the
        // paragraph's left edge. Not where it is drawn, nor its width: a justified line's spaces before the tab are
        // stretched there, and by another amount once the line is squeezed, the squeeze then taken off and put back
        // pass after pass: two samples never finished laying out. One taken to the next line with the word after it
        // isn't followed, nor one whose stop the words before it have passed.)
        const dom = view.nodeDOM(wd.from);
        const end = dom && dom.nodeType === 1 && !dom.classList.contains('odx-tabnl') ? parseFloat(dom.dataset.odxTabend) : NaN;
        const x = box.left + end;
        if (!Number.isFinite(end) || !(x < right) || x < lineX + width - 0.5) return;
        width = 0; lineFrom = wd.to; lineX = x; avail = right - x; spaces = 0; spaceW = 0; spaceEm = 0;
        continue;
      }
      if (k === 0 || width === 0) { width = wd.w; spaces = 0; spaceW = 0; spaceEm = 0; continue; }
      const add = wd.space + wd.w;
      const gap = wd.space > 0 ? 1 : 0;
      if (width + add <= avail + 0.01) { width += add; spaces += gap; spaceW += wd.space; spaceEm += gap * wd.em; continue; }
      const over = width + add - avail;
      const n = spaces + gap;
      // Measurement hook (the rule against Word's own line breaks); off unless a test sets it.
      // (`at`, `chars`: the letters of the document before the line and on it before the word, which tell a print's
      // lines apart; `width`, `avail`: the line without the word and its room.)
      if (typeof window !== 'undefined' && window.__sqdbg) {
        const letters = (a, b) => view.state.doc.textBetween(a, b, '', '').replace(/\s/g, '').length;
        window.__sqdbg.push({ word: view.state.doc.textBetween(wd.from, wd.to), before: view.state.doc.textBetween(Math.max(lineFrom, wd.from - 40), wd.from), over, n, spaceW: spaceW + wd.space, w: wd.w, em: wd.em, at: letters(0, lineFrom), chars: letters(lineFrom, wd.from), width, avail, space: wd.space, lsWord: wd.ls, lsPrev: words[k - 1]?.ls ?? 0 });
      }
      if (squeezeFits(over, n, spaceW + wd.space, spaceEm + gap * wd.em, wd.w, false, wd.space, spaces)) {
        // A little extra: canvas and layout widths differ slightly, and justification spreads it back out
        // (never so much that the word after it could fit too). Not on a paragraph's last line, or one ending in a
        // line break: nothing spreads those, so only half a pixel in all (Word squeezes those by just what fits).
        const after = words[k + 1]?.brk || words[k + 1]?.tab ? null : words[k + 1];
        const slack = after ? Math.min(0.3, (after.space + after.w) / 2 / n) : Math.min(0.3, 0.5 / n);
        out.push({ from: lineFrom, to: wd.to, squeeze: over / n + slack });
        width = 0; // the next word starts a new line
        lineFrom = wd.to;
        lineX = left;
        avail = right - left;
        continue;
      }
      // Break before this word.
      lineFrom = wd.from;
      lineX = left;
      avail = right - left;
      width = wd.w;
      spaces = 0;
      spaceW = 0;
      spaceEm = 0;
    }
  }
}

/** Recomputes the squeezed lines of every justified paragraph; true when the decorations changed. */
let last = '';
export function applySqueeze(view, session) {
  if (!(session.compatMode >= 15)) return false;
  ctx = ctx || document.createElement('canvas').getContext('2d');
  const decos = [];
  view.state.doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.paragraph) return true;
    const el = view.nodeDOM(pos);
    if (!el || el.nodeType !== 1 || getComputedStyle(el).textAlign !== 'justify') return false;
    // (Lines split beside a shape: wrapboth.js spreads those, and these only the lines above and below them.)
    const split = el.querySelector('.odx-wb') ? splitOf(view.state, pos, pos + node.nodeSize) : null;
    if (split === false) return false;
    const lines = squeezedLines(view, node, pos, el, split, session);
    for (const l of lines || []) decos.push(Decoration.inline(l.from, l.to, { style: `word-spacing:${(-l.squeeze).toFixed(2)}px` }, { key: `sq:${l.squeeze.toFixed(2)}` }));
    return false;
  });
  const sig = decos.map((d) => `${d.from}:${d.to}:${d.spec.key}`).join(','); // (positions here, not in the keys: pages.js applyPushes)
  if (sig === last) return false;
  last = sig;
  view.dispatch(view.state.tr.setMeta(key, decos).setMeta('addToHistory', false));
  return true;
}
