// Automatic hyphenation (w:autoHyphenation), as Word hyphenates a document's words at line ends. A browser that
// hyphenates English itself (Firefox; Chromium but on Windows) does it: CSS hyphens:auto on the paragraphs. One that
// doesn't (Edge and Chrome on Windows have no dictionaries) gets each word's break points, from the TeX English
// patterns, as soft hyphens: widgets on screen only, never in the document's text. Either way a line ends at the last
// break point that fits, with a hyphen, as Word's does. (Not widgets everywhere: Firefox's caret sticks in them.)
// Paragraphs and styles with w:suppressAutoHyphens are left alone, and words in capitals with w:doNotHyphenateCaps.
// The hyphenation zone (Word 2010's layout and older) and the consecutive-hyphen limit are followed too, once the
// lines are laid out (holdWords): a word Word would leave whole loses its soft hyphens, or, where the browser
// hyphenates itself, is marked hyphens:manual (odx-nohyph).
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { schema } from './schema.js';

const key = new PluginKey('odx-hyphenate');
const MIN_WORD = 5; // letters in a word before it is hyphenated
let hyphenate = null; // (word) => its parts, once the patterns are loaded

/** Loads the English patterns (a separate chunk, only for a document that hyphenates). */
export async function loadHyphenation() {
  if (hyphenate) return;
  const m = await import('hyphen/en-us/index.js');
  const sync = (m.default || m).hyphenateSync;
  // (At least two letters before a break and three after it, as English is set: never "complicat-ed".)
  hyphenate = (w) => {
    const parts = sync(w, { hyphenChar: '\u00ad', minWordLength: MIN_WORD }).split('\u00ad');
    while (parts.length > 1 && parts[parts.length - 1].length < 3) parts.splice(-2, 2, parts[parts.length - 2] + parts[parts.length - 1]);
    while (parts.length > 1 && parts[0].length < 2) parts.splice(0, 2, parts[0] + parts[1]);
    return parts;
  };
}

/** Whether this browser hyphenates English itself (hyphens:auto breaks an overlong word). */
let native = null;
export function nativeHyphens() {
  if (native != null) return native;
  try {
    const d = document.createElement('div');
    d.lang = 'en-US';
    d.style.cssText = 'position:absolute;visibility:hidden;width:48px;font:16px serif;hyphens:manual;white-space:normal';
    d.textContent = 'hyphenation characterization';
    document.body.append(d);
    const manual = d.getBoundingClientRect().height;
    d.style.hyphens = 'auto';
    native = d.getBoundingClientRect().height > manual;
    d.remove();
  } catch { native = false; }
  return native;
}

const WORD = /[A-Za-z\u00c0-\u024f]+(?:['\u2019][A-Za-z]+)?/g;
/** A soft-hyphen widget in the word starting at `w` (its offset in the paragraph). */
const shyIn = (w) => () => { const s = document.createElement('span'); s.className = 'odx-shy'; s.dataset.w = w; s.textContent = '\u00ad'; return s; };

// Words left whole (their break points dropped) for the hyphenation zone or the consecutive-hyphen limit, as Word
// leaves them: per paragraph node and layout (zoom and width: lines break a little differently at each), the offsets of
// the words' starts, to their lengths. (A paragraph edited is a new node: worked out again.)
const held = new WeakMap();
const layoutOf = new WeakMap(); // session -> the layout its words are held for now
const heldIn = (node, session) => held.get(node)?.get(layoutOf.get(session));
function hold(node, session, w, len) {
  const at = layoutOf.get(session);
  if (!held.has(node)) held.set(node, new Map());
  if (!held.get(node).has(at)) held.get(node).set(at, new Map());
  held.get(node).get(at).set(w, len);
}

/** The words of a paragraph that may be hyphenated: { w (offset in the paragraph), len, text node's marks }. */
function wordsOf(node, session) {
  const out = [];
  const caps = session.doNotHyphenateCaps;
  node.forEach((child, off) => {
    if (!child.isText) return;
    for (const m of child.text.matchAll(WORD)) {
      if (m[0].length < MIN_WORD || (caps && m[0] === m[0].toUpperCase())) continue;
      out.push({ w: off + m.index, text: m[0], marks: child.marks });
    }
  });
  return out;
}

/** The soft-hyphen widgets for one paragraph at `pos` (memoised on the node: unchanged paragraphs cost nothing). */
const memo = new WeakMap();
function paragraphDecos(node, pos, session) {
  let rel = memo.get(node);
  if (!rel) {
    rel = [];
    for (const { w, text, marks } of wordsOf(node, session)) {
      const parts = hyphenate(text);
      let at = w;
      for (let i = 0; i < parts.length - 1; i++) { at += parts[i].length; rel.push({ at, w, marks }); }
    }
    memo.set(node, rel);
  }
  const whole = heldIn(node, session);
  return rel.filter((r) => !whole?.has(r.w)).map(({ at, w, marks }) => Decoration.widget(pos + 1 + at, shyIn(w), { side: -1, marks, key: `shy${w}`, ignoreSelection: true }));
}

const MAX_ROUNDS = 60;

/**
 * Finds, in each paragraph, the first line ending in a hyphen that Word wouldn't hyphenate, and leaves that word whole:
 * one where the line, with the word wrapped whole instead, would leave no more room at its end than the hyphenation zone
 * (session.hyphenationZone, twips); or one more line ending in a hyphen than the limit allows in a row
 * (session.consecutiveHyphenLimit). True when a word was left whole (the widgets are to be made again).
 */
function holdWords(view, session) {
  const zone = session.hyphenationZone / 15; // px
  const limit = session.consecutiveHyphenLimit;
  const native = nativeHyphens();
  let changed = false;
  view.state.doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.paragraph) return true;
    const el = view.nodeDOM(pos);
    if (!el || !el.querySelectorAll) return false;
    const shown = native ? nativeBreaks(view, node, pos, session) : widgetBreaks(el);
    if (!shown.length) return false;
    const box = el.getBoundingClientRect();
    const scale = el.offsetWidth ? box.width / el.offsetWidth : 1;
    // (Measured unjustified: where the lines break doesn't change, and the room left at a line's end is plain to see.)
    const align = el.style.textAlign;
    el.style.textAlign = 'left';
    try {
      const cs = getComputedStyle(el);
      const right = box.right - (parseFloat(cs.paddingRight) + parseFloat(cs.borderRightWidth)) * scale;
      let run = 0;
      let nextLine = null; // where the text after the last hyphen goes on
      for (const { w, len, r, next } of shown) {
        run = nextLine && sameLine(nextLine, r) ? run + 1 : 1;
        nextLine = next;
        let whole = limit > 0 && run > limit;
        if (!whole && zone > 0) {
          // The end of the text before the word (on the same line: else the word starts the line, and must break).
          const before = node.textBetween(0, w, '', '\ufffc').replace(/\s+$/, '').length;
          if (before > 0) {
            const c = view.coordsAtPos(pos + 1 + before, -1);
            if (sameLine(c, r)) whole = (right - c.left) / scale <= zone;
          }
        }
        if (whole) {
          hold(node, session, w, len);
          changed = true;
          break;
        }
      }
    } finally { el.style.textAlign = align; }
    return false;
  });
  return changed;
}

const sameLine = (a, b) => Math.abs((a.top + a.bottom) / 2 - (b.top + b.bottom) / 2) < (b.bottom - b.top) / 2;

/**
 * The lines of a paragraph that end inside a word, in order: { w, len (the word's start and length), r (the rect of the
 * letter before the break), next (of the one after) }. With the soft hyphens, where the letters either side of a widget
 * are on different lines. (Not the hyphen's width: on a page's last line, pages.js's, it measures nothing.)
 */
function widgetBreaks(el) {
  const out = [];
  for (const s of el.querySelectorAll('.odx-shy')) {
    const a = charRect(s, el, -1);
    const b = charRect(s, el, 1);
    if (a && b && !sameLine(a, b)) out.push({ w: Number(s.dataset.w), len: 0, r: a, next: b });
  }
  return out;
}

/** The same where the browser hyphenates itself: words whose first and last letters are on different lines. */
function nativeBreaks(view, node, pos, session) {
  const out = [];
  for (const { w, text } of wordsOf(node, session)) {
    const at = (k) => letterRect(view, pos + 1 + w + k);
    const first = at(0);
    const last = at(text.length - 1);
    if (!first || !last || sameLine(first, last)) continue;
    for (let k = 1, prev = first; k < text.length; k++) {
      const r = at(k);
      if (!r) break;
      if (!sameLine(prev, r)) { out.push({ w, len: text.length, r: prev, next: r }); break; }
      prev = r;
    }
  }
  return out;
}

/** The rect of the letter at document position `p` (the one after it). */
function letterRect(view, p) {
  let { node: n, offset } = view.domAtPos(p);
  if (n.nodeType !== 3) {
    n = n.childNodes[offset];
    while (n && n.nodeType === 1) n = n.firstChild;
    offset = 0;
  }
  if (!n || n.nodeType !== 3 || offset >= n.data.length) return null;
  const range = document.createRange();
  range.setStart(n, offset);
  range.setEnd(n, offset + 1);
  return range.getClientRects()[0] || null;
}

/** The rect of the character just after (`dir` 1) or before (-1) `s` (a hyphen widget) inside `el`. */
function charRect(s, el, dir) {
  const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  walk.currentNode = s.firstChild || s;
  for (let t = dir > 0 ? walk.nextNode() : walk.previousNode(); t; t = dir > 0 ? walk.nextNode() : walk.previousNode()) {
    if (!t.data || t.parentElement.closest('.odx-shy')) continue;
    const range = document.createRange();
    const at = dir > 0 ? 0 : t.data.length - 1;
    range.setStart(t, at);
    range.setEnd(t, at + 1);
    const r = range.getClientRects()[0];
    if (r) return r;
  }
  return null;
}

function noHyph(node, session) {
  const own = node.attrs.view?.noHyph;
  return own != null ? own : !!session.styleProp(node.attrs.styleId, 'noHyph');
}

function build(doc, session) {
  const own = nativeHyphens();
  if (!own && !hyphenate) return DecorationSet.empty;
  const decos = [];
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.paragraph) return true;
    if (noHyph(node, session)) return false;
    if (own) {
      decos.push(Decoration.node(pos, pos + node.nodeSize, { class: 'odx-hyph' }));
      // (The words the zone or the limit leave whole: not hyphenated, editor.css.)
      for (const [w, len] of heldIn(node, session) || []) decos.push(Decoration.inline(pos + 1 + w, pos + 1 + w + len, { class: 'odx-nohyph' }));
    } else decos.push(...paragraphDecos(node, pos, session));
    return false;
  });
  return DecorationSet.create(doc, decos);
}

const sessions = new WeakMap(); // view -> its session
const busy = new WeakSet(); // views settling now (their own dispatches start nothing more)
const rules = (session) => session.autoHyphenation && (hyphenate || nativeHyphens()) && (session.hyphenationZone > 0 || session.consecutiveHyphenLimit > 0);
const redraw = (view) => view.dispatch(view.state.tr.setMeta(key, true).setMeta('addToHistory', false));

/**
 * The words the hyphenation zone and the consecutive-hyphen limit leave whole, for the lines as laid out now: on screen
 * after an edit, and as printed (pages.js, at full size after the zoom on screen is taken off, and again after).
 */
export function settleHyphens(view) {
  settle(view);
}

function settle(view) {
  const session = sessions.get(view);
  if (!session || view.isDestroyed || !rules(session)) return;
  busy.add(view);
  try {
    // (Another zoom or width: the words held for it, if it was laid out before, else none yet.)
    const at = `${view.dom.currentCSSZoom ?? ''}|${view.dom.offsetWidth}`;
    if (layoutOf.get(session) !== at) { layoutOf.set(session, at); redraw(view); }
    for (let n = 0; n < MAX_ROUNDS && holdWords(view, session); n++) redraw(view);
  } finally { busy.delete(view); }
}

/** The plugin; `ready` (a transaction meta) builds the widgets once the patterns have loaded. */
export function hyphenatePlugin(session) {
  return new Plugin({
    key,
    state: {
      init: (_, state) => (session.autoHyphenation ? build(state.doc, session) : DecorationSet.empty),
      apply(tr, set) {
        if (!session.autoHyphenation) return set;
        return tr.docChanged || tr.getMeta(key) ? build(tr.doc, session) : set;
      },
    },
    props: {
      decorations: (state) => key.getState(state),
      // (For the print: a line broken inside a word gets its hyphen written in, printlines.js.)
      attributes: () => (session.autoHyphenation ? { 'data-hyph': '' } : {}),
    },
    view(view) {
      if (session.autoHyphenation && !hyphenate && !nativeHyphens()) loadHyphenation().then(() => { if (!view.isDestroyed) view.dispatch(view.state.tr.setMeta(key, true).setMeta('addToHistory', false)); });
      // The zone and the limit, once the lines are laid out: a word to leave whole found, its break points go and the
      // lines are laid out again, until none is left (each round settles one more line in each paragraph). The rounds
      // go one after another at once (a dispatch redraws the paragraphs, and measuring lays them out), not each after
      // the pages have settled again: a document printed soon after opening had them half done.
      let timer = null;
      const later = () => { clearTimeout(timer); timer = setTimeout(() => settle(view), 60); };
      sessions.set(view, session);
      return {
        update() { if (!busy.has(view) && rules(session)) later(); },
        destroy() { clearTimeout(timer); },
      };
    },
  });
}
