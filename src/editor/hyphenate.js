// Automatic hyphenation (w:autoHyphenation), as Word hyphenates a document's words at line ends. A browser that
// hyphenates English itself (Firefox; Chromium but on Windows) does it: CSS hyphens:auto on the paragraphs. One that
// doesn't (Edge and Chrome on Windows have no dictionaries) gets each word's break points, from the TeX English
// patterns, as soft hyphens: widgets on screen only, never in the document's text. Either way a line ends at the last
// break point that fits, with a hyphen, as Word's does. (Not widgets everywhere: Firefox's caret sticks in them.)
// Paragraphs and styles with w:suppressAutoHyphens are left alone, and words in capitals with w:doNotHyphenateCaps.
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
const shy = () => { const s = document.createElement('span'); s.className = 'odx-shy'; s.textContent = '\u00ad'; return s; };

/** The soft-hyphen widgets for one paragraph at `pos` (memoised on the node: unchanged paragraphs cost nothing). */
const memo = new WeakMap();
function paragraphDecos(node, pos, session) {
  let rel = memo.get(node);
  if (!rel) {
    rel = [];
    const caps = session.doNotHyphenateCaps;
    node.forEach((child, off) => {
      if (!child.isText) return;
      for (const m of child.text.matchAll(WORD)) {
        if (m[0].length < MIN_WORD || (caps && m[0] === m[0].toUpperCase())) continue;
        const parts = hyphenate(m[0]);
        let at = off + m.index;
        for (let i = 0; i < parts.length - 1; i++) { at += parts[i].length; rel.push({ at, marks: child.marks }); }
      }
    });
    memo.set(node, rel);
  }
  return rel.map(({ at, marks }) => Decoration.widget(pos + 1 + at, shy, { side: -1, marks, key: 'shy', ignoreSelection: true }));
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
    if (own) decos.push(Decoration.node(pos, pos + node.nodeSize, { class: 'odx-hyph' }));
    else decos.push(...paragraphDecos(node, pos, session));
    return false;
  });
  return DecorationSet.create(doc, decos);
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
      return {};
    },
  });
}
