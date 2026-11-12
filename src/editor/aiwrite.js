// AI writing in the editor: the host's model (the `ai` option: an async function the host writes, so the editor sends
// nothing anywhere itself) asked to rewrite the selection (shorter, longer, more formal...), to summarise it or the
// document, or anything else about it. A rewrite goes into the document as a tracked change under the AI's name, for
// the person to accept or reject as any other; an answer is shown, then inserted below or put in place of the
// selection (tracked too) or copied, as they choose.
import { Fragment, Slice } from 'prosemirror-model';
import { schema } from './schema.js';
import { trackTransaction } from './track.js';

/** The rewrites offered: [action, label]. */
export const AI_REWRITES = [
  ['rewrite', 'Rewrite'],
  ['shorter', 'Make it shorter'],
  ['longer', 'Make it longer'],
  ['formal', 'More formal'],
  ['casual', 'More casual'],
  ['proofread', 'Fix spelling and grammar'],
];

const OBJ = '\ufffc'; // (a picture or other object in the text: kept out of what's sent, one character in the count)

/**
 * What a request is about: the selection, or with none the paragraph the cursor is in. { from, to, text, doc, before,
 * after } (the paragraphs either side, for context).
 */
export function aiTarget(view) {
  const { state } = view;
  let { from, to } = state.selection;
  if (from === to) {
    const $p = state.selection.$from;
    if ($p.parent.isTextblock) { from = $p.start(); to = $p.end(); }
  }
  const doc = state.doc;
  const text = doc.textBetween(from, to, '\n', OBJ);
  const $f = doc.resolve(from);
  const $t = doc.resolve(to);
  const para = (pos, dir) => {
    const $x = doc.resolve(pos);
    const d = $x.depth;
    if (!d) return '';
    const idx = $x.index(d - 1) + dir;
    const parent = $x.node(d - 1);
    return idx >= 0 && idx < parent.childCount ? parent.child(idx).textContent : '';
  };
  return { from, to, text, doc, before: $f.depth ? para(from, -1) : '', after: $t.depth ? para(to, 1) : '' };
}

/** The whole document's text, paragraph by paragraph (for a summary of it, or a question about it). */
export function documentText(view) {
  return view.state.doc.textBetween(0, view.state.doc.content.size, '\n', OBJ).replace(/\n{3,}/g, '\n\n');
}

/** Paragraphs of `text` (one a line), each in `marks`, in `styleId`. */
function paragraphsOf(text, marks, styleId = null) {
  return text.split(/\r?\n/).map((line) => schema.nodes.paragraph.create({ styleId }, line ? schema.text(line, marks) : null));
}

/** The first text's marks in [from, to] (the selection's formatting, which the new text takes). */
function marksAt(doc, from, to) {
  let marks = null;
  doc.nodesBetween(from, Math.max(to, from + 1), (n) => { if (!marks && n.isText) marks = n.marks; return !marks; });
  return marks || doc.resolve(from).marks();
}

/** A transaction applied as a tracked change under `author`. */
function applyTracked(view, tr, author) {
  view.dispatch(trackTransaction(view.state, tr, author).setMeta('odx-untracked', true).scrollIntoView());
  view.odx?.pages?.balloons?.();
}

/**
 * `text` in place of [from, to] as a tracked change under `author`. Within one paragraph only what differs is changed
 * (the common start and end are kept, so the change shows just the edit); across paragraphs, the whole.
 */
export function replaceTracked(view, from, to, text, author) {
  const doc = view.state.doc;
  const old = doc.textBetween(from, to, '\n', OBJ);
  const tr = view.state.tr;
  if (!old.includes('\n') && !text.includes('\n')) {
    let s = 0;
    while (s < old.length && s < text.length && old[s] === text[s]) s++;
    let e = 0;
    while (e < old.length - s && e < text.length - s && old[old.length - 1 - e] === text[text.length - 1 - e]) e++;
    if (s === old.length && s === text.length) return false;
    tr.insertText(text.slice(s, text.length - e), from + s, to - e);
  } else {
    const paras = paragraphsOf(text, marksAt(doc, from, to), doc.resolve(from).parent.attrs.styleId ?? null);
    tr.replaceRange(from, to, new Slice(Fragment.fromArray(paras), 1, 1));
  }
  applyTracked(view, tr, author);
  return true;
}

/** `text` as new paragraphs after the paragraph holding `pos`, as a tracked insertion under `author`. */
export function insertBelowTracked(view, pos, text, author) {
  const $p = view.state.doc.resolve(pos);
  const at = $p.parent.isTextblock ? $p.after() : pos;
  const tr = view.state.tr.insert(at, paragraphsOf(text.trim(), []));
  applyTracked(view, tr, author);
  return true;
}

/**
 * Asks the host's model: `ai({ action, instruction, prompt, text, before, after, document, signal })`, where action is
 * one of AI_REWRITES' or 'summarize' or 'ask'. Returns its text (a string, or { text }).
 */
export async function askAi(ai, request) {
  const out = await ai(request);
  const text = typeof out === 'string' ? out : out?.text;
  if (typeof text !== 'string') throw new Error('The AI gave no text');
  return text.replace(/\r\n/g, '\n').replace(/^\n+|\n+$/g, '');
}
