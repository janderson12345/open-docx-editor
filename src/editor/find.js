// Find and Replace over the document text (matches can span runs with different formatting).
import { TextSelection } from 'prosemirror-state';

/** Text of each textblock with the document position of every character. */
function blocks(doc) {
  const out = [];
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    let text = '';
    const at = [];
    node.forEach((child, off) => {
      const start = pos + 1 + off;
      if (child.isText) {
        for (let i = 0; i < child.text.length; i++) { text += child.text[i]; at.push(start + i); }
      } else { text += child.type.name === 'tab' ? '\t' : '￼'; at.push(start); }
    });
    at.push(pos + 1 + node.content.size);
    out.push({ text, at });
    return false;
  });
  return out;
}

/** All matches of `query` as [{from, to}]. */
export function findAll(doc, query, { matchCase = false, wholeWord = false } = {}) {
  if (!query) return [];
  const literal = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(wholeWord ? `\\b${literal}\\b` : literal, matchCase ? 'g' : 'gi');
  const out = [];
  for (const b of blocks(doc)) for (const m of b.text.matchAll(re)) out.push({ from: b.at[m.index], to: b.at[m.index + m[0].length - 1] + 1 });
  return out;
}

/** Select the next (dir 1) or previous (dir -1) match after the selection. Returns the match count. */
export function findNext(view, query, opts, dir = 1) {
  const all = findAll(view.state.doc, query, opts);
  if (!all.length) return 0;
  const { from, to } = view.state.selection;
  const m = dir > 0 ? all.find((x) => x.from >= to) || all[0] : [...all].reverse().find((x) => x.to <= from) || all[all.length - 1];
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, m.from, m.to)).scrollIntoView());
  return all.length;
}

/** Replace the selected match (if it is one) and move to the next. */
export function replaceOne(view, query, replacement, opts) {
  const { from, to } = view.state.selection;
  const hit = findAll(view.state.doc, query, opts).find((x) => x.from === from && x.to === to);
  if (hit) view.dispatch(view.state.tr.insertText(replacement, from, to));
  return findNext(view, query, opts, 1);
}

/** Replace every match; returns how many were replaced. */
export function replaceAll(view, query, replacement, opts) {
  const all = findAll(view.state.doc, query, opts);
  if (!all.length) return 0;
  const tr = view.state.tr;
  for (const m of [...all].reverse()) tr.insertText(replacement, m.from, m.to);
  view.dispatch(tr);
  return all.length;
}
