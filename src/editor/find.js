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

/**
 * The search as a regular expression: the query as typed, or (`regex`) a pattern of its own. Throws on a pattern that
 * isn't one (findbar.js says so).
 */
export function searchPattern(query, { matchCase = false, wholeWord = false, regex = false } = {}) {
  const body = regex ? query : query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp(wholeWord ? `\\b(?:${body})\\b` : body, matchCase ? 'gu' : 'giu');
}

/**
 * All matches of `query` as [{from, to, groups, named}], within each paragraph. With `regex`, the query is a regular
 * expression (groups: what its parentheses caught, for a replacement's $1, $2...); matches of nothing are skipped.
 */
export function findAll(doc, query, opts = {}) {
  if (!query) return [];
  const re = searchPattern(query, opts);
  const out = [];
  for (const b of blocks(doc)) {
    for (const m of b.text.matchAll(re)) {
      if (!m[0]) continue;
      out.push({ from: b.at[m.index], to: b.at[m.index + m[0].length - 1] + 1, groups: opts.regex ? [...m] : null, named: m.groups || null });
    }
  }
  return out;
}

/** The replacement for a match: with a pattern, $& (the match), $1-$99 and $<name> filled in, $$ for a dollar sign. */
export function replacementFor(m, replacement) {
  if (!m.groups) return replacement;
  return replacement.replace(/\$(\$|&|\d{1,2}|<([^>]+)>)/g, (all, k, name) => {
    if (k === '$') return '$';
    if (k === '&') return m.groups[0];
    if (name != null) return m.named?.[name] ?? '';
    // (Two digits where there aren't that many groups: the first digit's group and the second digit as it is, as
    // String.prototype.replace reads "$10" with one group.)
    const n = Number(k);
    if (n > 0 && n < m.groups.length) return m.groups[n] ?? '';
    if (k.length === 2 && Number(k[0]) > 0 && Number(k[0]) < m.groups.length) return (m.groups[Number(k[0])] ?? '') + k[1];
    return all;
  });
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
  if (hit) view.dispatch(view.state.tr.insertText(replacementFor(hit, replacement), from, to));
  return findNext(view, query, opts, 1);
}

/** Replace every match; returns how many were replaced. */
export function replaceAll(view, query, replacement, opts) {
  const all = findAll(view.state.doc, query, opts);
  if (!all.length) return 0;
  const tr = view.state.tr;
  for (const m of [...all].reverse()) tr.insertText(replacementFor(m, replacement), m.from, m.to);
  view.dispatch(tr);
  return all.length;
}
