// Find and Replace: the bar under the ribbon (Ctrl+F, Ctrl+H), through the main text, text boxes, headers, footers
// and notes (findstories.js).
import { el } from './ribbonkit.js';
import { findNextAll, replaceOneAll, replaceAllStories, findPosition } from './findstories.js';
import { refresh as refreshComments } from './comments.js';

/** The bar (hidden until opened) and `openFind(replace)`, a command that opens it with or without the replace row. */
export function createFindBar(getView, withView) {
  const findBox = el('input', { type: 'search', placeholder: 'Find', 'aria-label': 'Find', class: 'odx-findbox' });
  const replBox = el('input', { type: 'text', placeholder: 'Replace with', 'aria-label': 'Replace with', class: 'odx-findbox' });
  const matchCase = el('input', { type: 'checkbox', title: 'Match case' });
  const wholeWord = el('input', { type: 'checkbox', title: 'Whole words only' });
  // Patterns: the query is a regular expression (\d+, colou?r, ^Chapter), and a replacement can use what its
  // parentheses caught ($1, $2).
  const regex = el('input', { type: 'checkbox', title: 'Use patterns (regular expressions): \\d for a digit, . for any character, (…) to reuse as $1 in the replacement' });
  const findInfo = el('span', { class: 'odx-findinfo', role: 'status' }); // read out by screen readers (4.1.3)
  // Closed: back to the document, and a comment Find had opened in the margin closed too (its match is over).
  const closeFind = () => {
    findBar.hidden = true;
    const v = getView();
    const main = v?.odx?.main || v;
    const session = main?.odx?.session;
    if (session && session.findComment != null) {
      if (session.editingComment === session.findComment) session.editingComment = null;
      session.findComment = null;
      session.commentSelect = null;
      refreshComments(main);
    }
    v?.focus();
  };
  const fOpts = () => ({ matchCase: matchCase.checked, wholeWord: wholeWord.checked, regex: regex.checked });
  const fdo = (fn) => (e) => {
    e.preventDefault();
    const v = getView();
    // (The keys stay in the find bar, as in Word's find pane, when a match takes the search into a header, note, text box
    // or comment: their editors take the focus as they open, and the next Enter would have typed into them.)
    const from = findBar.contains(document.activeElement) ? document.activeElement : null;
    if (from) setTimeout(() => setTimeout(() => { if (!findBar.hidden && from.isConnected) from.focus(); }, 0), 0);
    let n;
    try { n = fn(v); } catch (err) {
      if (!(err instanceof SyntaxError)) throw err;
      findInfo.textContent = `Not a valid pattern: ${err.message.replace(/^Invalid regular expression: /, '')}`;
      return;
    }
    // Which match it went to, and where (said by screen readers each time: the status's text changes).
    const at = n ? findPosition() : null;
    findInfo.textContent = n == null ? '' : at ? `${at.index} of ${n}${at.part ? `, ${at.part}` : ''}` : n ? `${n} result${n === 1 ? '' : 's'}` : 'No results';
  };
  const replRow = el('span', { class: 'odx-replrow' }, replBox,
    el('button', { type: 'button', onmousedown: fdo((v) => replaceOneAll(v, findBox.value, replBox.value, fOpts())) }, 'Replace'),
    el('button', { type: 'button', onmousedown: fdo((v) => { const n = replaceAllStories(v, findBox.value, replBox.value, fOpts()); findInfo.textContent = ''; return n; }) }, 'Replace all'));
  const findBar = el('div', { class: 'odx-findbar', hidden: true },
    findBox,
    el('button', { type: 'button', title: 'Previous result', onmousedown: fdo((v) => findNextAll(v, findBox.value, fOpts(), -1)) }, '↑'),
    el('button', { type: 'button', title: 'Next result', onmousedown: fdo((v) => findNextAll(v, findBox.value, fOpts(), 1)) }, '↓'),
    el('label', {}, matchCase, 'Match case'), el('label', {}, wholeWord, 'Whole words'), el('label', {}, regex, 'Patterns'),
    replRow, findInfo,
    el('button', { type: 'button', title: 'Close', class: 'odx-close', onmousedown: (e) => { e.preventDefault(); closeFind(); } }, '✕'));
  findBox.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') fdo((v) => findNextAll(v, findBox.value, fOpts(), e.shiftKey ? -1 : 1))(e);
    if (e.key === 'Escape') closeFind();
  });
  replBox.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeFind(); });
  const openFind = (replace) => withView((v) => {
    findBar.hidden = false;
    replRow.hidden = !replace;
    const { from, to } = v.state.selection;
    if (to > from && to - from < 100) findBox.value = v.state.doc.textBetween(from, to);
    setTimeout(() => findBox.select());
    return true;
  });
  document.addEventListener('keydown', (e) => {
    const v = getView();
    // (In the document, or in the find bar itself: Ctrl+H there turns Find into Replace, as in Word, not the browser's history.)
    if (!v || !(e.ctrlKey || e.metaKey) || !(v.hasFocus() || findBar.contains(document.activeElement))) return;
    if (e.key === 'f' || e.key === 'h') { e.preventDefault(); openFind(e.key === 'h')(v.state, v.dispatch, v); }
  });
  return { findBar, openFind };
}
