// Find and Replace: the bar under the ribbon (Ctrl+F, Ctrl+H).
import { el } from './ribbonkit.js';
import { findNext, replaceOne, replaceAll } from './find.js';

/** The bar (hidden until opened) and `openFind(replace)`, a command that opens it with or without the replace row. */
export function createFindBar(getView, withView) {
  const findBox = el('input', { type: 'search', placeholder: 'Find', 'aria-label': 'Find', class: 'odx-findbox' });
  const replBox = el('input', { type: 'text', placeholder: 'Replace with', 'aria-label': 'Replace with', class: 'odx-findbox' });
  const matchCase = el('input', { type: 'checkbox', title: 'Match case' });
  const wholeWord = el('input', { type: 'checkbox', title: 'Whole words only' });
  const findInfo = el('span', { class: 'odx-findinfo', role: 'status' }); // read out by screen readers (4.1.3)
  const fOpts = () => ({ matchCase: matchCase.checked, wholeWord: wholeWord.checked });
  const fdo = (fn) => (e) => {
    e.preventDefault();
    const v = getView();
    const n = fn(v);
    findInfo.textContent = n == null ? '' : n ? `${n} result${n === 1 ? '' : 's'}` : 'No results';
  };
  const replRow = el('span', { class: 'odx-replrow' }, replBox,
    el('button', { type: 'button', onmousedown: fdo((v) => replaceOne(v, findBox.value, replBox.value, fOpts())) }, 'Replace'),
    el('button', { type: 'button', onmousedown: fdo((v) => { const n = replaceAll(v, findBox.value, replBox.value, fOpts()); findInfo.textContent = ''; return n; }) }, 'Replace all'));
  const findBar = el('div', { class: 'odx-findbar', hidden: true },
    findBox,
    el('button', { type: 'button', title: 'Previous result', onmousedown: fdo((v) => findNext(v, findBox.value, fOpts(), -1)) }, '↑'),
    el('button', { type: 'button', title: 'Next result', onmousedown: fdo((v) => findNext(v, findBox.value, fOpts(), 1)) }, '↓'),
    el('label', {}, matchCase, 'Match case'), el('label', {}, wholeWord, 'Whole words'),
    replRow, findInfo,
    el('button', { type: 'button', title: 'Close', class: 'odx-close', onmousedown: (e) => { e.preventDefault(); findBar.hidden = true; getView()?.focus(); } }, '✕'));
  findBox.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') fdo((v) => findNext(v, findBox.value, fOpts(), e.shiftKey ? -1 : 1))(e);
    if (e.key === 'Escape') { findBar.hidden = true; getView()?.focus(); }
  });
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
    if (!v || !(e.ctrlKey || e.metaKey) || !v.hasFocus()) return;
    if (e.key === 'f' || e.key === 'h') { e.preventDefault(); openFind(e.key === 'h')(v.state, v.dispatch, v); }
  });
  return { findBar, openFind };
}
