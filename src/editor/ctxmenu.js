// Right-click: Word's mini toolbar over its context menu (also Shift+F10 and the Menu key).
import { el, iconize } from './ribbonkit.js';
import { editChart, chartXmlAt } from './charts.js';
import * as cm from './comments.js';
import * as cmd from './commands.js';
import { WRAPS } from '../ooxml/wrap.js';
import { setWrapText, currentWrap } from './wraptext.js';

/**
 * The menu's element (drawn into the ribbon's container, hidden until opened). `kit` is the ribbon's (ribbonkit.js);
 * `ribbon` is what it borrows from the toolbar: its font and size fields (the mini toolbar starts from their values),
 * its clipboard, Format Painter and Bullets commands, `showTab(name)` and `panel(name)`, `editLink`, and
 * `comments()`, whether comments are on.
 */
export function createContextMenu(container, getView, kit, ribbon) {
  const { withView, pick, closeMenu } = kit;
  const { fonts, sizes, fontSel, sizeSel, clip, pasteFrom, formatPainter, bullets, showTab, panel, editLink, comments } = ribbon;
  const ctx = el('div', { class: 'odx-ctx', hidden: true });
  function closeCtx(refocus = false) { if (ctx.hidden) return; ctx.hidden = true; ctx.replaceChildren(); if (refocus) getView()?.focus(); }
  const ctxRun = (command) => (e) => { e.preventDefault(); closeCtx(); const v = getView(); if (!v) return; v.focus(); command(v.state, v.dispatch, v); };
  const ctxItem = (icon, label, keys, command, { disabled = false } = {}) => el('button', { type: 'button', role: 'menuitem', class: 'odx-ctxitem', disabled, onmousedown: disabled ? null : ctxRun(command) },
    el('span', { class: 'odx-ctxico', 'aria-hidden': 'true' }, icon), el('span', { class: 'odx-ctxlbl' }, label), el('small', {}, keys || ''));
  function openCtx(e) {
    const v = getView();
    if (!v || container.style.display === 'none' || v.odx?.readOnly) return;
    e.preventDefault();
    closeMenu();
    const { state } = v;
    const empty = state.selection.empty;
    const fsel = pick('Font', 'odx-font', [el('option', { value: '' }, ''), ...fonts.map((f) => el('option', { value: f }, f))], (val, w) => cmd.setMark('font', val ? { val } : null)(w.state, w.dispatch));
    fsel.append(...[...fontSel.options].filter((o) => ![...fsel.options].some((x) => x.value === o.value)).map((o) => el('option', { value: o.value }, o.textContent)));
    fsel.value = fontSel.value;
    const ssel = pick('Font size', 'odx-size', [el('option', { value: '' }, ''), ...sizes.map((x) => el('option', { value: x }, String(x)))], (val, w) => cmd.setMark('size', val ? { val: Math.round(Number(val) * 2) } : null)(w.state, w.dispatch));
    if (![...ssel.options].some((o) => o.value === sizeSel.value)) ssel.append(el('option', { value: sizeSel.value }, sizeSel.value));
    ssel.value = sizeSel.value;
    const tb = (icon, title, command, on = false) => el('button', { type: 'button', title, 'aria-label': title, 'aria-pressed': on ? 'true' : null, class: on ? 'on' : null, onmousedown: ctxRun(command) }, icon);
    const mini = el('div', { class: 'odx-mini', role: 'toolbar', 'aria-label': 'Mini toolbar' },
      fsel, ssel,
      tb('B', 'Bold (Ctrl+B)', cmd.toggle('bold'), cmd.markActive(state, 'bold')),
      tb('I', 'Italic (Ctrl+I)', cmd.toggle('italic'), cmd.markActive(state, 'italic')),
      tb('U', 'Underline (Ctrl+U)', cmd.toggle('underline'), cmd.markActive(state, 'underline')),
      tb('A', 'Font color: red', cmd.setMark('color', { val: 'FF0000' })),
      tb('🖍', 'Highlight: yellow', cmd.setMark('highlight', { val: 'yellow' })),
      tb('🖌', 'Format Painter', formatPainter),
      tb('•≡', 'Bullets', bullets),
      tb('Aa', 'Styles', withView(() => { showTab('Home'); setTimeout(() => container.querySelector('.odx-stile')?.focus()); return true; })),
      !comments() ? null : tb('💬', 'New comment', withView(cm.addComment)));
    mini.querySelectorAll('[title="Font color: red"]').forEach((b) => b.classList.add('odx-redline'));
    // (A chart that's selected: its data, as double-clicking it does.)
    const selView = getView();
    const chartAt = selView && state.selection.node && chartXmlAt(selView, state.selection.node) ? state.selection.from : null;
    const wrapNow = currentWrap(state);
    const list = el('div', { class: 'odx-ctxmenu', role: 'menu', 'aria-label': 'Context menu' },
      ctxItem('✂', 'Cut', 'Ctrl+X', clip('cut'), { disabled: empty }),
      ctxItem('⧉', 'Copy', 'Ctrl+C', clip('copy'), { disabled: empty }),
      ctxItem('📋', 'Paste', 'Ctrl+V', pasteFrom(false)),
      ctxItem('Ⓐ', 'Paste Text Only', 'Ctrl+Shift+V', pasteFrom(true)),
      el('hr'),
      ctxItem('¶', 'Paragraph Options', '', withView(() => { showTab('Layout'); setTimeout(() => panel('Layout').querySelector('input')?.focus()); return true; })),
      chartAt != null ? ctxItem('📊', 'Edit Chart Data...', '', withView((w) => editChart(w, chartAt))) : null,
      // A selected picture or shape: its Wrap Text choices (the one it has ticked).
      ...(wrapNow ? [el('hr'), el('div', { class: 'odx-ctxhead', role: 'presentation' }, 'Wrap Text'), ...WRAPS.map(([k, n]) => ctxItem(k === wrapNow ? '✓' : '', n, '', setWrapText(k)))] : []),
      el('hr'),
      ctxItem('🔗', 'Link', 'Ctrl+K', withView((w) => { showTab('Insert'); return editLink(w.state, w.dispatch, w); })),
      !comments() ? null : ctxItem('💬', 'New Comment', 'Ctrl+Alt+M', withView(cm.addComment)));
    ctx.replaceChildren(mini, list);
    iconize(ctx);
    ctx.hidden = false;
    const host = container.getBoundingClientRect();
    const w = ctx.offsetWidth, h = ctx.offsetHeight;
    const x = Math.min(e.clientX, window.innerWidth - w - 8);
    const y = e.clientY + h > window.innerHeight - 8 ? Math.max(8, e.clientY - h) : e.clientY - mini.offsetHeight - 8;
    ctx.style.left = `${x - host.left}px`;
    ctx.style.top = `${Math.max(8 - host.top, y - host.top)}px`;
    if (e.viaKeyboard || e.button !== 2) list.querySelector('button:not([disabled])')?.focus();
  }
  ctx.addEventListener('keydown', (e) => {
    const items = [...list().querySelectorAll('button:not([disabled])')];
    const i = items.indexOf(document.activeElement);
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeCtx(true); }
    else if (e.key === 'ArrowDown') { e.preventDefault(); items[(i + 1) % items.length]?.focus(); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); items[(i - 1 + items.length) % items.length]?.focus(); }
  });
  const list = () => ctx.querySelector('.odx-ctxmenu');
  document.addEventListener('mousedown', (e) => { if (!ctx.contains(e.target)) closeCtx(); }, true);
  document.addEventListener('scroll', () => closeCtx(), true);
  // (Opened with the mouse, focus stays in the document: Esc there closes it too, as in Word.)
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !ctx.hidden && !ctx.contains(e.target)) { e.preventDefault(); closeCtx(true); } }, true);
  document.addEventListener('contextmenu', (e) => {
    const v = getView();
    const main = v?.odx?.main || v;
    if (!main) return;
    // From the keyboard (Shift+F10, Menu key) the event can land on the page body, e.g. when a screen reader is
    // in browse mode: open at the cursor if the document was the last thing focused.
    const fromKeys = e.button !== 2 && e.clientX === 0 && e.clientY === 0;
    if (fromKeys && (e.target === document.body || main.dom.parentElement.contains(e.target)) && lastFocusDoc) {
      e.preventDefault();
      v.focus();
      const c = v.coordsAtPos(v.state.selection.head);
      openCtx({ preventDefault() {}, clientX: c.left, clientY: c.bottom, viaKeyboard: true });
      return;
    }
    if (!main.dom.parentElement.contains(e.target)) return;
    // Right-clicking outside the selection moves the cursor there first, as in Word.
    const at = v.posAtCoords({ left: e.clientX, top: e.clientY });
    const { from, to } = v.state.selection;
    if (at && (at.pos < from || at.pos > to)) v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.near(v.state.doc.resolve(at.pos))));
    openCtx(e);
  });
  let lastFocusDoc = false;
  document.addEventListener('focusin', (e) => { const v = getView(); lastFocusDoc = !!v && (v.dom.contains(e.target) || (v.odx?.main || v).dom.contains(e.target)); });
  // Shift+F10 and the Menu key open it from the keyboard.
  document.addEventListener('keydown', (e) => {
    if (!((e.key === 'F10' && e.shiftKey) || e.key === 'ContextMenu')) return;
    const v = getView();
    if (!v || !v.hasFocus()) return;
    const c = v.coordsAtPos(v.state.selection.head);
    openCtx({ preventDefault: () => e.preventDefault(), clientX: c.left, clientY: c.bottom, viaKeyboard: true });
  });
  return ctx;
}
