// The ribbon's building blocks: elements, buttons, groups, the one drop-down menu, select fields. The toolbar, the
// find bar and the context menu are made from these.

// Line icons drawn in the text color, so they show in light and dark mode (color emoji can be near-black).
const ICON_PATHS = {
  '🖍': 'M15 4l5 5-9 9H6v-5z M4 21h9',
  '🖌': 'M18 3l3 3-9 9-3-3z M9 12c-3 0-4 2-4 4 0 1.5-1 2.5-2 3 3 1 7 0 7-4',
  '🔗': 'M10 14a4 4 0 0 0 5.66 0l3-3a4 4 0 0 0-5.66-5.66l-1 1 M14 10a4 4 0 0 0-5.66 0l-3 3a4 4 0 0 0 5.66 5.66l1-1',
  '💬': 'M4 5h16v11H9l-5 4z',
  '🖶': 'M7 9V3h10v6 M7 17H4V9h16v8h-3 M7 14h10v7H7z',
  '🗑': 'M4 7h16 M9 7V4h6v3 M6 7l1 13h10l1-13',
  '🔍': 'M10.5 4a6.5 6.5 0 1 0 0 13a6.5 6.5 0 1 0 0-13z M15.5 15.5L21 21',
  '📋': 'M8 3h8v4H8z M6 5H5v16h14V5h-1',
  '✂': 'M3 6a3 3 0 1 0 6 0a3 3 0 1 0-6 0 M3 18a3 3 0 1 0 6 0a3 3 0 1 0-6 0 M8.5 7.5L20 19 M8.5 16.5L20 5',
  '📂': 'M3 6h6l2 2h10v11H3z',
  '🖼': 'M3 5h18v14H3z M3 16l5-5 4 4 3-3 6 6',
  '🔖': 'M6 3h12v18l-6-5-6 5z',
  '📖': 'M3 5h7a2 2 0 0 1 2 2v13a2 2 0 0 0-2-2H3z M21 5h-7a2 2 0 0 0-2 2v13a2 2 0 0 1 2-2h7z',
  '📏': 'M3 15L15 3l6 6L9 21z M7 11l2 2 M10 8l2 2 M13 5l2 2',
  '🎨': 'M12 3a9 9 0 1 0 0 18c1.5 0 2-1 2-2s-1-2 0-3 4 1 6-2a9 9 0 0 0-8-11z M8 10h.01 M12 7h.01 M16 10h.01',
  '♿': 'M12 4a1.5 1.5 0 1 0 0 3a1.5 1.5 0 1 0 0-3z M6 9h12 M12 9v5l-3 6 M12 14l3 6',
};
export function iconize(root) {
  root.querySelectorAll('.odx-ico, .odx-ctxico, .odx-mini > button').forEach((e) => {
    const d = ICON_PATHS[e.textContent.trim()];
    if (!d) return;
    e.innerHTML = `<svg viewBox="0 0 24 24" width="1.1em" height="1.1em" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" style="vertical-align:-0.15em"><path d="${d}"/></svg>`;
  });
}

export function el(tag, attrs = {}, ...children) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    // Controls act on mouse-down (so the editor keeps its selection); Enter and Space do the same from the keyboard.
    if (k === 'onmousedown' && v) e.addEventListener('keydown', (ev) => {
      if (ev.key !== 'Enter' && ev.key !== ' ') return;
      const r = e.getBoundingClientRect();
      ev.preventDefault();
      e.odxKeyAt = performance.now();
      v({ preventDefault() {}, stopPropagation() {}, target: e, currentTarget: e, clientX: r.left, clientY: r.bottom, button: 0, detail: 1, viaKeyboard: true });
    });
    // A screen reader in browse mode activates with a bare click (no mouse-down), which acts like a key press.
    if (k === 'onmousedown' && v) e.addEventListener('click', (ev) => {
      if (ev.detail !== 0 || performance.now() - (e.odxKeyAt || -1e9) < 500 || performance.now() - (e.odxDownAt || -1e9) < 1000) return;
      const r = e.getBoundingClientRect();
      v({ preventDefault() {}, stopPropagation() {}, target: e, currentTarget: e, clientX: r.left, clientY: r.bottom, button: 0, detail: 1, viaKeyboard: true });
    });
    if (k === 'onmousedown' && v) e.addEventListener('mousedown', () => { e.odxDownAt = performance.now(); }, true);
    else if (v != null && v !== false) e.setAttribute(k, v === true ? '' : v);
  }
  e.append(...children.filter((c) => c != null && c !== false));
  return e;
}

/** A tab or group label as a feature name: 'Header & Footer' -> 'header-and-footer'. */
export const slug = (t) => String(t || '').toLowerCase().replace(/&/g, 'and').replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');

/**
 * The controls of one ribbon, bound to its container (where the menu is placed) and to the current view.
 * A command is a ProseMirror command (state, dispatch, view).
 */
export function ribbonKit(container, getView) {
  let lastPointer = -1e9;
  document.addEventListener('pointerdown', () => { lastPointer = performance.now(); }, true);
  /** A command as a mouse-down handler. */
  const run = (command) => (e) => {
    e.preventDefault();
    closeMenu();
    const view = getView();
    if (!view) return;
    command(view.state, view.dispatch, view);
    // From the keyboard, focus stays on the button (as in Word) so its new state is heard; a click returns to the text.
    // A screen reader's click (Space in browse mode) has no pointer behind it, so it counts as keyboard too.
    // So does pressing a button that already had focus (a screen reader's click is a mouse event, but the user tabbed there).
    const keys = e.viaKeyboard || performance.now() - lastPointer > 500 || document.activeElement === e.target?.closest?.('button');
    // (Not an item of a menu, which has just closed: back to the document, as in Word. Its hidden item had kept focus:
    // NVDA said "document" after Review > Accept > Accept this change.)
    const inMenu = !!e.target?.closest?.('.odx-pop');
    if (keys && !inMenu && e.target?.isConnected) e.target.closest('button')?.focus(); else view.focus();
  };
  /** A command from a function of the view (only run when there is one). */
  const withView = (fn) => (state, dispatch, view) => (view ? fn(view) : true);

  /** A ribbon button: `icon` is a short glyph, `label` shows under it when `big`. */
  const btn = (icon, title, command, { label = null, big = false, disabled = false, cls = '' } = {}) =>
    el('button', { type: 'button', title, 'aria-label': label && !title.toLowerCase().includes(String(label).toLowerCase()) ? `${label}: ${title}` : title, class: `${big ? 'odx-big' : ''} ${cls}`.trim() || null, disabled, onmousedown: disabled ? null : run(command) },
      el('span', { class: 'odx-ico' }, icon), label && el('span', { class: 'odx-lbl' }, label));
  const grp = (label, ...kids) => el('div', { class: 'odx-group', 'data-feature': slug(label) }, el('div', { class: 'odx-row' }, ...kids), label && el('div', { class: 'odx-glabel' }, label));
  const col = (...kids) => el('div', { class: 'odx-col' }, ...kids);
  const line = (...kids) => el('div', { class: 'odx-line' }, ...kids);

  // --- Dropdown menus (one open at a time) ---
  const pop = el('div', { class: 'odx-pop', hidden: true });
  let popOwner = null;
  pop.setAttribute('role', 'menu');
  function closeMenu(refocus = false) { const o = popOwner; pop.hidden = true; popOwner = null; o?.setAttribute('aria-expanded', 'false'); if (refocus) o?.focus(); }
  pop.addEventListener('keydown', (e) => {
    const items = [...pop.querySelectorAll('button:not([disabled]), input, select, textarea')];
    const i = items.indexOf(document.activeElement);
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeMenu(true); }
    else if (e.key === 'ArrowDown' || e.key === 'ArrowRight') { e.preventDefault(); items[(i + 1) % items.length]?.focus(); }
    else if (e.key === 'ArrowUp' || e.key === 'ArrowLeft') { e.preventDefault(); items[(i - 1 + items.length) % items.length]?.focus(); }
  });
  /** Opens the menu under `owner` with the elements `build()` returns (a second press on the owner closes it). */
  function openMenu(owner, build) {
    if (popOwner === owner) return closeMenu();
    pop.replaceChildren(...build());
    iconize(pop);
    pop.hidden = false;
    popOwner = owner;
    owner.setAttribute('aria-expanded', 'true');
    pop.setAttribute('aria-label', owner.getAttribute('aria-label') || owner.textContent);
    // (One with fields to fill in, as Restrict Editing's, is a dialog rather than a menu.)
    const form = !!pop.querySelector('input, select, textarea');
    pop.setAttribute('role', form ? 'dialog' : 'menu');
    pop.querySelectorAll('button').forEach((b) => { if (!form && !b.getAttribute('role')) b.setAttribute('role', 'menuitem'); if (!b.getAttribute('aria-label') && b.title && !b.textContent.trim()) b.setAttribute('aria-label', b.title); });
    if (owner.contains(document.activeElement) || owner === document.activeElement) setTimeout(() => pop.querySelector('button:not([disabled]), input, select')?.focus(), 300); // after the reader has said "expanded"
    const r = owner.getBoundingClientRect();
    const host = container.getBoundingClientRect();
    pop.style.left = `${Math.max(0, Math.min(r.left - host.left, host.width - 240))}px`;
    pop.style.top = `${r.bottom - host.top + 2}px`;
  }
  document.addEventListener('mousedown', (e) => { if (!pop.hidden && !pop.contains(e.target) && !popOwner?.contains(e.target)) closeMenu(); });
  const item = (label, command, { on = false, hint = '' } = {}) => el('button', { type: 'button', class: `odx-item${on ? ' on' : ''}`, onmousedown: run(command) }, el('span', {}, label), hint && el('small', {}, hint));
  /** A button with a ▾ that opens `items()` (an array of elements). */
  const menu = (icon, title, items, opts = {}) => {
    const b = el('button', { type: 'button', title, 'aria-label': title, 'aria-haspopup': 'menu', 'aria-expanded': 'false', class: `odx-menu ${opts.big ? 'odx-big' : ''}`, onmousedown: (e) => { e.preventDefault(); openMenu(b, items); } },
      el('span', { class: 'odx-ico' }, icon), opts.label && el('span', { class: 'odx-lbl' }, opts.label), el('span', { class: 'odx-caret' }, '▾'));
    return b;
  };
  /** A select whose choice goes to `onPick(value, view)`, then back to the document. */
  const pick = (title, cls, options, onPick) => el('select', { title, 'aria-label': title, class: cls, onchange: (e) => { const v = getView(); onPick(e.target.value, v); v.focus(); } }, ...options);

  return { run, withView, btn, grp, col, line, pop, openMenu, closeMenu, item, menu, pick };
}
