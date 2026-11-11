// Command search, as Word for the web's "Search" box (Alt+Q): type part of a command's name and pick it from the list:
// every command on the ribbon, with its tab and keyboard shortcut. Up/Down move, Enter runs it (a menu opens, a list
// takes the focus), Esc goes back to the document. Read by screen readers as a combo box with its list (ARIA 1.2).
import { el } from './ribbonkit.js';
import { announce } from './live.js';

/**
 * The search box for the ribbon `panels` (tab name -> panel), shown with `showTab(name)`; `reveal(control)` brings a
 * control a narrow ribbon hid behind More into view. Returns { box (the element), open() }.
 */
export function createCommandSearch({ panels, showTab, reveal, getView, container }) {
  const listId = `odx-cs${Math.random().toString(36).slice(2, 8)}`;
  const input = el('input', { type: 'search', class: 'odx-cmdsearch', placeholder: 'Search (Alt+Q)', 'aria-label': 'Search for a command', role: 'combobox', 'aria-autocomplete': 'list', 'aria-expanded': 'false', 'aria-controls': listId, autocomplete: 'off', spellcheck: 'false' });
  const list = el('div', { class: 'odx-cmdlist', role: 'listbox', id: listId, 'aria-label': 'Commands', hidden: true });
  const box = el('div', { class: 'odx-cmdbox' }, input, list);
  let found = [];
  let at = -1;

  /** Every command on the ribbon now: { label, tab, keys, control }. */
  function commands() {
    const out = [];
    const seen = new Set();
    for (const [tab, panel] of Object.entries(panels)) {
      for (const c of panel.querySelectorAll('button, select')) {
        // (Every tab's, the shown one or not; not a group turned off, nor one of its hidden parts.)
        let off = false;
        for (let e = c; e && e !== panel; e = e.parentElement) if (e.hidden) off = true;
        if (off || c.disabled) continue;
        const raw = (c.getAttribute('aria-label') || c.title || c.textContent || '').replace(/\s+/g, ' ').trim();
        if (!raw || raw.length < 2) continue;
        const keys = /\(((?:Ctrl|Alt|Shift|F\d)[^)]*)\)/.exec(raw)?.[1] || null;
        const label = raw.replace(/\s*\((?:Ctrl|Alt|Shift|F\d)[^)]*\)/, '').replace(/:\s*$/, '');
        const k = `${tab}|${label}`;
        if (seen.has(k)) continue;
        seen.add(k);
        out.push({ label, tab, keys, control: c });
      }
    }
    return out;
  }

  /** The commands matching `q`: every word of it in the command's name or tab, those starting with it first. */
  function search(q) {
    const words = q.toLowerCase().split(/\s+/).filter(Boolean);
    if (!words.length) return [];
    const hits = commands().filter((c) => { const t = `${c.label} ${c.tab}`.toLowerCase(); return words.every((w) => t.includes(w)); });
    const first = (c) => (c.label.toLowerCase().startsWith(words[0]) ? 0 : 1);
    return hits.sort((a, b) => first(a) - first(b)).slice(0, 12);
  }

  function draw() {
    list.replaceChildren(...found.map((c, i) => el('div', { role: 'option', id: `${listId}-${i}`, class: `odx-cmdopt${i === at ? ' on' : ''}`, 'aria-selected': String(i === at),
      onmousedown: (e) => { e.preventDefault(); run(i); } },
    el('span', {}, c.label), el('small', {}, [c.keys, c.tab].filter(Boolean).join(' · ')))));
    list.hidden = !found.length;
    input.setAttribute('aria-expanded', String(found.length > 0));
    if (at >= 0) input.setAttribute('aria-activedescendant', `${listId}-${at}`);
    else input.removeAttribute('aria-activedescendant');
  }

  let sayTimer = 0;
  input.addEventListener('input', () => {
    found = search(input.value);
    at = found.length ? 0 : -1;
    draw();
    clearTimeout(sayTimer);
    if (input.value.trim()) sayTimer = setTimeout(() => announce(found.length ? `${found.length} command${found.length === 1 ? '' : 's'}` : 'No commands'), 500);
  });

  const close = (back = true) => {
    found = [];
    at = -1;
    draw();
    input.value = '';
    if (back) getView()?.focus();
  };

  /** Runs found[i]: its tab shown, then the command as a keyboard press on it (a menu opens, a list takes the focus). */
  function run(i) {
    const c = found[i];
    if (!c) return;
    close(false);
    showTab(c.tab);
    reveal?.(c.control);
    if (c.control.tagName === 'SELECT') { c.control.focus(); return; }
    const menu = c.control.getAttribute('aria-haspopup');
    if (menu) c.control.focus();
    c.control.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
    // (A plain command done: back to the text, as after Word's search.)
    if (!menu) setTimeout(() => getView()?.focus(), 0);
  }

  input.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      if (!found.length) return;
      at = (at + (e.key === 'ArrowDown' ? 1 : -1) + found.length) % found.length;
      draw();
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (at >= 0) run(at);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      close();
    }
  });
  input.addEventListener('blur', () => setTimeout(() => { if (!box.contains(document.activeElement)) { found = []; at = -1; draw(); } }, 150));

  // Alt+Q, from the document or the ribbon (as in Word for the web).
  const open = () => { input.focus(); input.select(); };
  document.addEventListener('keydown', (e) => {
    if (!e.altKey || e.ctrlKey || e.metaKey || (e.key !== 'q' && e.key !== 'Q' && e.code !== 'KeyQ')) return;
    const v = getView();
    if (!container.isConnected || container.offsetParent === null || !(v?.hasFocus() || container.contains(document.activeElement))) return;
    e.preventDefault();
    open();
  });
  return { box, open };
}
