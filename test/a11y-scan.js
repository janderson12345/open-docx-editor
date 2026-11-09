// Runs in the page: accessibility pattern scan over the rendered editor UI (method from the WBS VPAT work).
window.__a11yScan = (root = document.querySelector('#app') || document.body) => {
  const out = {};
  const add = (k, el, why = '') => (out[k] ??= []).push(`${el.tagName.toLowerCase()}.${(el.className && el.className.baseVal === undefined ? el.className : '').toString().split(' ').filter(Boolean).slice(0, 2).join('.')} "${(el.getAttribute('aria-label') || el.title || el.textContent || '').trim().slice(0, 40)}" ${why}`);
  const vis = (e) => e.getClientRects().length > 0 && getComputedStyle(e).visibility !== 'hidden';
  const name = (e) => (e.getAttribute('aria-label') || '').trim() || (e.getAttribute('aria-labelledby') ? 'lb' : '') || [...e.childNodes].map((n) => (n.nodeType === 3 ? n.textContent : n.getAttribute?.('aria-hidden') === 'true' ? '' : n.textContent)).join('').trim();
  const doc = root.querySelector('.odx-content');
  const ui = [...root.querySelectorAll('button, [role=button], [role=tab], [role=menuitem], select, input, textarea, a[href]')].filter((e) => !doc?.contains(e) && vis(e));
  out._counts = { controls: ui.length };
  for (const e of ui) {
    const n = name(e);
    const isField = /^(SELECT|INPUT|TEXTAREA)$/.test(e.tagName);
    if (isField) {
      const labelled = e.getAttribute('aria-label') || e.getAttribute('aria-labelledby') || e.closest('label') || (e.id && root.querySelector(`label[for="${e.id}"]`));
      if (!labelled && e.type !== 'hidden') add(e.placeholder ? 'field: placeholder-only name' : 'field: no label', e);
    } else if (!n) add(e.title ? 'button: named only by title' : 'button: no name', e);
    // aria-label must contain the visible text (2.5.3)
    const visible = [...e.childNodes].filter((c) => c.nodeType === 3 || (c.getAttribute?.('aria-hidden') !== 'true' && !c.matches?.('.odx-ico,.odx-ctxico,.odx-caret,svg'))).map((c) => c.textContent).join('').trim();
    const al = e.getAttribute('aria-label');
    if (e.tagName !== 'SELECT' && al && visible && /[a-z]{3}/i.test(visible) && !al.toLowerCase().includes(visible.toLowerCase())) add('aria-label lacks visible text', e, `visible="${visible.slice(0, 30)}"`);
    const r = e.getBoundingClientRect();
    if (r.width < 24 || r.height < 24) add('target under 24px', e, `${Math.round(r.width)}x${Math.round(r.height)}`);
    if (e.tabIndex > 0) add('positive tabindex', e);
    if (e.closest('[aria-hidden=true]')) add('focusable inside aria-hidden', e);
  }
  root.querySelectorAll('[role=tab]').forEach((e) => { if (!e.hasAttribute('aria-selected')) add('tab without aria-selected', e); });
  root.querySelectorAll('[role=dialog],[role=menu],[role=toolbar],[role=region]').forEach((e) => { if (vis(e) && !e.getAttribute('aria-label') && !e.getAttribute('aria-labelledby')) add('container role with no name', e); });
  root.querySelectorAll('[aria-labelledby],[aria-describedby],[aria-controls]').forEach((e) => {
    for (const a of ['aria-labelledby', 'aria-describedby', 'aria-controls']) for (const id of (e.getAttribute(a) || '').split(/\s+/).filter(Boolean)) if (!document.getElementById(id)) add('broken id reference', e, `${a}=${id}`);
  });
  [...root.querySelectorAll('svg')].filter((s) => !doc?.contains(s) && vis(s)).forEach((s) => { if (!s.closest('[aria-hidden=true]') && s.getAttribute('aria-hidden') !== 'true' && !s.getAttribute('role')) add('svg not hidden or named', s); });
  [...root.querySelectorAll('img')].filter((i) => !doc?.contains(i)).forEach((i) => { if (!i.hasAttribute('alt')) add('img without alt', i); });
  // Clickable things that aren't controls (2.1.1)
  [...root.querySelectorAll('div, span, li, p')].filter((e) => !doc?.contains(e) && vis(e) && (e.onclick || e.onmousedown) && !e.getAttribute('role') && e.tabIndex < 0).forEach((e) => add('mouse-only element', e));
  // Status messages need a live region (4.1.3)
  out._counts.liveRegions = root.querySelectorAll('[role=status],[role=alert],[aria-live]').length;
  return out;
};
