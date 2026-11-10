// Table rows as Word sizes them where the browser doesn't: a merged cell's extra height in the last of its rows, and
// exact-height rows that cut their content instead of growing.

const fitted = new WeakMap(); // table element -> what it was fitted for (rows, width, text)

/**
 * A cell merged down several rows that needs more height than those rows give: Word puts all the extra in the last of
 * them (sample S099's header: the logo cell spans "Title" and "Document No.", and only the second row grows); the
 * browser spreads it over all of them. Each such last row gets the height that takes it, from the rows' own content.
 * True when a height changed.
 */
export function fitRowspans(root) {
  let changed = false;
  const own = (td) => {
    const last = td.lastElementChild;
    if (!last) return 0;
    const cs = getComputedStyle(td);
    const r = td.getBoundingClientRect();
    return last.getBoundingClientRect().bottom + (parseFloat(getComputedStyle(last).marginBottom) || 0) - r.top + (parseFloat(cs.paddingBottom) || 0);
  };
  for (const table of root.querySelectorAll('table')) {
    // (Only tables with cells merged down, and of those only one whose text or width changed since it was fitted:
    // measuring every table each pass took 450ms a keystroke in a 20MB document of tables.)
    if (!table.querySelector('td[rowspan], th[rowspan]')) continue;
    const sig = `${document.fonts?.status}|${table.rows.length}|${table.offsetWidth}|${table.textContent}`; // (again once fonts load)
    if (fitted.get(table) === sig) continue;
    fitted.set(table, sig);
    const rows = [...table.rows];
    if (rows.some((tr) => tr.hasAttribute('data-exact'))) continue;
    const need = new Map(); // row index -> extra height its last-row role needs
    let any = false;
    for (let i = 0; i < rows.length; i++) {
      for (const td of rows[i].cells) {
        if (td.rowSpan < 2 || i + td.rowSpan > rows.length) continue;
        any = true;
        const span = rows.slice(i, i + td.rowSpan);
        const natural = span.map((tr) => Math.max(0, ...[...tr.cells].filter((c) => c.rowSpan === 1).map(own)));
        const want = own(td) - natural.reduce((a, b) => a + b, 0);
        const k = i + td.rowSpan - 1;
        if (want > 0.5) need.set(k, Math.max(need.get(k) || 0, natural[natural.length - 1] + want));
      }
    }
    if (!any) continue;
    rows.forEach((tr, k) => {
      const h = need.has(k) ? `${need.get(k).toFixed(2)}px` : '';
      if (tr.style.height !== h) { tr.style.height = h; changed = true; }
    });
  }
  return changed;
}

/**
 * Exact-height rows in a drawn header or footer: a cell holding more than the row doesn't grow it (Word keeps the row's
 * height and cuts the cell at its edge: sample S109's header rows add up exactly in Word's PDF). Its paragraphs after the first
 * take no room; centred or bottom-aligned content moves by the offset all of it would have.
 */
export function clipExactRows(root) {
  // A table ending in an exact row has its bottom border below that row, not in it (sample S108's first-page header: the
  // 1.92px rule under the last 20.13px row, the body 1.9px lower than ours).
  for (const table of root.querySelectorAll('table')) {
    const last = table.rows[table.rows.length - 1];
    if (!last || !last.hasAttribute('data-exact')) continue;
    const below = Math.max(0, ...[...last.cells].map((td) => { const m = /-(\d+(?:\.\d+)?)px 0px 0px inset/.exec(td.style.boxShadow || ''); return m ? Number(m[1]) : 0; }));
    // (Added to that row, its border drawn in the band: a margin would collapse into the next paragraph's space before,
    // the first-page footer's "Page 1 of 8".)
    if (below) last.style.setProperty('--odx-trh-add', `${below}px`);
  }
  for (const tr of root.querySelectorAll('tr[data-exact]')) {
    const h = parseFloat(tr.style.getPropertyValue('--odx-trh'));
    if (!h) continue;
    for (const td of tr.cells) {
      const ps = [...td.children].filter((e) => e.tagName === 'P');
      if (ps.length < 2 || td.rowSpan > 1) continue; // (a cell merged down several rows has their height, not this one's)
      const cs = getComputedStyle(td);
      const room = h - (parseFloat(cs.paddingTop) || 0) - (parseFloat(cs.paddingBottom) || 0);
      const first = ps[0].getBoundingClientRect();
      const total = ps[ps.length - 1].getBoundingClientRect().bottom - first.top;
      if (total <= room + 0.5) continue;
      const shown = Math.min(first.height, room);
      const lift = cs.verticalAlign === 'middle' ? (total - shown) / 2 : cs.verticalAlign === 'bottom' ? total - shown : 0;
      ps.slice(1).forEach((p) => { p.style.maxHeight = '0'; p.style.minHeight = '0'; });
      // (By a transform: a relative top would read as the print baseline shift, and printlines takes that back off.)
      if (lift) ps.forEach((p) => { p.style.transform = `translateY(${-lift}px)`; });
    }
  }
}
