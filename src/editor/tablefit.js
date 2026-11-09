// Autofit tables (no fixed layout): Word never breaks a word inside a cell; a column too narrow for its longest word
// is widened to fit it, and the other columns give up the difference (each in proportion to its room above its own
// longest word), so the table keeps its width. Browsers only widen, so the column widths are worked out here: a copy
// of each table is laid out at its narrowest to measure every column's minimum, and the real columns are adjusted.
// The copies are laid out together, once a pass (each on its own forced a layout of the whole document: 120 tables in
// a 440-page document took 6 s), and a table's minimums are kept until its text changes.

const memo = new WeakMap(); // table element -> { sig, min }

/** Adjusts the columns of autofit tables; true when any width changed (measure again once drawn). */
export function fitAutofitColumns(view) {
  let changed = false;
  const host = view.dom.parentElement;
  const todo = [];
  for (const table of view.dom.querySelectorAll(':scope > .tableWrapper > table, :scope table')) {
    if (table.hasAttribute('data-fixed') || table.classList.contains('odx-autofit') || table.closest('.odx-repeat')) continue;
    const cols = [...(table.querySelector(':scope > colgroup')?.children || [])];
    if (cols.length < 2) continue;
    const cur = cols.map((c) => parseFloat(c.style.width) || 0);
    if (cur.some((w) => !w)) continue;
    // (The minimums depend on the cells' text alone, not the widths: the copy is laid out without them.)
    const sig = `${table.rows.length}|${table.textContent}`;
    const m = memo.get(table);
    todo.push({ table, cols, cur, sig, min: m && m.sig === sig ? m.min : undefined });
  }
  const fresh = todo.filter((t) => t.min === undefined);
  if (fresh.length) {
    const mins = minimums(fresh.map((t) => t.table), fresh.map((t) => t.cols.length), host, view.dom.className);
    fresh.forEach((t, k) => { t.min = mins[k]; memo.set(t.table, { sig: t.sig, min: t.min }); });
  }
  for (const { cols, cur, min } of todo) {
    if (!min) continue;
    const over = min.map((m, i) => Math.max(0, m - cur[i]));
    const need = over.reduce((a, b) => a + b, 0);
    if (need < 0.5) continue;
    const room = cur.map((w, i) => (over[i] ? 0 : Math.max(0, w - min[i])));
    const spare = room.reduce((a, b) => a + b, 0);
    if (spare < need) continue; // nothing to take it from: leave it as the grid says
    cols.forEach((col, i) => {
      const w = over[i] ? min[i] : cur[i] - (room[i] / spare) * need;
      col.style.width = `${Math.round(w * 100) / 100}px`;
    });
    changed = true;
  }
  return changed;
}

/** Each table's columns' narrowest widths (px): its longest word plus the cell's padding, from cells spanning one
 *  column; null for a table with none. All the tables' copies laid out at once. */
function minimums(tables, ns, host, contentClass) {
  const box = document.createElement('div');
  box.className = contentClass;
  box.setAttribute('aria-hidden', 'true');
  box.style.cssText = 'position:absolute;left:-99999px;top:0;width:auto;visibility:hidden;contain:layout';
  const copies = tables.map((table) => {
    const copy = table.cloneNode(true);
    copy.style.tableLayout = 'auto';
    copy.style.width = '0';
    copy.style.minWidth = '0';
    for (const col of copy.querySelectorAll(':scope > colgroup > col')) col.style.width = '';
    // Cells spanning columns emptied: laid out at its narrowest, the table would spread their long text over the columns
    // they span, and a column would seem to need far more than its own cells (sample S099 form: a checkbox column, 25px, at 77px).
    for (const tr of copy.rows) for (const td of tr.cells) if (td.colSpan > 1) td.replaceChildren();
    const wrap = document.createElement('div'); // (one per table: each laid out at its own narrowest)
    wrap.append(copy);
    box.append(wrap);
    return copy;
  });
  host.append(box);
  const out = copies.map((copy, t) => {
    const n = ns[t];
    const min = new Array(n).fill(0);
    let seen = 0;
    // Each cell's own grid column: past the columns a merged cell from a row above still covers (its rowspan).
    const covered = new Array(n).fill(0); // rows each column stays covered for
    for (const tr of copy.rows) {
      let c = 0;
      for (const td of tr.cells) {
        while (c < n && covered[c] > 0) c++;
        const span = td.colSpan || 1;
        if (span === 1 && c < n) { min[c] = Math.max(min[c], td.getBoundingClientRect().width); seen++; }
        for (let k = c; k < Math.min(n, c + span); k++) covered[k] = td.rowSpan || 1;
        c += span;
      }
      for (let k = 0; k < n; k++) if (covered[k] > 0) covered[k]--;
    }
    return seen ? min : null;
  });
  box.remove();
  return out;
}
