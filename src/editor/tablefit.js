// Autofit tables (no fixed layout): Word never breaks a word inside a cell; a column too narrow for its longest word
// is widened to fit it, and the other columns give up the difference (each in proportion to its room above its own
// longest word), so the table keeps its width. Browsers only widen, so the column widths are worked out here: a copy
// of each table is laid out at its narrowest to measure every column's minimum, and the real columns are adjusted.
// The copies are laid out together, once a pass (each on its own forced a layout of the whole document: 120 tables in
// a 440-page document took 6 s), and a table's minimums are kept until its text changes.

const memo = new WeakMap(); // table element -> { sig, min }
// A column's narrowest width is its longest word's in the fonts as loaded: measured before a font arrived (the first pass
// of a document just opened), it was kept, and the columns stayed as the file's cells give them or went to the grid's
// from one opening to the next (a five-column table, its second column 74.6px for a word of 83.8px: its header row one
// line taller every few openings). Measured again when a font has loaded since, and from the widths the table came with.
let fontLoads = 0;
if (typeof document !== 'undefined' && document.fonts) document.fonts.addEventListener('loadingdone', () => { fontLoads++; });
const given = new WeakMap(); // table element -> { from: the widths it came with, set: the widths put on it here }

/**
 * A table drawn again by its view has its columns back at the widths it came with: the fitted ones put back at once
 * (the next pass that would fit them again doesn't always come: the same document opened with the fitted columns most
 * times and the file's own a few, a header row a line taller for it).
 */
export function keepFitted(table) {
  const g = given.get(table);
  if (!g || g.set === g.from) return;
  const cols = [...(table.querySelector(':scope > colgroup')?.children || [])];
  if (cols.length !== g.from.length || !cols.every((c, i) => Math.abs((parseFloat(c.style.width) || 0) - g.from[i]) < 0.011)) return;
  cols.forEach((c, i) => { c.style.width = `${g.set[i]}px`; });
}

/** Adjusts the columns of autofit tables; true when any width changed (measure again once drawn). */
export function fitAutofitColumns(view) {
  let changed = false;
  const host = view.dom.parentElement;
  const all = [];
  for (const table of view.dom.querySelectorAll(':scope > .tableWrapper > table, :scope table')) {
    if (table.hasAttribute('data-fixed') || table.classList.contains('odx-autofit') || table.closest('.odx-repeat')) continue;
    const cols = [...(table.querySelector(':scope > colgroup')?.children || [])];
    if (cols.length < 2) continue;
    let depth = 0;
    for (let e = table.parentElement?.closest('td'); e; e = e.parentElement?.closest('td')) depth++;
    all.push({ table, cols, depth });
  }
  // Tables in cells first, the deepest first: a nested table widened for its longest words makes its cell's column
  // need more, so the table round it is measured with it as it now is.
  for (let depth = Math.max(0, ...all.map((t) => t.depth)); depth >= 0; depth--) {
    const todo = [];
    for (const { table, cols, depth: d } of all) {
      if (d !== depth) continue;
      let cur = cols.map((c) => parseFloat(c.style.width) || 0);
      if (cur.some((w) => !w)) continue;
      // (Fitted from the widths it came with, not from an earlier fitting's: those are still on it unless something
      // else has set its columns since.)
      const g = given.get(table);
      if (g && g.set.length === cur.length && g.set.every((w, i) => Math.abs(w - cur[i]) < 0.011)) cur = g.from;
      else given.set(table, { from: cur, set: cur });
      // (The minimums depend on the cells' text alone, not the widths: the copy is laid out without them. And on the
      // widths of the tables in its cells.)
      const nested = [...table.querySelectorAll('td table > colgroup > col')].map((c) => c.style.width).join(',');
      const sig = `${fontLoads}|${table.rows.length}|${table.textContent}|${nested}`;
      const m = memo.get(table);
      todo.push({ table, cols, cur, sig, nest: depth > 0, min: m && m.sig === sig ? m.min : undefined });
    }
    const fresh = todo.filter((t) => t.min === undefined);
    if (fresh.length) {
      const mins = minimums(fresh.map((t) => t.table), fresh.map((t) => t.cols.length), host, view.dom.className);
      fresh.forEach((t, k) => { t.min = mins[k]; memo.set(t.table, { sig: t.sig, min: t.min }); });
    }
    for (const { table, cols, cur, min, nest } of todo) {
      if (!min) continue;
      const before = cols.map((c) => c.style.width);
      const put = (ws) => {
        cols.forEach((col, i) => { col.style.width = `${Math.round(ws[i] * 100) / 100}px`; });
        given.set(table, { from: cur, set: cols.map((c) => parseFloat(c.style.width) || 0) });
        if (cols.some((c, i) => c.style.width !== before[i])) changed = true;
      };
      const over = min.map((m, i) => Math.max(0, m - cur[i]));
      const need = over.reduce((a, b) => a + b, 0);
      if (need < 0.5) { put(cur); continue; }
      const room = cur.map((w, i) => (over[i] ? 0 : Math.max(0, w - min[i])));
      const spare = room.reduce((a, b) => a + b, 0);
      // Nothing (or not enough) to take it from: a table in a cell grows, as in Word (its cell's column then makes room
      // for it); one on the page is left as the grid says.
      if (spare < need && !nest) { put(cur); continue; }
      const take = Math.min(spare, need);
      put(cur.map((w0, i) => (over[i] ? min[i] : w0 - (spare ? (room[i] / spare) * take : 0))));
      if (spare < need) cols[0].closest('table').style.width = `${Math.round(cols.reduce((a, c) => a + parseFloat(c.style.width), 0) * 100) / 100}px`;
    }
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
