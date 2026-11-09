// Editing a chart: double-click it (or the right-click menu's "Edit Chart Data...") for a dialog with its title, its
// type and its data as a grid (categories down, series across), as Word's chart data sheet. Applying it redraws the
// chart and keeps the edited chart XML on the picture's node (one undoable step); saving writes it into the chart
// part and its embedded workbook (ooxml/chartedit.js).
import { Plugin } from 'prosemirror-state';
import { readChart, writeChart } from '../ooxml/chartedit.js';
import { announce } from './live.js';

const LABEL = { column: 'Column', bar: 'Bar', line: 'Line', area: 'Area', pie: 'Pie', doughnut: 'Doughnut' };

/** The chart XML a picture node shows: its own edit, else its part's. Null for a picture that isn't a chart. */
export function chartXmlAt(view, node) {
  if (node?.type.name !== 'image' || !/<c:chart\b/.test(node.attrs.xml || '')) return null;
  return node.attrs.chart || view.odx?.session?.chartPart(node.attrs.xml)?.xml || null;
}

const h = (tag, props = {}, ...kids) => {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (k === 'class') e.className = v;
    else if (k.startsWith('on')) e[k] = v;
    else if (v != null && v !== false) e.setAttribute(k, v === true ? '' : v);
  }
  for (const k of kids) if (k != null) e.append(k);
  return e;
};

/** Opens the chart dialog for the picture at `pos`; true when it opened. */
export function editChart(view, pos) {
  const node = view.state.doc.nodeAt(pos);
  const xml = chartXmlAt(view, node);
  const data = xml && readChart(xml);
  if (!data) return false;
  if (!data.editable) { announce('This kind of chart can\'t be edited here yet'); return false; }
  document.querySelector('.odx-chartdlg-back')?.remove();
  const back = h('div', { class: 'odx-chartdlg-back' });
  const dlg = h('div', { class: 'odx-chartdlg', role: 'dialog', 'aria-modal': 'true', 'aria-labelledby': 'odx-chartdlg-h' });
  const title = h('input', { type: 'text', value: data.title, 'aria-label': 'Chart title', placeholder: 'No title' });
  const type = h('select', { 'aria-label': 'Chart type', disabled: data.types.length < 2 });
  for (const t of data.types) type.append(h('option', { value: t, selected: t === data.type }, LABEL[t] || t));
  const grid = h('table', { class: 'odx-chartdlg-grid' });
  let cats = [...data.cats];
  let series = data.series.map((s) => ({ name: s.name, vals: [...s.vals] }));
  const num = (v) => (v == null || Number.isNaN(v) ? '' : String(v));
  // The grid's inputs write straight into cats/series; drawing it again keeps what was typed.
  const draw = (focus) => {
    grid.replaceChildren();
    const head = h('tr', {}, h('th', { scope: 'col' }, 'Category'));
    series.forEach((s, i) => head.append(h('th', { scope: 'col' }, h('input', { type: 'text', value: s.name, 'aria-label': `Series ${i + 1} name`, oninput: (e) => { s.name = e.target.value; relabel(); } }))));
    grid.append(h('thead', {}, head));
    const body = h('tbody');
    cats.forEach((c, r) => {
      const tr = h('tr', {}, h('th', { scope: 'row' }, h('input', { type: 'text', value: c, 'aria-label': `Category ${r + 1}`, oninput: (e) => { cats[r] = e.target.value; relabel(); } })));
      series.forEach((s, i) => tr.append(h('td', {}, h('input', {
        type: 'text', inputmode: 'decimal', value: num(s.vals[r]), 'data-cell': `${r},${i}`,
        oninput: (e) => { const t = e.target.value.trim().replace(',', '.'); s.vals[r] = t === '' ? null : Number(t); e.target.setAttribute('aria-invalid', t !== '' && Number.isNaN(Number(t)) ? 'true' : 'false'); },
      }))));
      body.append(tr);
    });
    grid.append(body);
    relabel();
    if (focus) grid.querySelector(focus)?.focus();
  };
  // Each value is named by its series and category, as they are typed ("2027, East").
  function relabel() {
    for (const input of grid.querySelectorAll('input[data-cell]')) {
      const [r, i] = input.dataset.cell.split(',').map(Number);
      input.setAttribute('aria-label', `${series[i].name || `Series ${i + 1}`}, ${cats[r] || `category ${r + 1}`}`);
    }
  }
  draw();
  const button = (label, fn, cls = '') => h('button', { type: 'button', class: cls, onclick: fn }, label);
  const close = () => { back.remove(); view.focus(); };
  const apply = () => {
    if (series.some((s) => s.vals.some((v) => v != null && Number.isNaN(v)))) { announce('Some values aren\'t numbers'); dlg.querySelector('[aria-invalid="true"]')?.focus(); return; }
    // The title only when it was changed: one the dialog couldn't read as text (Word's automatic title, a title taken
    // from a cell) opens empty, and applying other changes must leave it as it is, not delete it.
    const typed = title.value.trim();
    const titleChange = typed !== data.title ? { title: typed, hasTitle: !!typed } : { title: null };
    const next = writeChart(xml, { type: type.value, ...titleChange, cats, series });
    const fresh = view.state.doc.nodeAt(pos);
    if (!fresh || fresh.type.name !== 'image') { close(); return; }
    const src = view.odx.session.chartUrlFromXml(next, fresh.attrs.width, fresh.attrs.height) || fresh.attrs.src;
    view.dispatch(view.state.tr.setNodeMarkup(pos, null, { ...fresh.attrs, chart: next, src }));
    close();
    announce('Chart updated');
  };
  const rows = h('div', { class: 'odx-chartdlg-row' },
    button('Add category', () => { cats.push(''); series.forEach((s) => s.vals.push(null)); draw('tbody tr:last-child input'); }),
    button('Remove last category', () => { if (cats.length > 1) { cats.pop(); series.forEach((s) => s.vals.pop()); draw(); } }),
    button('Add series', () => { series.push({ name: `Series ${series.length + 1}`, vals: cats.map(() => null) }); draw('thead th:last-child input'); }),
    button('Remove last series', () => { if (series.length > 1) { series.pop(); draw(); } }));
  dlg.append(
    h('h2', { id: 'odx-chartdlg-h' }, 'Edit chart'),
    h('label', {}, 'Title ', title), h('label', {}, 'Type ', type),
    h('div', { class: 'odx-chartdlg-scroll' }, grid), rows,
    h('div', { class: 'odx-chartdlg-row odx-chartdlg-end' }, button('Apply', apply, 'odx-primary'), button('Cancel', close)));
  // Esc cancels; Tab stays in the dialog; Enter in a field applies.
  dlg.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { e.preventDefault(); close(); }
    else if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); apply(); }
    else if (e.key === 'Tab') {
      const f = [...dlg.querySelectorAll('input, select:not([disabled]), button')];
      const i = f.indexOf(document.activeElement);
      if (e.shiftKey && i <= 0) { e.preventDefault(); f[f.length - 1].focus(); } else if (!e.shiftKey && i === f.length - 1) { e.preventDefault(); f[0].focus(); }
    }
    e.stopPropagation();
  });
  back.onmousedown = (e) => { if (e.target === back) close(); };
  back.append(dlg);
  // (Inside the editor, whose colours (light or dark) it takes.)
  (view.dom.closest('.odx-canvas')?.parentElement || document.body).append(back);
  title.focus();
  return true;
}

/** Double-clicking a chart edits it. */
export function chartPlugin() {
  return new Plugin({
    props: {
      // (The chart's own position, not the click's: on its right half the click falls after it.)
      handleDoubleClickOn(view, pos, node, nodePos) {
        if (view.odx?.readOnly || !chartXmlAt(view, node)) return false;
        return editChart(view, nodePos);
      },
      // And Enter on a selected chart, from the keyboard.
      handleKeyDown(view, e) {
        const node = view.state.selection.node;
        if (e.key !== 'Enter' || e.shiftKey || e.ctrlKey || e.altKey || view.odx?.readOnly || !chartXmlAt(view, node)) return false;
        e.preventDefault();
        // (Taken even when this kind of chart can't be edited, which says so: the keymap's Enter would replace it.)
        editChart(view, view.state.selection.from);
        return true;
      },
    },
  });
}
