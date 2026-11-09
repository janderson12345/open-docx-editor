// Editing a chart's data, title and type (word/charts/chartN.xml). Word draws a chart from the values cached in the
// chart part, and edits them in the workbook embedded beside it: both are rewritten, the cache and the workbook's
// first sheet (and its table, when the data is one), so Word shows the new chart and "Edit Data" opens the same
// numbers. Column, bar, line and area charts can become one another; a pie or doughnut stays one.
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { parseXml, serializeXml, esc, elementKids as els } from './xml.js';
import { ck, pts, textIn, seriesName } from './chart.js';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const AXIS = ['barChart', 'lineChart', 'areaChart'];
const PIE = ['pieChart', 'doughnutChart'];

const cks = (el, name) => els(el).filter((n) => n.localName === name && n.namespaceURI === C);

function typeOf(plot) {
  const el = els(plot).find((n) => n.namespaceURI === C && /Chart$/.test(n.localName));
  if (!el) return { el: null, type: null };
  const name = el.localName;
  if (name === 'barChart') return { el, type: ck(el, 'barDir')?.getAttribute('val') === 'bar' ? 'bar' : 'column' };
  if (name === 'lineChart') return { el, type: 'line' };
  if (name === 'areaChart') return { el, type: 'area' };
  if (name === 'pieChart') return { el, type: 'pie' };
  if (name === 'doughnutChart') return { el, type: 'doughnut' };
  return { el, type: name.replace(/Chart$/, '') }; // 3D, scatter, radar...: shown, not editable
}

/**
 * The chart's data: { type, title, cats, series: [{ name, vals }], editable, types } (types: what it may become).
 * Null when it isn't a chart part.
 */
export function readChart(xml) {
  const doc = parseXml(xml);
  const chart = doc.getElementsByTagNameNS(C, 'chart')[0];
  const plot = chart && ck(chart, 'plotArea');
  if (!plot) return null;
  const { el, type } = typeOf(plot);
  if (!el) return null;
  const sers = cks(el, 'ser');
  const series = sers.map((s, i) => {
    const tx = ck(s, 'tx');
    return { name: tx ? seriesName(tx) : `Series ${i + 1}`, vals: pts(ck(s, 'val')).map((v) => (v === '' ? null : Number(v))) };
  });
  const cats = sers.length ? pts(ck(sers[0], 'cat')) : [];
  const n = Math.max(cats.length, ...series.map((s) => s.vals.length), 0);
  while (cats.length < n) cats.push('');
  for (const s of series) while (s.vals.length < n) s.vals.push(null);
  const titleEl = ck(chart, 'title');
  const title = titleEl ? textIn(ck(titleEl, 'tx')) : '';
  // (One chart in the plot: a combination (bars and a line) is shown but not edited, its other groups' data would go.)
  const groups = els(plot).filter((n) => n.namespaceURI === C && /Chart$/.test(n.localName)).length;
  const editable = groups === 1 && (AXIS.includes(el.localName) || PIE.includes(el.localName));
  const types = PIE.includes(el.localName) ? [type] : editable ? ['column', 'bar', 'line', 'area'] : [type];
  return { type, title, hasTitle: !!titleEl, cats, series, editable, types, sheet: sheetOf(el) };
}

/** The sheet the chart's data comes from (from its first reference), else Sheet1. */
function sheetOf(el) {
  const f = el.getElementsByTagNameNS(C, 'f')[0]?.textContent || '';
  // (A quoted name doubles its own quotes: 'Bob''s data'!$A$1.)
  const m = /^(?:'((?:[^']|'')*)'|([^!]*))!/.exec(f);
  return m ? (m[1] != null ? m[1].replace(/''/g, "'") : m[2]) : 'Sheet1';
}
const colName = (i) => { let s = ''; for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const ref = (sheet, col, r1, r2) => `${/[^A-Za-z0-9_]/.test(sheet) ? `'${sheet.replace(/'/g, "''")}'` : sheet}!$${col}$${r1}${r2 != null ? `:$${col}$${r2}` : ''}`;

/** The chart part's XML with `data` ({ type, title, cats, series }) written in. */
export function writeChart(xml, data) {
  const doc = parseXml(xml);
  const chart = doc.getElementsByTagNameNS(C, 'chart')[0];
  const plot = ck(chart, 'plotArea');
  let { el } = typeOf(plot);
  const sheet = sheetOf(el);
  const make = (name, attrs = {}, kids = []) => {
    const e = doc.createElementNS(C, `c:${name}`);
    for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
    for (const k of kids) e.appendChild(k);
    return e;
  };
  const v = (name, val) => make(name, { val: String(val) });

  // The type: a new element of the kind wanted, its series and axes carried over.
  const want = data.type;
  const current = typeOf(plot).type;
  if (want && want !== current && AXIS.includes(el.localName) && ['column', 'bar', 'line', 'area'].includes(want)) {
    const name = want === 'column' || want === 'bar' ? 'barChart' : `${want}Chart`;
    const next = make(name);
    if (name === 'barChart') next.appendChild(v('barDir', want === 'bar' ? 'bar' : 'col'));
    next.appendChild(v('grouping', name === 'barChart' ? 'clustered' : 'standard'));
    next.appendChild(v('varyColors', 0));
    // (A series keeps its colours only within bars or within lines and areas: a column's "no line" outline would leave
    // a line chart's series as markers alone, as Word drew it. Its theme colour comes back from its index.)
    const family = (k) => (k === 'barChart' ? 'bar' : 'line');
    for (const s of cks(el, 'ser')) {
      if (family(el.localName) !== family(name)) for (const sp of cks(s, 'spPr').concat(cks(s, 'dPt'))) s.removeChild(sp);
      next.appendChild(s);
    }
    if (name === 'barChart') { next.appendChild(v('gapWidth', 219)); next.appendChild(v('overlap', -27)); }
    if (name === 'lineChart') next.appendChild(v('marker', 1));
    for (const a of cks(el, 'axId')) next.appendChild(a);
    plot.replaceChild(next, el);
    el = next;
  } else if (want && want !== current && el.localName === 'barChart' && (want === 'bar' || want === 'column')) {
    ck(el, 'barDir')?.setAttribute('val', want === 'bar' ? 'bar' : 'col');
  }

  // The series: as many as wanted (a new one a copy of the last, in the next theme colour), each its name, categories
  // and values, as cached values and as references to the sheet.
  const n = data.cats.length;
  let sers = cks(el, 'ser');
  const after = (node) => { let x = node.nextSibling; while (x && x.nodeType !== 1) x = x.nextSibling; return x; };
  while (sers.length > data.series.length && sers.length > 1) { el.removeChild(sers.pop()); }
  while (sers.length < data.series.length) {
    const last = sers[sers.length - 1];
    const copy = last.cloneNode(true);
    const sp = ck(copy, 'spPr'); if (sp) copy.removeChild(sp); // (its own colour: the theme's next accent)
    for (const d of cks(copy, 'dPt')) copy.removeChild(d);
    el.insertBefore(copy, after(last));
    sers.push(copy);
  }
  const usedIdx = new Set();
  sers.forEach((s, i) => {
    const kind = el.localName;
    // Series elements in the schema's order for the chart kind: what the kind doesn't allow is dropped.
    if (kind !== 'barChart') for (const x of cks(s, 'invertIfNegative').concat(cks(s, 'shape'))) s.removeChild(x);
    if (kind !== 'lineChart') for (const x of cks(s, 'marker').concat(cks(s, 'smooth'))) s.removeChild(x);
    let idx = Number(ck(s, 'idx')?.getAttribute('val'));
    if (!Number.isFinite(idx) || usedIdx.has(idx)) { idx = 0; while (usedIdx.has(idx)) idx++; }
    usedIdx.add(idx);
    const set = (name, value, before) => {
      let e = ck(s, name);
      if (!e) { e = v(name, value); s.insertBefore(e, before || null); } else e.setAttribute('val', String(value));
      return e;
    };
    const idxEl = set('idx', idx, els(s)[0]);
    set('order', idx, after(idxEl));
    const ser = data.series[i];
    const col = colName(i + 1);
    const cache = (tag, values, numeric) => {
      const c = make(tag, {}, numeric ? [make('formatCode', {}, [doc.createTextNode('General')])] : []);
      c.appendChild(v('ptCount', values.length));
      values.forEach((x, k) => { if (x == null || x === '') return; c.appendChild(make('pt', { idx: String(k) }, [make('v', {}, [doc.createTextNode(String(x))])])); });
      return c;
    };
    const refEl = (tag, f, cacheEl) => make(tag, {}, [make('f', {}, [doc.createTextNode(f)]), cacheEl]);
    const tx = make('tx', {}, [refEl('strRef', ref(sheet, col, 1), cache('strCache', [ser.name], false))]);
    const cat = make('cat', {}, [refEl('strRef', ref(sheet, 'A', 2, n + 1), cache('strCache', data.cats, false))]);
    const val = make('val', {}, [refEl('numRef', ref(sheet, col, 2, n + 1), cache('numCache', ser.vals, true))]);
    const swap = (name, next, anchorNames) => {
      const old = ck(s, name);
      if (old) { s.replaceChild(next, old); return; }
      const anchor = els(s).find((x) => anchorNames.includes(x.localName));
      s.insertBefore(next, anchor || null);
    };
    swap('tx', tx, ['spPr', 'invertIfNegative', 'pictureOptions', 'marker', 'explosion', 'dPt', 'dLbls', 'trendline', 'errBars', 'cat', 'val', 'smooth', 'shape', 'extLst']);
    swap('cat', cat, ['val', 'smooth', 'shape', 'extLst']);
    swap('val', val, ['smooth', 'shape', 'extLst']);
  });

  // The title: rich text of its own, or none.
  let titleEl = ck(chart, 'title');
  const auto = ck(chart, 'autoTitleDeleted');
  if (data.title) {
    const rich = `<c:rich xmlns:c="${C}" xmlns:a="${A}"><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${esc(data.title)}</a:t></a:r></a:p></c:rich>`;
    const richEl = doc.importNode(parseXml(rich).documentElement, true);
    const tx = make('tx', {}, [richEl]);
    if (!titleEl) {
      titleEl = make('title', {}, [tx, v('overlay', 0)]);
      chart.insertBefore(titleEl, els(chart)[0] || null);
    } else {
      const old = ck(titleEl, 'tx');
      if (old) titleEl.replaceChild(tx, old); else titleEl.insertBefore(tx, els(titleEl)[0] || null);
    }
    if (auto) auto.setAttribute('val', '0');
  } else if (titleEl && data.hasTitle === false) {
    chart.removeChild(titleEl);
    if (auto) auto.setAttribute('val', '1'); else chart.insertBefore(v('autoTitleDeleted', 1), els(chart)[0] || null);
  }
  return serializeXml(doc);
}

/** The relationship target of the chart's embedded workbook (from its own .rels), or null. */
export function workbookRel(chartXml, chartRels) {
  const id = /<c:externalData\b[^>]*r:id="([^"]+)"/.exec(chartXml)?.[1];
  return id ? chartRels.find((r) => r.id === id && !r.external) || null : null;
}

/** The embedded workbook with the chart's data on its sheet (and that sheet's table resized to it), or null. */
export function updateWorkbook(bytes, data, sheetName = 'Sheet1') {
  let files;
  try { files = unzipSync(bytes); } catch { return null; }
  const wb = files['xl/workbook.xml'] && strFromU8(files['xl/workbook.xml']);
  if (!wb) return null;
  const sheets = [...wb.matchAll(/<sheet\b[^>]*name="([^"]*)"[^>]*r:id="([^"]+)"/g)];
  const pick = sheets.find((m) => m[1] === sheetName) || sheets[0];
  if (!pick) return null;
  const wbRels = strFromU8(files['xl/_rels/workbook.xml.rels'] || new Uint8Array());
  const target = new RegExp(`Id="${pick[2]}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${pick[2]}"`).exec(wbRels);
  const sheetPath = target ? 'xl/' + (target[1] || target[2]).replace(/^\/?xl\//, '') : null;
  if (!sheetPath || !files[sheetPath]) return null;
  const cols = data.series.length + 1, rows = data.cats.length + 1;
  const last = `${colName(cols - 1)}${rows}`;
  const hasTable = /<tablePart\b/.test(strFromU8(files[sheetPath]));
  const cell = (c, r, value) => {
    const at = `${colName(c)}${r}`;
    if (value == null || value === '') return '';
    if (typeof value === 'number' && Number.isFinite(value)) return `<c r="${at}"><v>${value}</v></c>`;
    return `<c r="${at}" t="inlineStr"><is><t xml:space="preserve">${esc(String(value))}</t></is></c>`;
  };
  const head = [hasTable ? ' ' : '', ...data.series.map((s, i) => s.name || `Series ${i + 1}`)];
  const lines = [`<row r="1">${head.map((h, c) => cell(c, 1, h)).join('')}</row>`];
  data.cats.forEach((cat, k) => lines.push(`<row r="${k + 2}">${cell(0, k + 2, cat)}${data.series.map((s, i) => cell(i + 1, k + 2, s.vals[k])).join('')}</row>`));
  let sheet = strFromU8(files[sheetPath]);
  sheet = sheet.replace(/<sheetData\s*\/>|<sheetData>[\s\S]*?<\/sheetData>/, `<sheetData>${lines.join('')}</sheetData>`);
  sheet = sheet.replace(/<dimension\b[^>]*\/>/, `<dimension ref="A1:${last}"/>`);
  files[sheetPath] = strToU8(sheet);
  // Its table (Word's data sheet is one): the new range, a column for each, named as the header cells.
  if (hasTable) {
    const sheetRels = files[sheetPath.replace(/([^/]+)$/, '_rels/$1.rels')];
    for (const m of strFromU8(sheetRels || new Uint8Array()).matchAll(/Target="([^"]*tables\/[^"]+)"/g)) {
      const tPath = 'xl/' + m[1].replace(/^\.\.\//, '').replace(/^\/?xl\//, '');
      if (!files[tPath]) continue;
      let t = strFromU8(files[tPath]);
      t = t.replace(/(<table\b[^>]*\bref=")[^"]*"/, `$1A1:${last}"`).replace(/(<autoFilter\b[^>]*\bref=")[^"]*"/, `$1A1:${last}"`);
      const columns = head.map((h, i) => `<tableColumn id="${i + 1}" name="${esc(h)}"/>`).join('');
      t = t.replace(/<tableColumns\b[^>]*>[\s\S]*?<\/tableColumns>/, `<tableColumns count="${head.length}">${columns}</tableColumns>`);
      files[tPath] = strToU8(t);
    }
  }
  return zipSync(files, { level: 6 });
}
