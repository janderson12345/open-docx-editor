// DrawingML charts (word/charts/chartN.xml) drawn as SVG from the values Word
// caches in the chart part. Covers the common kinds: column/bar, line, area,
// scatter, pie and doughnut. The chart part itself is never rewritten.
import { parseXml, esc } from './xml.js';
import { OFFICE } from './shapes.js';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const TYPES = ['barChart', 'bar3DChart', 'lineChart', 'line3DChart', 'areaChart', 'area3DChart', 'scatterChart', 'pieChart', 'pie3DChart', 'doughnutChart'];

export const ck = (el, name) => (el ? [...el.childNodes].find((n) => n.nodeType === 1 && n.localName === name && n.namespaceURI === C) || null : null);
const val = (el, name) => ck(el, name)?.getAttribute('val') ?? null;
// A cache's points at their indexes ('' where one is left out: a blank cell), as many as its ptCount.
export const pts = (el) => {
  if (!el) return [];
  const all = [...el.getElementsByTagNameNS(C, 'pt')];
  const count = Number(el.getElementsByTagNameNS(C, 'ptCount')[0]?.getAttribute('val'));
  const out = Array.from({ length: Number.isFinite(count) ? count : all.length }, () => '');
  for (const p of all) { const i = Number(p.getAttribute('idx')); if (i >= 0) out[i] = p.getElementsByTagNameNS(C, 'v')[0]?.textContent ?? ''; }
  return out;
};
export const textIn = (el) => (el ? [...el.getElementsByTagNameNS(A, 't')].map((t) => t.textContent).join('') : '');
/** A series' name (c:tx): cached from a cell, rich text, or a literal value (c:v). */
export const seriesName = (tx) => (tx ? pts(tx)[0] || textIn(tx) || tx.textContent.trim() : '');

function fillOf(spPr, colors) {
  const fill = spPr && [...spPr.childNodes].find((n) => n.localName === 'solidFill');
  const ln = spPr && [...spPr.childNodes].find((n) => n.localName === 'ln');
  const solid = fill || (ln && [...ln.childNodes].find((n) => n.localName === 'solidFill'));
  const c = solid && [...solid.childNodes].find((n) => n.nodeType === 1);
  if (!c) return null;
  if (c.localName === 'srgbClr') return c.getAttribute('val');
  if (c.localName === 'schemeClr') {
    const key = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' }[c.getAttribute('val')] || c.getAttribute('val');
    return colors[key] || OFFICE[key] || null;
  }
  return null;
}

/** Round axis steps to 1, 2 or 5 times a power of ten, like Excel. */
function niceAxis(lo, hi) {
  if (lo === hi) hi = lo + 1;
  // Excel leaves headroom: a maximum within 5% of a tick gets the next tick.
  const room = (hi - lo) * 0.05;
  if (hi > 0) hi += room;
  if (lo < 0) lo -= room;
  const raw = (hi - lo) / 6;
  const p = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 5, 10].map((m) => m * p).find((s) => s >= raw);
  return { min: Math.floor(lo / step) * step, max: Math.ceil(hi / step - 1e-9) * step, step };
}

const fmt = (v) => String(+v.toFixed(10));

/**
 * @param {string} xml chart part
 * @param {{ width: number, height: number, font?: string, colors?: object }} opts sizes in px
 * @returns {string|null} SVG markup
 */
export function chartSvg(xml, { width, height, font = 'Calibri', colors = {} }) {
  const doc = parseXml(xml);
  const chart = doc.getElementsByTagNameNS(C, 'chart')[0];
  const plot = ck(chart, 'plotArea');
  const typeEl = plot && [...plot.childNodes].find((n) => n.nodeType === 1 && TYPES.includes(n.localName));
  if (!typeEl || !width || !height) return null;
  const kind = typeEl.localName.replace('3D', '').replace('Chart', '');
  const accent = (i) => colors['accent' + ((i % 6) + 1)] || OFFICE['accent' + ((i % 6) + 1)];
  const series = [...typeEl.childNodes].filter((n) => n.localName === 'ser').map((s, i) => {
    const idx = Number(val(s, 'idx') ?? i);
    const nameEl = ck(s, 'tx');
    return {
      name: nameEl ? seriesName(nameEl) : `Series ${idx + 1}`,
      cats: pts(ck(s, 'cat') || ck(s, 'xVal')),
      vals: pts(ck(s, 'val') || ck(s, 'yVal')).map((v) => (v === '' ? null : Number(v))), // (a blank cell: no point)
      color: fillOf(ck(s, 'spPr'), colors) || accent(idx),
      points: [...s.childNodes].filter((n) => n.localName === 'dPt').map((d) => [Number(val(d, 'idx')), fillOf(ck(d, 'spPr'), colors)]),
    };
  });
  if (!series.length) return null;

  const gray = '595959';
  const out = [];
  const text = (x, y, s, { size = 9, anchor = 'middle', color = gray } = {}) =>
    out.push(`<text x="${x.toFixed(1)}" y="${y.toFixed(1)}" font-size="${size}pt" fill="#${color}" text-anchor="${anchor}">${esc(s)}</text>`);
  const pt = 4 / 3; // px per pt

  // Title: explicit text, else Word's placeholder when a title element exists.
  let top = 8;
  const titleEl = ck(chart, 'title');
  const single = series.length === 1;
  const title = titleEl ? textIn(ck(titleEl, 'tx')) || (single ? series[0].name : 'Chart Title') : val(chart, 'autoTitleDeleted') === '0' && single ? series[0].name : '';
  if (title) {
    text(width / 2, top + 14 * pt, title, { size: 14 });
    top += 14 * pt + 10;
  }

  // Legend
  const legendEl = ck(chart, 'legend');
  const pos = legendEl ? val(legendEl, 'legendPos') || 'r' : null;
  const pie = kind === 'pie' || kind === 'doughnut';
  const entries = pie ? series[0].cats.map((c, i) => [c, series[0].points.find((p) => p[0] === i)?.[1] || accent(i)]) : series.map((s) => [s.name, s.color]);
  let bottom = height - 8;
  let right = width - 16;
  const sw = 7;
  if (pos === 'b' || pos === 't') {
    const widths = entries.map(([n]) => sw + 4 + n.length * 5.2 + 14);
    let x = (width - widths.reduce((a, b) => a + b, 0)) / 2;
    const y = pos === 'b' ? height - 12 : top + 4;
    entries.forEach(([n, c], i) => {
      out.push(`<rect x="${x.toFixed(1)}" y="${(y - sw).toFixed(1)}" width="${sw}" height="${sw}" fill="#${c}"/>`);
      text(x + sw + 4, y, n, { anchor: 'start' });
      x += widths[i];
    });
    if (pos === 'b') bottom -= 22;
    else top += 22;
  } else if (pos) {
    const w = Math.max(...entries.map(([n]) => n.length * 5.2)) + sw + 14;
    // Entries 24px apart, as in Word's PDF (file-sample_1MB's "Column 1-3"), centred on the plot.
    let y = (top + bottom) / 2 - (entries.length * 24) / 2 + 14;
    for (const [n, c] of entries) {
      out.push(`<rect x="${(width - w).toFixed(1)}" y="${(y - sw).toFixed(1)}" width="${sw}" height="${sw}" fill="#${c}"/>`);
      text(width - w + sw + 4, y, n, { anchor: 'start' });
      y += 24;
    }
    right -= w + 6;
  }

  if (pie) {
    const s = series[0];
    const total = s.vals.reduce((a, v) => a + Math.max(v, 0), 0) || 1;
    const cx = (10 + right) / 2;
    const cy = (top + bottom) / 2;
    const r = Math.max(4, Math.min(right - 10, bottom - top) / 2 - 4);
    const hole = kind === 'doughnut' ? r * (Number(val(typeEl, 'holeSize') || 50) / 100) : 0;
    let a0 = ((Number(val(typeEl, 'firstSliceAng') || 0) - 90) * Math.PI) / 180;
    s.vals.forEach((v, i) => {
      const a1 = a0 + (Math.max(v, 0) / total) * 2 * Math.PI;
      const big = a1 - a0 > Math.PI ? 1 : 0;
      const p = (a, rr) => `${(cx + rr * Math.cos(a)).toFixed(2)} ${(cy + rr * Math.sin(a)).toFixed(2)}`;
      const d = hole
        ? `M${p(a0, r)}A${r} ${r} 0 ${big} 1 ${p(a1, r)}L${p(a1, hole)}A${hole} ${hole} 0 ${big} 0 ${p(a0, hole)}Z`
        : `M${cx} ${cy}L${p(a0, r)}A${r} ${r} 0 ${big} 1 ${p(a1, r)}Z`;
      out.push(`<path d="${d}" fill="#${s.points.find((q) => q[0] === i)?.[1] || accent(i)}" stroke="#fff" stroke-width="1"/>`);
      a0 = a1;
    });
  } else {
    const horizontal = kind === 'bar' && val(typeEl, 'barDir') === 'bar';
    const grouping = val(typeEl, 'grouping') || 'clustered';
    const stacked = grouping === 'stacked' || grouping === 'percentStacked';
    const percent = grouping === 'percentStacked';
    const cats = series.find((s) => s.cats.length)?.cats || series[0].vals.map((_, i) => String(i + 1));
    const n = Math.max(cats.length, ...series.map((s) => s.vals.length));
    const totals = Array.from({ length: n }, (_, i) => series.reduce((a, s) => a + Math.abs(s.vals[i] || 0), 0) || 1);
    const value = (s, i) => (percent ? ((s.vals[i] || 0) / totals[i]) * 100 : s.vals[i] || 0);
    let lo = 0;
    let hi = 0;
    for (let i = 0; i < n; i++) {
      let pos = 0;
      let neg = 0;
      for (const s of series) {
        const v = value(s, i);
        if (stacked) (v >= 0 ? (pos += v) : (neg += v));
        else (hi = Math.max(hi, v), lo = Math.min(lo, v));
      }
      if (stacked) (hi = Math.max(hi, pos), lo = Math.min(lo, neg));
    }
    const ax = niceAxis(lo, hi);
    const ticks = [];
    for (let v = ax.min; v <= ax.max + ax.step / 2; v += ax.step) ticks.push(v);
    const labelW = Math.max(...ticks.map((t) => fmt(t).length)) * 5.5 + (percent ? 6 : 0);
    const x0 = horizontal ? 10 + Math.max(...cats.map((c) => c.length)) * 5.2 + 6 : 10 + labelW + 6;
    const y1 = bottom - (horizontal ? 18 : 20);
    const x1 = right;
    const y0 = top + 12;
    const span = (v) => (v - ax.min) / (ax.max - ax.min);
    const vpos = (v) => (horizontal ? x0 + span(v) * (x1 - x0) : y1 - span(v) * (y1 - y0));
    for (const t of ticks) {
      const p = vpos(t);
      const label = fmt(t) + (percent ? '%' : '');
      if (horizontal) {
        out.push(`<line x1="${p.toFixed(1)}" y1="${y0}" x2="${p.toFixed(1)}" y2="${y1.toFixed(1)}" stroke="#D9D9D9" stroke-width="0.75"/>`);
        text(p, y1 + 14, label);
      } else {
        out.push(`<line x1="${x0.toFixed(1)}" y1="${p.toFixed(1)}" x2="${x1.toFixed(1)}" y2="${p.toFixed(1)}" stroke="#D9D9D9" stroke-width="0.75"/>`);
        text(x0 - 5, p + 3.5, label, { anchor: 'end' });
      }
    }
    const band = ((horizontal ? y1 - y0 : x1 - x0) / n);
    cats.forEach((c, i) => {
      const mid = (horizontal ? y1 - (i + 0.5) * band : x0 + (i + 0.5) * band);
      if (horizontal) text(x0 - 5, mid + 3.5, c, { anchor: 'end' });
      else text(mid, y1 + 14, c);
    });
    const base = vpos(Math.max(ax.min, Math.min(0, ax.max)));
    if (kind === 'bar') {
      const gap = Number(val(typeEl, 'gapWidth') ?? 150) / 100;
      const overlap = stacked ? 1 : Number(val(typeEl, 'overlap') ?? 0) / 100;
      const k = stacked ? 1 : series.length;
      const bw = band / (k - (k - 1) * overlap + gap);
      const stackPos = new Array(n).fill(0);
      const stackNeg = new Array(n).fill(0);
      series.forEach((s, si) => {
        for (let i = 0; i < n; i++) {
          const v = value(s, i);
          let from = 0;
          if (stacked) {
            from = v >= 0 ? stackPos[i] : stackNeg[i];
            if (v >= 0) stackPos[i] += v;
            else stackNeg[i] += v;
          }
          const a = vpos(from);
          const b = vpos(from + v);
          const off = (gap / 2) * bw + (stacked ? 0 : si * bw * (1 - overlap));
          const col = s.points.find((q) => q[0] === i)?.[1] || s.color;
          if (horizontal) {
            const y = y1 - (i + 1) * band + off;
            out.push(`<rect x="${Math.min(a, b).toFixed(1)}" y="${y.toFixed(1)}" width="${Math.abs(b - a).toFixed(1)}" height="${bw.toFixed(1)}" fill="#${col}"/>`);
          } else {
            const x = x0 + i * band + off;
            out.push(`<rect x="${x.toFixed(1)}" y="${Math.min(a, b).toFixed(1)}" width="${bw.toFixed(1)}" height="${Math.abs(b - a).toFixed(1)}" fill="#${col}"/>`);
          }
        }
      });
    } else {
      const acc = new Array(n).fill(0);
      const xs = (i) => x0 + (i + 0.5) * band;
      const layers = series.map((s) => {
        const ys = [];
        for (let i = 0; i < n; i++) {
          const v = value(s, i) + (stacked ? acc[i] : 0);
          ys.push(vpos(v));
          if (stacked) acc[i] = v;
        }
        return ys;
      });
      series.forEach((s, si) => {
        const line = layers[si].map((y, i) => `${xs(i).toFixed(1)} ${y.toFixed(1)}`).join('L');
        // (A line leaves a gap at a blank cell, as Word draws one by default: dispBlanksAs gap.)
        const gaps = kind === 'line' && !stacked && s.vals.some((v) => v == null);
        const broken = gaps ? layers[si].map((y, i) => (s.vals[i] == null ? null : `${xs(i).toFixed(1)} ${y.toFixed(1)}`)).reduce((d, p, i, all) => (p == null ? d : d + (i && all[i - 1] != null ? 'L' : 'M') + p), '') : null;
        if (kind === 'area') {
          const under = stacked && si ? layers[si - 1].map((y, i) => `${xs(i).toFixed(1)} ${y.toFixed(1)}`).reverse().join('L') : `${xs(n - 1).toFixed(1)} ${base.toFixed(1)}L${xs(0).toFixed(1)} ${base.toFixed(1)}`;
          out.push(`<path d="M${line}L${under}Z" fill="#${s.color}"/>`);
        } else {
          out.push(`<path d="${broken ?? `M${line}`}" fill="none" stroke="#${s.color}" stroke-width="2.25" stroke-linejoin="round" stroke-linecap="round"/>`);
          if (kind === 'scatter') layers[si].forEach((y, i) => out.push(`<circle cx="${xs(i).toFixed(1)}" cy="${y.toFixed(1)}" r="3" fill="#${s.color}"/>`));
        }
      });
    }
    out.push(horizontal
      ? `<line x1="${base.toFixed(1)}" y1="${y0}" x2="${base.toFixed(1)}" y2="${y1.toFixed(1)}" stroke="#BFBFBF" stroke-width="0.75"/>`
      : `<line x1="${x0.toFixed(1)}" y1="${base.toFixed(1)}" x2="${x1.toFixed(1)}" y2="${base.toFixed(1)}" stroke="#BFBFBF" stroke-width="0.75"/>`);
  }

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" font-family="${esc(font)}, Carlito, sans-serif"><rect width="100%" height="100%" fill="#fff" stroke="#D9D9D9" stroke-width="0.75"/>${out.join('')}</svg>`;
}
