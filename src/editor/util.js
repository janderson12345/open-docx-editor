// Small helpers the editor's modules share.

/**
 * Now, as Word writes a revision's or comment's date (w:date): ISO 8601 to the second, the local wall-clock time
 * marked "Z" (Word writes the time on the author's clock; the true UTC time, when kept, goes in w16du:dateUtc).
 */
export const now = () => {
  const d = new Date();
  return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().replace(/\.\d+Z$/, 'Z');
};

/**
 * A revision's or comment's time to show: its UTC time (w16du:dateUtc) when it has one, else its w:date read as the
 * wall-clock time it is, as Word shows it. Null when there is none.
 */
export function wordDate(date, utc) {
  const d = utc && !Number.isNaN(Date.parse(utc)) ? new Date(utc) : date ? new Date(String(date).replace(/Z$/i, '')) : null;
  return d && !Number.isNaN(d.getTime()) ? d : null;
}

/** The w:id in an atom node's raw XML (a comment or note reference), or undefined. */
export const xmlId = (node) => (/w:id="([^"]*)"/.exec(node.attrs.xml) || [])[1];

/**
 * A paragraph's line height as getComputedStyle's lineHeight gives it ('15.33px' or 'normal'): document paragraphs keep
 * theirs in --odx-lhv, their own line-height being 0 (textspans.js). `cs` is the paragraph's computed style.
 */
export function lineHeightOf(cs) {
  const v = cs.getPropertyValue('--odx-lhv').trim();
  if (!v) return cs.lineHeight;
  if (v === 'normal' || v.endsWith('px')) return v;
  const n = Number(v);
  return Number.isFinite(n) ? `${n * parseFloat(cs.fontSize)}px` : cs.lineHeight;
}

/**
 * An element's box as laid out, before its CSS transform (a turned shape's picture, turned about its centre and moved
 * back where its margins put it, reader.js turned): what takes room on its line and page, where its bounding rect is
 * the turned outline's. (At 100%, as the layout measures.)
 */
export function layoutRect(el) {
  const r = el.getBoundingClientRect();
  const t = getComputedStyle(el).transform;
  if (!t || t === 'none') return r;
  const m = new DOMMatrix(t);
  // (Not turned, only moved or flipped: its own rect, moved back. Otherwise its exact size: offsetWidth is rounded.)
  if (Math.abs(m.b) < 1e-9 && Math.abs(m.c) < 1e-9 && Math.abs(Math.abs(m.a) - 1) < 1e-9 && Math.abs(Math.abs(m.d) - 1) < 1e-9) {
    return { left: r.left - m.e, right: r.right - m.e, top: r.top - m.f, bottom: r.bottom - m.f, width: r.width, height: r.height };
  }
  const cs = getComputedStyle(el);
  const cx = (r.left + r.right) / 2 - m.e, cy = (r.top + r.bottom) / 2 - m.f;
  const w = parseFloat(cs.width) || el.offsetWidth, h = parseFloat(cs.height) || el.offsetHeight;
  return { left: cx - w / 2, right: cx + w / 2, top: cy - h / 2, bottom: cy + h / 2, width: w, height: h };
}
