// Sections: each one's page (size, margins, its headers and footers, its page numbering) and where it ends.
import { schema } from './schema.js';
import { TWIPS_PER_PX, twipsToPx } from '../ooxml/units.js';

/** A section's page from its w:sectPr: size and margins in px, what it leaves out taken from `page`. */
export const marginsIn = (xml, page, gutter) => {
  const num = (a, d) => {
    const m = new RegExp(`<w:pgMar[^>]*\\sw:${a}="(-?\\d+)"`).exec(xml);
    return m ? Math.abs(Number(m[1])) / TWIPS_PER_PX : d;
  };
  const h = /<w:pgSz[^>]*\sw:h="(\d+)"/.exec(xml);
  const w = /<w:pgSz[^>]*\sw:w="(\d+)"/.exec(xml);
  const hasTop = /<w:pgMar[^>]*\sw:top=/.test(xml);
  const hasLeft = /<w:pgMar[^>]*\sw:left=/.test(xml);
  return { top: hasTop ? num('top', 0) + gutter(xml).top : page.top, bottom: num('bottom', page.bottom), height: h ? twipsToPx(h[1]) : page.height, width: w ? twipsToPx(w[1]) : page.width,
    left: hasLeft ? num('left', 0) + gutter(xml).left : page.left, right: num('right', page.right) };
};

/**
 * A section's newspaper columns (w:cols) on `page`, in px, or null for a single column: { num, space, width, rtl, sep,
 * widths, xs }. Columns of their own widths (w:equalWidth="0", each w:col its width and the gap after it) give each
 * one's width and its left edge from the text's (`xs`); `width` is then the widest, and `unequal` is set.
 */
export const columnsIn = (xml, page) => {
  const m = /<w:cols\b[^>]*>/.exec(xml || '');
  const rtl = /<w:bidi(?=[\s/>])(?![^>]*w:val="(?:0|false|off)")/.test(xml || '');
  const sep = m ? /\sw:sep="(?:1|true|on)"/.test(m[0]) : false; // a line between the columns
  if (m && /\sw:equalWidth="(?:0|false|off)"/.test(m[0])) {
    const own = [...(/<w:cols\b[^>]*>([\s\S]*?)<\/w:cols>/.exec(xml)?.[1] || '').matchAll(/<w:col\b[^>]*>/g)]
      .map((c) => ({ w: Number(/w:w="(\d+)"/.exec(c[0])?.[1] || 0) / TWIPS_PER_PX, s: Number(/w:space="(\d+)"/.exec(c[0])?.[1] || 0) / TWIPS_PER_PX }));
    if (own.length >= 2 && own.every((c) => c.w > 0)) {
      const xs = [];
      let x = 0;
      for (const c of own) { xs.push(x); x += c.w + c.s; }
      const widths = own.map((c) => c.w);
      return { num: own.length, space: own[0].s, width: Math.max(...widths), rtl, sep, widths, xs, spaces: own.map((c) => c.s), unequal: true };
    }
  }
  const num = m ? Number((/w:num="(\d+)"/.exec(m[0]) || [])[1] || 1) : 1;
  if (num < 2) return null;
  const space = Number((/w:space="(\d+)"/.exec(m[0]) || [])[1] ?? 720) / TWIPS_PER_PX;
  const width = (page.width - page.left - page.right - space * (num - 1)) / num;
  const widths = Array(num).fill(width);
  return { num, space, width, rtl, sep, widths, xs: widths.map((_, c) => c * (width + space)), spaces: widths.map(() => space) };
};

/**
 * Section breaks: a paragraph carrying a w:sectPr ends a section, and the next section's w:type says whether it starts
 * on a new page (the default) or continues. `page` is the last section's page; `storyHeight(doc)` measures a header or
 * footer (Word moves the body clear of one that reaches past the margin).
 * Returns { ends: block index of each section end -> { page: new page?, next: next section's margins, index }, first:
 * the first section's margins }.
 */
/** The section ends a paragraph carries (its w:sectPr), or null. */
export const sectionEnd = (node) => (node.attrs.extra || []).find((e) => e[0] === 'sectPr')?.[1] ?? null;

/** Every section's w:sectPr XML in order: each one a paragraph ends, then the document's last. */
export function sectionXmls(doc) {
  const out = [];
  doc.forEach((n) => { const x = sectionEnd(n); if (x != null) out.push(x); });
  out.push(doc.attrs.sectPr || '');
  return out;
}

export function sectionBreaks(doc, session, page, storyHeight) {
  const ends = [];
  doc.forEach((node, pos, i) => {
    const s = node.type === schema.nodes.paragraph && (node.attrs.extra || []).find((e) => e[0] === 'sectPr');
    if (s) ends.push({ i, xml: s[1] });
  });
  const out = new Map();
  const xmls = ends.map((e) => e.xml).concat([doc.attrs.sectPr || '']);
  const hfs = session.headerFooters(xmls);
  const same = (a, b) => a.top === b.top && a.bottom === b.bottom && a.left === b.left && a.right === b.right && a.height === b.height && a.width === b.width && a.hf === b.hf && a.numStart === b.numStart && a.numFmt === b.numFmt;
  let prev = null;
  const all = xmls.map((x, n) => {
    const m = n === xmls.length - 1 ? { ...page } : marginsIn(x, page, (y) => session.gutter(y));
    m.hf = hfs[n];
    // Page numbers restart at the section's start (w:pgNumType w:start), else go on from the section before.
    const ns = /<w:pgNumType\b[^>]*w:start="(\d+)"/.exec(x);
    m.numStart = ns ? Number(ns[1]) : null;
    // ... in the section's number format (w:fmt: lowerRoman for a front matter's "Page i").
    m.numFmt = (/<w:pgNumType\b[^>]*w:fmt="(\w+)"/.exec(x) || [])[1] || null;
    // Word moves the body down when the header reaches past the top margin.
    const fit = (type) => ({
      top: Math.max(m.top, m.hf.headerDist + storyHeight(m.hf.header[type])),
      bottom: Math.max(m.bottom, m.hf.footerDist + storyHeight(m.hf.footer[type])),
    });
    if (m.hf.titlePg) m.first = fit('first');
    Object.assign(m, fit('default'));
    return prev && same(m, prev) ? prev : (prev = m);
  });
  ends.forEach((e, n) => {
    const next = n + 1 < ends.length ? ends[n + 1].xml : doc.attrs.sectPr || '';
    out.set(e.i, { page: !/<w:type\b[^>]*w:val="continuous"/.test(next), next: all[n + 1], index: n + 1 });
  });
  return { ends: out, first: all[0] };
}
