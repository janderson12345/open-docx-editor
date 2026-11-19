// Sections: each one's page (size, margins, its headers and footers, its page numbering) and where it ends.
import { Plugin } from 'prosemirror-state';
import { ReplaceStep } from 'prosemirror-transform';
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
 * widths, xs, fnCols (w15:footnoteColumns, when set) }. Columns of their own widths (w:equalWidth="0", each w:col its width and the gap after it) give each
 * one's width and its left edge from the text's (`xs`); `width` is then the widest, and `unequal` is set.
 */
export const columnsIn = (xml, page) => {
  const m = /<w:cols\b[^>]*>/.exec(xml || '');
  const rtl = /<w:bidi(?=[\s/>])(?![^>]*w:val="(?:0|false|off)")/.test(xml || '');
  const sep = m ? /\sw:sep="(?:1|true|on)"/.test(m[0]) : false; // a line between the columns
  // (w15:footnoteColumns: the footnotes' own columns at the page's foot, when it says.)
  const fn = /<w15:footnoteColumns\b[^>]*w:val="(\d+)"/.exec(xml || '');
  // (Their gap: w:cols' own w:space, as Word spaces them, also when the section's columns have widths of their own.)
  const fnCols = fn ? { fnCols: Math.min(45, Number(fn[1])), fnGap: Number((m && /\sw:space="(\d+)"/.exec(m[0])?.[1]) ?? 720) / TWIPS_PER_PX } : {};
  if (m && /\sw:equalWidth="(?:0|false|off)"/.test(m[0])) {
    const own = [...(/<w:cols\b[^>]*>([\s\S]*?)<\/w:cols>/.exec(xml)?.[1] || '').matchAll(/<w:col\b[^>]*>/g)].slice(0, 45)
      .map((c) => ({ w: Number(/w:w="(\d+)"/.exec(c[0])?.[1] || 0) / TWIPS_PER_PX, s: Number(/w:space="(\d+)"/.exec(c[0])?.[1] || 0) / TWIPS_PER_PX }));
    if (own.length >= 2 && own.every((c) => c.w > 0)) {
      const xs = [];
      let x = 0;
      for (const c of own) { xs.push(x); x += c.w + c.s; }
      const widths = own.map((c) => c.w);
      return { num: own.length, space: own[0].s, width: Math.max(...widths), rtl, sep, widths, xs, spaces: own.map((c) => c.s), unequal: true, ...fnCols };
    }
  }
  // (At most 45 columns, Word's limit: a file's huge w:num allocated that many columns on every layout pass.)
  const num = Math.min(45, m ? Number((/w:num="(\d+)"/.exec(m[0]) || [])[1] || 1) : 1);
  if (num < 2) return null;
  const space = Number((/w:space="(\d+)"/.exec(m[0]) || [])[1] ?? 720) / TWIPS_PER_PX;
  const width = (page.width - page.left - page.right - space * (num - 1)) / num;
  const widths = Array(num).fill(width);
  return { num, space, width, rtl, sep, widths, xs: widths.map((_, c) => c * (width + space)), spaces: widths.map(() => space), ...fnCols };
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

const BODY_LINE = 20; // px of body text a page keeps under the tallest header or footer (a line of 11pt text)
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
    // (But the page keeps a line of body text however tall they are, as Word's does: a footer longer than the page leaves
    // one line at its top, each page of the body one line.)
    // (A page with no header (or footer) of its kind has none to move the body: its margin, whatever its header distance.
    // A first-page header alone, a 1in header distance over a 0.26in top margin: pages 2 on had their text from 1in down,
    // where Word's start at the margin.)
    const fit = (type) => {
      const H = m.height || page.height;
      const hs = m.hf.header[type], fs = m.hf.footer[type];
      const top = hs ? Math.min(Math.max(m.top, m.hf.headerDist + storyHeight(hs)), Math.max(m.top, H - m.bottom - BODY_LINE)) : m.top;
      const bottom = fs ? Math.min(Math.max(m.bottom, m.hf.footerDist + storyHeight(fs)), Math.max(m.bottom, H - top - BODY_LINE)) : m.bottom;
      return { top, bottom };
    };
    if (m.hf.titlePg) m.first = fit('first');
    Object.assign(m, fit('default'));
    return prev && same(m, prev) ? prev : (prev = m);
  });
  ends.forEach((e, n) => {
    const next = n + 1 < ends.length ? ends[n + 1].xml : doc.attrs.sectPr || '';
    // (An odd- or even-page break starts the section on the next page of that parity, a blank page between if need be.)
    const parity = (/<w:type\b[^>]*w:val="(oddPage|evenPage)"/.exec(next) || [])[1] || null;
    out.set(e.i, { page: !/<w:type\b[^>]*w:val="continuous"/.test(next), parity, next: all[n + 1], index: n + 1 });
  });
  return { ends: out, first: all[0] };
}

/**
 * A section break deleted (Delete or Backspace at it, or a selection across it): the paragraph mark that carried the
 * w:sectPr is gone, and with it that section's setup, as in Word. The text before the break joins the section after it,
 * the one whose mark is left: the joined paragraph ends with the following paragraph's w:sectPr (or none), not the one
 * ProseMirror's join keeps from the first paragraph. (Word, columns on selected text, the break before or after them
 * deleted: the text joins the following section, in its columns. Not a tracked deletion, Undo or a co-editor's step.)
 */
export const sectionJoinPlugin = new Plugin({
  appendTransaction(trs, old, state) {
    const fixes = []; // [from, to, sectPr]: the top-level paragraphs in [from, to] of state.doc, the last ending with sectPr
    const later = (tr, t, i, p, assoc) => {
      for (const m of [tr.mapping.slice(i + 1), ...trs.slice(t + 1).map((x) => x.mapping)]) p = m.map(p, assoc);
      return p;
    };
    trs.forEach((tr, t) => {
      if (!tr.docChanged || tr.getMeta('odx-remote') || tr.getMeta('history$') || tr.getMeta('y-sync$')) return;
      tr.steps.forEach((step, i) => {
        if (!(step instanceof ReplaceStep) || step.from >= step.to) return;
        const doc = tr.docs[i];
        const $a = doc.resolve(step.from);
        const $b = doc.resolve(step.to);
        // Delete at the end of a paragraph before an empty one that ends a section (Word's own section breaks are often
        // on an empty paragraph): ProseMirror takes out the empty paragraph, Word the mark before it, so the break stays,
        // now at the end of the paragraph the caret was in.
        if ($a.depth === 0 && $b.depth === 0 && $b.index(0) === $a.index(0) + 1 && $a.index(0) > 0 && old.selection.empty && t === 0 && i === 0) {
          const gone = doc.child($a.index(0));
          const caret = old.selection.$from;
          if (gone.type === schema.nodes.paragraph && !gone.content.size && sectionEnd(gone) != null && caret.depth === 1
            && caret.index(0) === $a.index(0) - 1 && caret.pos === caret.end() && doc.child($a.index(0) - 1).type === schema.nodes.paragraph) {
            const at = later(tr, t, i, step.from, -1) - 1;
            fixes.push([at, at, sectionEnd(gone)]);
          }
          return;
        }
        if ($a.depth < 1 || $b.depth < 1) return;
        const ia = $a.index(0);
        const ib = $b.index(0);
        if (ia >= ib || $a.node(1).type !== schema.nodes.paragraph || $b.node(1).type !== schema.nodes.paragraph) return;
        let any = false;
        for (let k = ia; k <= ib; k++) if (sectionEnd(doc.child(k)) != null) any = true;
        if (!any) return;
        fixes.push([later(tr, t, i, step.from, -1), later(tr, t, i, step.getMap().map(step.to, 1), 1), sectionEnd($b.node(1))]);
      });
    });
    if (!fixes.length) return null;
    const out = state.tr;
    for (const [a, b, sect] of fixes) {
      if (a > b || b > out.doc.content.size) continue;
      const $a = out.doc.resolve(a);
      const $b = out.doc.resolve(b);
      if ($a.depth < 1 || $b.depth < 1) continue;
      for (let k = $a.index(0); k <= $b.index(0); k++) {
        let pos = 0;
        for (let j = 0; j < k; j++) pos += out.doc.child(j).nodeSize;
        const node = out.doc.child(k);
        if (node.type !== schema.nodes.paragraph) continue;
        const want = k === $b.index(0) ? sect : null;
        if (sectionEnd(node) === want) continue;
        const extra = (node.attrs.extra || []).filter((e) => e[0] !== 'sectPr');
        out.setNodeMarkup(pos, null, { ...node.attrs, extra: want == null ? extra : [...extra, ['sectPr', want]] });
      }
    }
    return out.docChanged ? out : null;
  },
});
