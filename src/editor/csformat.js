// Formatting given in the editor to complex-script text (runs marked w:rtl or w:cs: Hebrew, Arabic...). Word draws such
// text with the complex-script members of its run properties (w:bCs, w:iCs, w:szCs), not w:b, w:i, w:sz, and its own
// Bold, Italic and Size commands set both. So bold, italic or a size applied to such a run here sets the matching member
// too (kept in the run's raw properties), and removing one removes it: on screen and in Word the run then changes.
import { Plugin } from 'prosemirror-state';
import { AddMarkStep, RemoveMarkStep } from 'prosemirror-transform';
import { schema } from './schema.js';

const CS = { bold: 'bCs', italic: 'iCs', size: 'szCs' };
const off = (x) => /w:val="(0|false|off)"/.test(x);
export const isComplexRun = (items) => items.some(([n, x]) => (n === 'rtl' || n === 'cs') && !off(x));

/** The complex-script member for mark type `name` as the node now has it (null: none). */
function member(node, name) {
  const m = node.marks.find((k) => k.type.name === name);
  if (!m) return null;
  if (name === 'size') return `<w:szCs w:val="${m.attrs.val}"/>`;
  return m.attrs.on === false ? `<w:${CS[name]} w:val="0"/>` : `<w:${CS[name]}/>`;
}

// Hebrew, Arabic, Syriac, Thaana and N'Ko letters (and their presentation forms).
const RTL_TEXT = /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/;

/**
 * Text typed or pasted in a right-to-left script is a right-to-left run (w:rtl), as Word makes it: so Word orders it
 * and its punctuation right to left, and gives it its complex-script formatting.
 */
function markTypedRtl(trs, state, out) {
  const type = schema.marks.rextra;
  trs.forEach((tr, t) => tr.steps.forEach((s, i) => {
    if (!s.slice || !s.slice.size || tr.getMeta('odx-remote')) return;
    let from = s.from, to = s.from + s.slice.size;
    const rest = [tr.mapping.slice(i + 1), ...trs.slice(t + 1).map((x) => x.mapping)];
    for (const m of rest) { from = m.map(from, -1); to = m.map(to, 1); }
    out.doc.nodesBetween(from, Math.min(to, out.doc.content.size), (node, pos) => {
      if (!node.isText || !RTL_TEXT.test(node.text)) return;
      const rx = node.marks.find((k) => k.type === type);
      if (rx && rx.attrs.items.some(([n]) => n === 'rtl')) return;
      const s0 = Math.max(pos, from), e0 = Math.min(pos + node.nodeSize, to);
      out.removeMark(s0, e0, type).addMark(s0, e0, type.create({ items: [...(rx?.attrs.items || []), ['rtl', '<w:rtl/>']] }));
    });
  }));
}

export const csFormatPlugin = new Plugin({
  appendTransaction(trs, _old, state) {
    if (trs.some((tr) => tr.docChanged)) {
      const out = state.tr;
      markTypedRtl(trs, state, out);
      if (out.docChanged) { if (state.storedMarks) out.setStoredMarks(state.storedMarks); return out; }
    }
    const ranges = [];
    trs.forEach((tr, t) => tr.steps.forEach((s, i) => {
      if (!(s instanceof AddMarkStep || s instanceof RemoveMarkStep) || !CS[s.mark.type.name]) return;
      let from = s.from, to = s.to;
      const rest = [tr.mapping.slice(i + 1), ...trs.slice(t + 1).map((x) => x.mapping)];
      for (const m of rest) { from = m.map(from, 1); to = m.map(to, -1); }
      if (to > from) ranges.push([from, to, s.mark.type.name]);
    }));
    if (!ranges.length) return null;
    const out = state.tr;
    const type = schema.marks.rextra;
    for (const [from, to, name] of ranges) {
      out.doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isText) return;
        const rx = node.marks.find((k) => k.type === type);
        if (!rx || !isComplexRun(rx.attrs.items)) return;
        const want = member(node, name);
        const items = rx.attrs.items.filter(([n]) => n !== CS[name]);
        if (want) items.push([CS[name], want]);
        const had = rx.attrs.items.find(([n]) => n === CS[name]);
        if ((had ? had[1] : null) === want) return;
        const s = Math.max(pos, from), e = Math.min(pos + node.nodeSize, to);
        out.removeMark(s, e, type).addMark(s, e, type.create({ items }));
      });
    }
    if (out.docChanged && state.storedMarks) out.setStoredMarks(state.storedMarks);
    return out.docChanged ? out : null;
  },
});
