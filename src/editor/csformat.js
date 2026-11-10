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

export const csFormatPlugin = new Plugin({
  appendTransaction(trs, _old, state) {
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
    return out.docChanged ? out : null;
  },
});
