// A table beside a float (a picture or shape wrapped square, tight or through, on one side: reader.js anchorCss, a
// browser float, and one placed against the page lifted to its page's first paragraph, pages.js liftFloats). Word sets
// the table in the room the float leaves beside it, moved over to the float's wrap distance if it would overlap it, as
// the browser does; but a table that doesn't fit in that room, its right edge past it (its own indent taking it there, or
// its width), goes below the float, where the browser keeps it beside the float and lets it run past the margin (Word's
// PDFs: a table indented 153pt and 319.5pt wide, 4.5pt past the right margin, under a 71.4pt logo reaching 47.5pt into
// the left margin, starts under the logo; 4.5pt narrower it starts beside it; one with no indent starts beside it at its
// wrap distance). The table is cleared (a node decoration), and the layout measures again (pages.js).
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';

const key = new PluginKey('odx-tableclear');
const decos = (doc, list) => DecorationSet.create(doc, list.flatMap((c) => {
  const node = doc.nodeAt(c.pos);
  return node && node.type.name === 'table' ? [Decoration.node(c.pos, c.pos + node.nodeSize, { style: `clear:${c.side}` }, { key: `tc:${c.side}` })] : [];
}));

export const tableClearPlugin = new Plugin({
  key,
  state: {
    // (passes: how many times the set changed since the document last did, to stop rather than go on for ever.)
    init: () => ({ list: [], set: DecorationSet.empty, passes: 0 }),
    apply(tr, st) {
      const next = tr.getMeta(key);
      if (next) return { list: next, set: decos(tr.doc, next), passes: st.passes + 1 };
      if (!tr.docChanged) return st;
      if (!st.list.length) return st.passes ? { ...st, passes: 0 } : st;
      const list = [];
      for (const c of st.list) {
        const r = tr.mapping.mapResult(c.pos, 1);
        if (!r.deleted) list.push({ pos: r.pos, side: c.side });
      }
      return { list, set: decos(tr.doc, list), passes: 0 };
    },
  },
  props: { decorations: (state) => key.getState(state).set },
});

const MAX_PASSES = 12;
const SLACK = 0.5; // px

/** Clear the tables that don't fit beside the floats they start by, and only those. True when that changed. */
export function applyTableClear(view) {
  const st = key.getState(view.state);
  if (!st || st.passes > MAX_PASSES) return false;
  const dom = view.dom;
  const floats = [];
  for (const img of dom.querySelectorAll('img[style*="float"]')) {
    const cs = getComputedStyle(img);
    if (cs.float !== 'left' && cs.float !== 'right') continue;
    const r = img.getBoundingClientRect();
    if (!r.width || !r.height) continue;
    const m = (s) => parseFloat(cs[s]) || 0;
    floats.push({ img, side: cs.float, left: r.left - m('marginLeft'), right: r.right + m('marginRight'), top: r.top - Math.max(0, m('marginTop')), bottom: r.bottom + m('marginBottom') });
  }
  if (!floats.length && !st.list.length) return false;
  const was = new Map(st.list.map((c) => [c.pos, c.side]));
  const next = [];
  if (floats.length) {
    const dr = dom.getBoundingClientRect(), dcs = getComputedStyle(dom);
    const colL = dr.left + (parseFloat(dcs.paddingLeft) || 0), colR = dr.right - (parseFloat(dcs.paddingRight) || 0);
    // (Only those reaching into the column: one in the margin leaves the table its whole width.)
    const inCol = floats.filter((f) => (f.side === 'left' ? f.right > colL + SLACK : f.left < colR - SLACK));
    view.state.doc.forEach((node, pos) => {
      if (node.type.name !== 'table' || !inCol.length) return;
      const el = view.nodeDOM(pos);
      if (!el || el.nodeType !== 1 || el.parentElement !== dom) return;
      const ecs = getComputedStyle(el);
      if (ecs.float !== 'none' || ecs.position === 'absolute') return; // (a floating table: placed on its own)
      const table = el.tagName === 'TABLE' ? el : el.querySelector(':scope > table');
      if (!table) return;
      const tr = table.getBoundingClientRect();
      // (Where it starts in the flow: cleared, under the block before it, whose place the clearing doesn't move.)
      const prev = el.previousElementSibling;
      const top0 = was.has(pos) && prev ? prev.getBoundingClientRect().bottom + Math.max(0, parseFloat(getComputedStyle(prev).marginBottom) || 0) : tr.top;
      const row = table.rows[0]?.getBoundingClientRect();
      const bottom0 = top0 + (row ? row.height : tr.height);
      const near = inCol.filter((f) => !el.contains(f.img) && f.top < bottom0 - SLACK && f.bottom > top0 + SLACK);
      if (!near.length) return;
      let bandL = colL, bandR = colR;
      for (const f of near) {
        if (f.side === 'left') bandL = Math.max(bandL, f.right);
        else bandR = Math.min(bandR, f.left);
      }
      if (Math.max(tr.left, bandL) + tr.width <= bandR + SLACK) return;
      const sides = new Set(near.map((f) => f.side));
      next.push({ pos, side: sides.size > 1 ? 'both' : [...sides][0] });
    });
  }
  const sig = (l) => l.map((c) => `${c.pos}:${c.side}`).join(',');
  if (sig(next) === sig(st.list)) return false;
  view.dispatch(view.state.tr.setMeta(key, next).setMeta('addToHistory', false));
  return true;
}
