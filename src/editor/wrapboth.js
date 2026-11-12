// Text on both sides of a floating shape (w:wrapSquare, Tight or Through, wrapText="bothSides"; the one-sided kinds are
// floats, reader.js anchorCss): Word splits each line beside the shape, filling the part
// from the margin to the shape's left wrap distance with as many whole words as fit, then going on from its right one
// (its PDFs). A browser float keeps text off one side only, so such a shape is set where it is, out of the flow
// (reader.js anchorCss, --odx-both: the room kept clear round it), and each line that runs into it gets a spacer at the
// first word that doesn't fit before it, as wide as the jump to its right side. Lines are measured as drawn; a pass
// adds one spacer to each paragraph that needs one, or takes away those that no longer fit, and the layout measures
// again (pages.js), so a paragraph is settled line by line, top down. The spacers are view decorations, kept through
// edits (mapped) and checked again on each pass.
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';

const key = new PluginKey('odx-wrapboth');
const spacer = (w) => () => {
  const el = document.createElement('span');
  el.className = 'odx-wb';
  el.contentEditable = 'false';
  el.style.width = `${w}px`;
  return el;
};
const decos = (doc, list) => DecorationSet.create(doc, list.map((s) => Decoration.widget(s.pos, spacer(s.w), { side: -1, key: `wb:${s.pos}:${s.w.toFixed(2)}`, ignoreSelection: true })));

export const wrapBothPlugin = new Plugin({
  key,
  state: {
    // (passes: how many times the spacers changed since the document last did, to stop rather than go on for ever.)
    init: () => ({ list: [], set: DecorationSet.empty, passes: 0 }),
    apply(tr, st) {
      const next = tr.getMeta(key);
      if (next) return { list: next, set: decos(tr.doc, next), passes: st.passes + 1 };
      if (!tr.docChanged) return st;
      if (!st.list.length) return st.passes ? { ...st, passes: 0 } : st;
      // (Kept where their words went; one whose word was deleted goes, and the next pass puts what's needed back.)
      const list = [];
      for (const s of st.list) {
        const r = tr.mapping.mapResult(s.pos, 1);
        if (!r.deleted) list.push({ pos: r.pos, w: s.w });
      }
      return { list, set: decos(tr.doc, list), passes: 0 };
    },
  },
  props: { decorations: (state) => key.getState(state).set },
});

// The room kept clear round each such shape, as drawn (its turned outline when turned): { left, right, top, bottom }.
function zones(view) {
  const out = [];
  for (const img of view.dom.querySelectorAll('img[style*="--odx-both"]')) {
    const v = getComputedStyle(img).getPropertyValue('--odx-both').trim().split(/\s+/).map((x) => parseFloat(x) || 0);
    const r = img.getBoundingClientRect();
    if (!r.width || !r.height) continue;
    const [t, rt, b, l] = v.length === 4 ? v : [0, 0, 0, 0];
    out.push({ left: r.left - l, right: r.right + rt, top: r.top - t, bottom: r.bottom + b });
  }
  // A float drawn up to end at its page's edge (pages.js floatsPast) over lines of paragraphs above its own: Word lays
  // those lines out round it too (its PDF: the short lines above such a float go on to its right). Its own paragraph's
  // lines the browser wraps (it's a float there); text to its right only (a left float) is all these handle.
  for (const img of view.dom.querySelectorAll('img[style*="translate"]')) {
    const cs = getComputedStyle(img);
    if (cs.float !== 'left' || !/^0px -/.test(img.style.translate)) continue;
    const r = img.getBoundingClientRect();
    out.push({ left: r.left - (parseFloat(cs.marginLeft) || 0), right: r.right + (parseFloat(cs.marginRight) || 0), top: r.top, bottom: r.bottom, own: img.closest('.odx-content > *') });
  }
  return out;
}

// A paragraph's characters as drawn, in order, with their document positions, and its spacers: { chars, spacers }.
function charsOf(view, el) {
  const chars = [];
  const range = document.createRange();
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  for (let t = walker.nextNode(); t; t = walker.nextNode()) {
    if (t.parentElement.closest('.odx-wb, [contenteditable="false"]')) continue;
    let base;
    try { base = view.posAtDOM(t, 0); } catch { continue; }
    for (let i = 0; i < t.length; i++) {
      range.setStart(t, i);
      range.setEnd(t, i + 1);
      const rs = [...range.getClientRects()].filter((r) => r.height > 0);
      const r = rs[rs.length - 1];
      if (!r) continue;
      chars.push({ x0: r.left, x1: r.right, top: r.top, bottom: r.bottom, pos: base + i, space: /\s/.test(t.data[i]) });
    }
  }
  return chars;
}

// The characters in lines: a new line where one starts left of the last and lower down.
function linesOf(chars) {
  const lines = [];
  for (const c of chars) {
    const last = lines[lines.length - 1];
    if (last && !(c.x0 < last.x1 - 1 && c.top >= last.bottom - 2)) {
      last.chars.push(c);
      last.x1 = c.x1; last.top = Math.min(last.top, c.top); last.bottom = Math.max(last.bottom, c.bottom);
    } else lines.push({ chars: [c], x1: c.x1, top: c.top, bottom: c.bottom });
  }
  return lines;
}

/**
 * One pass: for each paragraph beside a both-sides shape, the first line that runs into its room gets a spacer at the
 * start of the first word that doesn't fit before it, as wide as takes that word to the room's right side; a spacer no
 * longer where its line meets the room (the text before it changed) goes, with those after it in the paragraph. True
 * when the spacers changed (lines move: measure again).
 */
export function applyWrapBoth(view) {
  const st = key.getState(view.state);
  // (A paragraph settles in a pass for each line beside a shape; this many means something keeps undoing them.)
  if (!st || st.passes > 80) return false;
  const zs = zones(view);
  if (!zs.length && !st.list.length) return false;
  let list = st.list.slice().sort((a, b) => a.pos - b.pos);
  let changed = false;
  if (!zs.length) { list = []; changed = true; }
  const doc = view.state.doc;
  doc.forEach((node, offset) => {
    if (!zs.length || !node.isTextblock) return;
    const el = view.nodeDOM(offset);
    if (!el || el.nodeType !== 1) return;
    const er = el.getBoundingClientRect();
    // (Its box down through its space after, as its last line's: below.)
    const after = Math.max(0, parseFloat(getComputedStyle(el).marginBottom) || 0);
    const near = zs.filter((z) => z.own !== el && er.bottom + after > z.top && er.top < z.bottom);
    const from = offset, to = offset + node.nodeSize;
    const own = list.filter((s) => s.pos > from && s.pos < to);
    if (!near.length) {
      if (own.length) { list = list.filter((s) => !own.includes(s)); changed = true; }
      return;
    }
    const lines = linesOf(charsOf(view, el));
    // (A line's box, as Word measures it against the shape: from the line above's foot, or the paragraph's top, to its
    // own. Word's PDF: a line whose text starts 1pt below the shape but whose box reaches it is split too.)
    const cs = getComputedStyle(el);
    let above = er.top + (parseFloat(cs.borderTopWidth) || 0) + (parseFloat(cs.paddingTop) || 0);
    for (const line of lines) { line.boxTop = Math.min(line.top, above); above = line.bottom; }
    // (A paragraph's last line down through its space after: Word's PDF, a line 7pt above a float, its space after
    // reaching it, goes beside it.)
    if (lines.length) lines[lines.length - 1].bottom += after;
    const beside = (top, bottom, z) => bottom > z.top + 0.25 && top < z.bottom - 0.25;
    // A spacer no longer ending at the room's right side on a line beside it (the text before it changed, or the lines
    // moved): it and those after it in the paragraph go, to be put back as the lines settle. (Its word may have gone on
    // to the next line, not fitting beside the room: the spacer still ends there. Its line by where it sits; one holding
    // nothing else, a line's height above it.)
    const lh = parseFloat(cs.lineHeight) || parseFloat(cs.fontSize) * 1.2 || 16;
    const els = [...el.querySelectorAll('.odx-wb')];
    const stale = own.find((s, i) => {
      const r = els[i]?.getBoundingClientRect();
      if (!r) return true;
      const line = lines.find((l) => r.top >= l.top - 1 && r.top <= l.bottom + 1);
      const [top, bottom] = line ? [line.boxTop, line.bottom] : [r.bottom - lh, r.bottom];
      return !near.some((z) => beside(top, bottom, z) && Math.abs(r.right - z.right) <= 1);
    });
    if (stale) {
      list = list.filter((s) => !(s.pos >= stale.pos && s.pos < to));
      changed = true;
      return;
    }
    for (const line of lines) {
      const z = near.find((zz) => beside(line.boxTop, line.bottom, zz));
      if (!z) continue;
      // The first character (not a space) in the room.
      const k = line.chars.findIndex((c) => !c.space && c.x1 > z.left + 0.5 && c.x0 < z.right - 0.5);
      if (k < 0) continue;
      let s = k;
      while (s > 0 && !line.chars[s - 1].space) s--;
      const c = line.chars[s];
      const w = z.right - c.x0;
      if (w <= 0.5 || list.some((x) => x.pos === c.pos)) continue;
      list.push({ pos: c.pos, w });
      list.sort((a, b) => a.pos - b.pos);
      changed = true;
      return; // (one a pass: the lines after it move)
    }
  });
  if (!changed) return false;
  view.dispatch(view.state.tr.setMeta(key, list).setMeta('addToHistory', false));
  return true;
}
