// Text on both sides of a floating shape (w:wrapSquare, Tight or Through, wrapText="bothSides"; the one-sided kinds are
// floats, reader.js anchorCss): Word splits each line beside the shape, filling the part
// from the margin to the shape's left wrap distance with as many whole words as fit, then going on from its right one
// (its PDFs). A browser float keeps text off one side only, so such a shape is set where it is, out of the flow
// (reader.js anchorCss, --odx-both: the room kept clear round it), and each line that runs into it gets a spacer at the
// first word that doesn't fit before it, as wide as the jump to its right side. Lines are measured as drawn; a pass
// adds one spacer to each paragraph that needs one, or takes away those that no longer fit, and the layout measures
// again (pages.js), so a paragraph is settled line by line, top down. The spacers are view decorations, kept through
// edits (mapped) and checked again on each pass.
// In a justified paragraph Word justifies each side of such a line on its own: the words before the shape spread to its
// left wrap distance, those after it to the margin (wrap-word2 wj-both). The browser would spread the line's room over
// all its spaces, so each side gets a word spacing of its own (decorations kept with the spacer), worked out from the
// words' widths unspread, and the spacer is as wide as takes it from the left side's end to the room's right side.
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { padOf } from '../ooxml/place.js';
import { dropSqueeze, squeezeFits } from './justify.js';

const key = new PluginKey('odx-wrapboth');
const spacer = (w) => () => {
  const el = document.createElement('span');
  el.className = 'odx-wb';
  el.contentEditable = 'false';
  el.style.width = `${w}px`;
  return el;
};
const spread = (r) => Decoration.inline(r.from, r.to, { style: `word-spacing:${r.ws.toFixed(3)}px` });
const decos = (doc, list) => DecorationSet.create(doc, list.flatMap((s) => [
  Decoration.widget(s.pos, spacer(s.w), { side: -1, key: `wb:${s.pos}:${s.w.toFixed(2)}`, ignoreSelection: true }),
  ...[s.l, s.r].filter((r) => r && r.to > r.from).map(spread),
]));
// A range's place after an edit; null when it went.
const mapRange = (m, r) => {
  if (!r) return null;
  const from = m.map(r.from, 1), to = m.map(r.to, -1);
  return to > from ? { from, to, ws: r.ws } : null;
};

/**
 * The lines split beside shapes in the paragraph from `from` to `to` (justify.js squeezes the others): the first
 * spacer's position (`first`: the words before the room on its line may yet go on the line above, squeezed), and where
 * the line after the last split one starts (`next`, null when not known). False when it has no spacers.
 */
export function splitOf(state, from, to) {
  const own = key.getState(state)?.list.filter((s) => s.pos > from && s.pos < to) || [];
  if (!own.length) return false;
  const lastOne = own[own.length - 1];
  return { first: own[0].pos, next: lastOne.next ?? null };
}

/**
 * Floats drawn up to end at their page's edge (pages.js floatsPast): the edge's place below the editor's top, by
 * picture. Where such a float is drawn depends on where its paragraph is, which the lines it wraps above it move: its
 * room is taken from the edge, which stays put, not from where a pass before drew it.
 */
export const drawnUpTo = new WeakMap();

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
        if (!r.deleted) list.push({ pos: r.pos, w: s.w, l: mapRange(tr.mapping, s.l), r: mapRange(tr.mapping, s.r), next: s.next == null ? null : tr.mapping.map(s.next) });
      }
      return { list, set: decos(tr.doc, list), passes: 0 };
    },
  },
  props: { decorations: (state) => key.getState(state).set },
});

/**
 * Shapes wrapped round in the headers and footers drawn on the sheets under `root` (reader.js anchorCss, --odx-hfwrap:
 * the sides the text takes and the room round it): Word wraps the body's text round them on every page showing them,
 * as round its own. The room as drawn on each page, reaching to the margin on a side the text doesn't take: a line
 * beside it then goes on past it (text on its right only) or stops before it (on its left only). pagefloat-word
 * pf-footer: a rectangle 400pt above its footer's paragraph, Text 4's lines on both sides of it.
 */
function storyZones(root) {
  const out = [];
  if (!root) return out;
  const MIN_GAP = 6; // px: room for no word at all, as reader.js anchorCss
  for (const img of root.querySelectorAll('.odx-hf img[style*="--odx-hfwrap"]')) {
    const [side, ...d] = img.style.getPropertyValue('--odx-hfwrap').trim().split(/\s+/);
    const [t, rt, b, l] = d.map((x) => parseFloat(x) || 0);
    const r = img.getBoundingClientRect();
    const col = img.closest('.odx-hf').getBoundingClientRect();
    if (!r.width || !r.height) continue;
    const z = { left: r.left - l, right: r.right + rt, top: r.top - t, bottom: r.bottom + b };
    const gapL = z.left - col.left, gapR = col.right - z.right;
    // (The side the text takes: both, its left or right only, or the larger; both when there's room for a word on each.)
    const takes = side === 'left' ? 'left' : side === 'right' ? 'right' : side === 'largest' ? (gapL > gapR ? 'left' : 'right')
      : gapL < MIN_GAP ? 'right' : gapR < MIN_GAP ? 'left' : 'both';
    // (On its left only: the line's spacer fills it to the margin, a px short so it stays on its line.)
    if (takes === 'right') z.left = Math.min(z.left, col.left);
    else if (takes === 'left') { z.right = col.right - 1; z.toEnd = true; }
    z.story = true;
    out.push(z);
  }
  return out;
}

// The room kept clear round each such shape, as drawn (its turned outline when turned): { left, right, top, bottom }.
function zones(view, stories) {
  const out = storyZones(stories);
  for (const img of view.dom.querySelectorAll('img[style*="--odx-both"]')) {
    const v = getComputedStyle(img).getPropertyValue('--odx-both').trim().split(/\s+/).map((x) => parseFloat(x) || 0);
    const r = img.getBoundingClientRect();
    if (!r.width || !r.height) continue;
    const [t, rt, b, l] = v.length === 4 ? v : [0, 0, 0, 0];
    // (Its own box: less the room drawn round its outline, reader.js padPlace's --odx-pad. floatshapes-word's 6pt-outlined
    // rectangle by Text 2, its lines' words otherwise kept that much further off each side.)
    const pad = padOf(img.getAttribute('style'));
    out.push({ left: r.left + pad - l, right: r.right - pad + rt, top: r.top + pad - t, bottom: r.bottom - pad + b });
  }
  // A float drawn up to end at its page's edge (pages.js floatsPast) over lines of paragraphs above its own: Word lays
  // those lines out round it too (its PDF: the short lines above such a float go on to its right). Its own paragraph's
  // lines the browser wraps (it's a float there). Its room from the edge it ends at (drawnUpTo); text on its left only
  // (a right float; wrap-word2 wl-up: the lines above it stop at its left) reaching to the line's end (toEnd).
  const domTop = view.dom.getBoundingClientRect().top;
  for (const img of view.dom.querySelectorAll('img[style*="translate"]')) {
    const cs = getComputedStyle(img);
    if ((cs.float !== 'left' && cs.float !== 'right') || !/^0px -/.test(img.style.translate)) continue;
    const r = img.getBoundingClientRect();
    const edge = drawnUpTo.get(img);
    const bottom = edge == null ? r.bottom : domTop + edge;
    out.push({ left: r.left - (parseFloat(cs.marginLeft) || 0), right: r.right + (parseFloat(cs.marginRight) || 0), top: bottom - r.height, bottom, own: img.closest('.odx-content > *'), toEnd: cs.float === 'right' });
  }
  return out;
}

let ctx = null;
// A text's widths unspread, in the font it's drawn in, as justify.js measures words (whose squeezing rule was fitted
// to those widths): its characters' (each the step from the word up to it, kerning as the page sets it) and its font
// size.
function measurer(host) {
  ctx = ctx || document.createElement('canvas').getContext('2d');
  const cs = getComputedStyle(host);
  const font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
  const ls = cs.letterSpacing === 'normal' ? '0px' : cs.letterSpacing;
  const kern = cs.fontKerning === 'none' ? 'none' : 'normal';
  const em = parseFloat(cs.fontSize) || 16;
  return (text) => {
    ctx.font = font;
    ctx.letterSpacing = ls;
    ctx.fontKerning = kern;
    const out = [];
    let prev = 0, word = 0;
    for (let i = 0; i < text.length; i++) {
      if (text[i] === ' ') { out.push(ctx.measureText(' ').width); word = i + 1; prev = 0; continue; }
      // (Measured over the word so far, so kerning counts; past 32 letters over the last 32 only, so a run without
      // spaces costs its length, not its square.)
      const from = Math.max(word, i - 31);
      if (from > word) prev = ctx.measureText(text.slice(from, i)).width;
      const w = ctx.measureText(text.slice(from, i + 1)).width;
      out.push(w - prev);
      prev = w;
    }
    return { w: out, em };
  };
}

// A paragraph's characters as drawn, in order, with their document positions; when `measure` (justified text), `nat`,
// a character's width unspread (measurer: Word's choice of words), `dw` as drawn unspread (the browser's: the word
// spacing that makes it fit), and `em` its font size; `plain`, not a tab or other white space.
function charsOf(view, el, measure) {
  const chars = [];
  const range = document.createRange();
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  for (let t = walker.nextNode(); t; t = walker.nextNode()) {
    if (t.parentElement.closest('.odx-wb, [contenteditable="false"]')) continue;
    let base;
    try { base = view.posAtDOM(t, 0); } catch { continue; }
    const nat = measure ? measurer(t.parentElement)(t.data) : null;
    for (let i = 0; i < t.length; i++) {
      range.setStart(t, i);
      range.setEnd(t, i + 1);
      const rs = [...range.getClientRects()].filter((r) => r.height > 0);
      const r = rs[rs.length - 1];
      if (!r) continue;
      const ch = t.data[i];
      const space = /\s/.test(ch);
      chars.push({ x0: r.left, x1: r.right, top: r.top, bottom: r.bottom, pos: base + i, space, plain: !space || ch === ' ', nat: nat ? nat.w[i] : r.right - r.left, dw: nat && ch === ' ' ? nat.w[i] : r.right - r.left, em: nat?.em || 0 });
    }
  }
  return chars;
}

/**
 * A justified line's split beside room `z` (Word: each side justified on its own), worked out from the characters'
 * widths unspread: where the spacer goes (`s`, the line's character it goes before) and starts, and the word spacing
 * of the words before it (`l`, spread to the room's left side) and after it (`r`, spread to the line's end `end`; none
 * on the paragraph's last line). `all`: the paragraph's characters, for the words that go on after the room. Null when
 * it can't be worked out (a tab, or anything but text, in the line).
 */
function splitJustified(line, all, z, end, squeeze) {
  const L = line.chars;
  if (L.some((c, i) => !c.plain || (i && c.pos !== L[i - 1].pos + 1))) return null;
  // (Unspread: where Word has each character, `xn`, and where the browser draws it, `xd`.)
  const xn = [], xd = [];
  let x = L[0].x0, y = L[0].x0;
  for (const c of L) { xn.push(x); xd.push(y); x += c.nat; y += c.dw; }
  const k = L.findIndex((c, i) => !c.space && xn[i] + c.nat > z.left + 0.5 && xn[i] < z.right - 0.5);
  if (k < 0) return null;
  let s = k;
  while (s > 0 && !L[s - 1].space) s--;
  // (A word Word squeezes in before the room, as on a line of its own: squeezeFits.)
  if (squeeze && s > 0) {
    let we = s;
    while (we + 1 < L.length && !L[we + 1].space) we++;
    let f = 0;
    while (L[f].space) f++;
    const sp = L.slice(f, s).filter((c) => c.space);
    const over = xn[we] + L[we].nat - z.left;
    if (we + 1 < L.length && squeezeFits(over, sp.length, sp.reduce((a, c) => a + c.nat, 0), sp.reduce((a, c) => a + c.em, 0), xn[we] + L[we].nat - xn[s])) {
      s = we + 1;
      while (s < L.length && L[s].space) s++;
      if (s >= L.length) return null;
    }
  }
  let e = s - 1;
  while (e >= 0 && L[e].space) e--;
  let start = xd[s], l = null;
  if (e >= 0) {
    let first = 0;
    while (L[first].space) first++;
    const n = L.slice(first, e + 1).filter((c) => c.space).length;
    const gap = z.left - (xd[e] + L[e].dw);
    const ws = n ? gap / n : 0;
    // (The spaces after the last word before the room take it too: the spacer starts after them.)
    start += (n ? gap : 0) + ws * (s - e - 1);
    l = { from: L[0].pos, to: L[s].pos, ws };
  }
  if (z.toEnd) return { s, start, l, r: null };
  // The words after the room: as many as fit to the line's end unspread, or Word squeezes in, spread to it unless
  // they're the paragraph's last (as drawn, a little short of it, so the browser fits them too).
  const room = end - z.right;
  const at = all.indexOf(L[s]);
  const n = { width: 0, spaces: 0, spaceW: 0, em: 0, drawn: 0 }; // the words taken
  const p = { space: 0, n: 0, em: 0, drawn: 0, word: 0, wordD: 0 }; // the spaces and word after them
  let last = -1, wordEnd = -1;
  // (The next word: on the line if it fits (1), or Word squeezes it in (2).)
  const takes = () => {
    if (last < 0 || n.width + p.space + p.word <= room) return 1;
    return squeeze && squeezeFits(n.width + p.space + p.word - room, n.spaces + p.n, n.spaceW + p.space, n.em + p.em, p.word) ? 2 : 0;
  };
  const take = () => {
    n.width += p.space + p.word; n.spaces += p.n; n.spaceW += p.space; n.em += p.em; n.drawn += p.drawn + p.wordD;
    last = wordEnd;
    Object.assign(p, { space: 0, n: 0, em: 0, drawn: 0, word: 0, wordD: 0 });
  };
  let i = at, full = false;
  for (; i < all.length; i++) {
    const c = all[i];
    if (i > at && (!c.plain || c.pos !== all[i - 1].pos + 1)) return { s, start, l, r: null };
    if (!c.space) { p.word += c.nat; p.wordD += c.dw; wordEnd = i; continue; }
    if (p.word) {
      const t = takes();
      if (!t) { full = true; break; }
      take();
      if (t === 2) { full = true; break; }
    }
    p.space += c.nat; p.n++; p.em += c.em; p.drawn += c.dw;
  }
  if (!full && p.word) {
    const t = takes();
    if (t) take();
    full = !t || t === 2;
  }
  if (last < 0) return { s, start, l, r: null };
  return { s, start, l, r: { from: L[s].pos, to: all[last].pos + 1, ws: !full || !n.spaces ? 0 : (room - 0.3 - n.drawn) / n.spaces } };
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
 * when the spacers changed (lines move: measure again). `stories`: the sheets' headers and footers, for their shapes
 * (storyZones).
 */
export function applyWrapBoth(view, stories = null) {
  const st = key.getState(view.state);
  // (A paragraph settles in a pass for each line beside a shape; this many means something keeps undoing them.)
  if (!st || st.passes > 80) return false;
  const zs = zones(view, stories);
  if (!zs.length && !st.list.length) return false;
  let list = st.list.slice().sort((a, b) => a.pos - b.pos);
  let changed = false;
  const fresh = []; // (the lines split this pass: from, to)
  if (!zs.length) { list = []; changed = true; }
  const doc = view.state.doc;
  doc.forEach((node, offset) => {
    if (!zs.length || !node.isTextblock) return;
    const el = view.nodeDOM(offset);
    if (!el || el.nodeType !== 1) return;
    const er = el.getBoundingClientRect();
    // (Its box down through its space after, as its last line's: below.)
    const after = Math.max(0, parseFloat(getComputedStyle(el).marginBottom) || 0);
    const cs = getComputedStyle(el);
    const justified = cs.textAlign === 'justify';
    // (The lines' end: a room taking text on its left only reaches to it, a px short so its spacer stays on its line;
    // half that when justified, the browser spreading what's left over the line, moving the spacer that much.)
    const end = er.right - (parseFloat(cs.borderRightWidth) || 0) - (parseFloat(cs.paddingRight) || 0);
    const near = zs.filter((z) => z.own !== el && er.bottom + after > z.top && er.top < z.bottom).map((z) => (z.toEnd ? { ...z, right: end - (justified ? 0.5 : 1) } : z));
    const from = offset, to = offset + node.nodeSize;
    const own = list.filter((s) => s.pos > from && s.pos < to);
    if (!near.length) {
      if (own.length) { list = list.filter((s) => !own.includes(s)); changed = true; }
      return;
    }
    const chars = charsOf(view, el, justified);
    const lines = linesOf(chars);
    // (A line's box, as Word measures it against the shape: from the line above's foot, or the paragraph's top, to its
    // own. Word's PDF: a line whose text starts 1pt below the shape but whose box reaches it is split too.)
    let above = er.top + (parseFloat(cs.borderTopWidth) || 0) + (parseFloat(cs.paddingTop) || 0);
    for (const line of lines) { line.boxTop = Math.min(line.top, above); above = line.bottom; }
    // (A paragraph's last line down through its space after: Word's PDF, a line 7pt above a float, its space after
    // reaching it, goes beside it. Not beside a header's or footer's shape: pagefloat-word pf-footer, Text 3's last line
    // above its rectangle, its space after reaching it, whole.)
    if (lines.length) lines[lines.length - 1].after = after;
    const beside = (top, bottom, z, extra = 0) => bottom + (z.story ? 0 : extra) > z.top + 0.25 && top < z.bottom - 0.25;
    // A spacer no longer ending at the room's right side on a line beside it (the text before it changed, or the lines
    // moved): it and those after it in the paragraph go, to be put back as the lines settle. (Its word may have gone on
    // to the next line, not fitting beside the room: the spacer still ends there. Its line by where it sits; one holding
    // nothing else, a line's height above it.)
    const lh = parseFloat(cs.lineHeight) || parseFloat(cs.fontSize) * 1.2 || 16;
    const els = [...el.querySelectorAll('.odx-wb')];
    const stale = own.find((s, i) => {
      const r = els[i]?.getBoundingClientRect();
      if (!r) return true;
      const line = lines.find((l) => r.top >= l.top - 1 && r.top <= l.bottom + (l.after || 0) + 1);
      const [top, bottom] = line ? [line.boxTop, line.bottom] : [r.bottom - lh, r.bottom];
      return !near.some((z) => beside(top, bottom, z, line?.after) && Math.abs(r.right - z.right) <= 1);
    });
    if (stale) {
      list = list.filter((s) => !(s.pos >= stale.pos && s.pos < to));
      changed = true;
      return;
    }
    for (const line of lines) {
      const z = near.find((zz) => beside(line.boxTop, line.bottom, zz, line.after));
      if (!z) continue;
      // Justified: each side spread on its own (splitJustified).
      const sp = justified ? splitJustified(line, chars, z, end, view.odx?.session?.compatMode >= 15) : null;
      if (sp) {
        const c = line.chars[sp.s];
        const w = z.right - sp.start;
        if (w <= 0.5 || list.some((x) => x.pos === c.pos)) continue;
        // (next: where the line after it starts, for justify.js; the word pushed on, when nothing follows the room.)
        list.push({ pos: c.pos, w, l: sp.l, r: sp.r, next: z.toEnd ? c.pos : sp.r ? sp.r.to : null });
        fresh.push([sp.l ? sp.l.from : c.pos, sp.r ? sp.r.to : c.pos + 1]);
        list.sort((a, b) => a.pos - b.pos);
        changed = true;
        return;
      }
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
  // (Word's squeezing of justified lines, justify.js, worked out for lines all the column's width: gone from those split
  // now, to be worked out again for the lines above and below the split ones.)
  view.dispatch(dropSqueeze(view.state.tr, fresh).setMeta(key, list).setMeta('addToHistory', false));
  return true;
}
