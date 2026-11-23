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
// view -> the sets of spacers its document has had, to see them come round again (applyWrapBoth). Afresh once fonts load.
const settled = new WeakMap();
let fontLoads = 0;
if (typeof document !== 'undefined' && document.fonts) document.fonts.addEventListener('loadingdone', () => { fontLoads++; });
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
 * Floats drawn up to end at their page's edge (pages.js floatsPast), and shapes with text on both sides at its bottom
 * margin (shapesPast): the edge's place below the editor's top, by picture. Where such a float is drawn depends on
 * where its paragraph is, which the lines it wraps above it move: its room is taken from the edge, which stays put, not
 * from where a pass before drew it.
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
        if (!r.deleted) list.push({ pos: r.pos, w: s.w, l: mapRange(tr.mapping, s.l), r: mapRange(tr.mapping, s.r), next: s.next == null ? null : tr.mapping.map(s.next), tail: s.tail, flush: s.flush });
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
  const domTop = view.dom.getBoundingClientRect().top;
  for (const img of view.dom.querySelectorAll('img[style*="--odx-both"]')) {
    const v = getComputedStyle(img).getPropertyValue('--odx-both').trim().split(/\s+/).map((x) => parseFloat(x) || 0);
    const r = img.getBoundingClientRect();
    if (!r.width || !r.height) continue;
    const [t, rt, b, l, fxT = 0] = v.length >= 4 ? v : [0, 0, 0, 0];
    // (Its own box: less the room drawn round its outline, reader.js padPlace's --odx-pad. floatshapes-word's 6pt-outlined
    // rectangle by Text 2, its lines' words otherwise kept that much further off each side.)
    const pad = padOf(img.getAttribute('style'));
    // (Drawn up to end at its page's bottom margin, pages.js shapesPast: its room from there, as a float's drawn up
    // (below), not from where a pass before drew it, a line lower once a line above it split.)
    const edge = img.style.translate ? drawnUpTo.get(img) : null;
    const dy = edge == null ? 0 : domTop + edge - r.bottom;
    out.push({ left: r.left + pad - l, right: r.right - pad + rt, top: r.top + dy + pad - t, bottom: r.bottom + dy - pad + b, fxT, host: img.closest('.odx-content > *') });
  }
  // A float drawn up to end at its page's edge (pages.js floatsPast) over lines of paragraphs above its own: Word lays
  // those lines out round it too (its PDF: the short lines above such a float go on to its right). Its own paragraph's
  // lines the browser wraps (it's a float there). Its room from the edge it ends at (drawnUpTo); text on its left only
  // (a right float; wrap-word2 wl-up: the lines above it stop at its left) reaching to the line's end (toEnd).
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
    if (we + 1 < L.length && squeezeFits(over, sp.length, sp.reduce((a, c) => a + c.nat, 0), sp.reduce((a, c) => a + c.em, 0), xn[we] + L[we].nat - xn[s], true)) {
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
    // (The first word too: one that doesn't fit leaves the side empty, as Word does: wrap-word3 sq-none, Text 23's first
    // line, nothing in the 28pt right of its rectangle.)
    if (n.width + p.space + p.word <= room) return 1;
    if (last < 0) return 0;
    return squeeze && squeezeFits(n.width + p.space + p.word - room, n.spaces + p.n, n.spaceW + p.space, n.em + p.em, p.word, true) ? 2 : 0;
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
  // (No word fits after the room: the spacer ends the line, and the browser spreading the line slid it to the line's end;
  // it reaches there itself, half a px short as one to the line's end (applyWrapBoth's `end`).)
  if (last < 0) return { s, start, l, r: null, flush: Math.max(0, room - 0.5) };
  // (One word after the room, the line going on: Word sets it at the line's end (wrap-word3 sq-*, a rectangle near the
  // right margin, "garden" alone flush right beside it). With no space there to spread, the browser spread its room over
  // the words before the room, the spacer moved off the room's side and went and came back until the passes ran out:
  // the spacer takes it, `flush`, past the room's right side. Drawn a little wider than it measures, with no space to
  // narrow, the spacer gives up what it needs, up to 3px: sq-none's Text 23, "harbor" by 0.3px.)
  const short = room - 0.3 - n.drawn;
  const flush = full && !n.spaces ? (short >= -3 ? short : 0) : 0;
  return { s, start, l, r: { from: L[s].pos, to: all[last].pos + 1, ws: !full || !n.spaces ? 0 : (room - 0.3 - n.drawn) / n.spaces }, flush };
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
  let seen = settled.get(view);
  if (!seen || seen.doc !== view.state.doc || seen.fonts !== fontLoads) settled.set(view, seen = { doc: view.state.doc, fonts: fontLoads, states: [], done: false });
  if (seen.done) return false;
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
    // (A room anchored in a later paragraph, `later`: one placed above its paragraph's text, reaching up to this one's.)
    const near = zs.filter((z) => z.own !== el && er.bottom + after > z.top && er.top < z.bottom)
      .map((z) => (z.toEnd ? { ...z, right: end - (justified ? 0.5 : 1) } : z))
      .map((z) => (z.host && z.host !== el && el.compareDocumentPosition(z.host) & Node.DOCUMENT_POSITION_FOLLOWING ? { ...z, later: true } : z))
      // (And an earlier paragraph's placed below this one's top, `below`, as a later one's, its space after not reaching it:
      // wrap-word4 vt-thin, Text 20's rectangle saved 586pt below its paragraph at page 4's foot, Text 25's last line
      // whole above it in Word's print, its space after reaching 4.7pt into it; split, Text 25 grew a line and Text 26
      // went a line down, its own rectangle drawn up over Text 25.)
      .map((z) => (!z.later && z.host && z.host !== el && z.top > er.top ? { ...z, below: true } : z));
    const from = offset, to = offset + node.nodeSize;
    const own = list.filter((s) => s.pos > from && s.pos < to);
    if (!near.length) {
      if (own.length) { list = list.filter((s) => !own.includes(s)); changed = true; }
      return;
    }
    const chars = charsOf(view, el, justified);
    const lines = linesOf(chars);
    // (A line's box, as Word measures it against the shape: from its own top, the lead the line has above its text as the
    // paragraph's first line shows it, to its text's foot. Not from the line above's foot: the room under a shape ending
    // in the 2.6px a line spaced 1.15 keeps below its text took the next line in too, where Word sets that line whole
    // (wrap-word3 vt-*: rectangles 0.5pt apart in height, the rows beside each as in Word's print).)
    const top0 = er.top + (parseFloat(cs.borderTopWidth) || 0) + (parseFloat(cs.paddingTop) || 0);
    const lead = lines.length ? Math.max(0, lines[0].top - top0) : 0;
    let above = top0;
    for (const line of lines) { line.boxTop = Math.max(Math.min(line.top, above), line.top - lead); above = line.bottom; }
    // (A paragraph's last line down through its space after: Word's PDF, a line 7pt above a float, its space after
    // reaching it, goes beside it. Not beside a header's or footer's shape: pagefloat-word pf-footer, Text 3's last line
    // above its rectangle, its space after reaching it, whole.)
    // (Nor does it, nor a line, reach into the room a later paragraph's shape's effect extent takes above it: wrap-word3
    // vt-thick, each paragraph's last line, its space after reaching 3pt into the extent of the 6pt-outlined rectangle
    // below, whole in Word's print; sq-thick, Text 5's last line into that of Text 6's, placed 7.5pt above its
    // paragraph. Split, the paragraph grew a line, moving the room down a line too, and the spacers came and went until
    // the passes ran out.) Nor its space after into a later paragraph's shape at all: wrap-word3 vt-none's Text 19 and
    // sq-none's Text 5, their last lines whole in Word's print above rectangles Word drew up 4.7 and 7.2pt (saved so),
    // their space after reaching them; split, the paragraph below and its rectangle went a line down, Text 6 moved on.
    if (lines.length) lines[lines.length - 1].after = after;
    // (Into the room by any of it, at its top or its foot: wrap-word4 vt-thin, Text 1's fourth line reaching 0.11px into
    // the room under its rectangle, split beside it in Word's print; vt-none, Text 3's 0.06px short of it, whole.)
    const beside = (top, bottom, z, extra = 0) => bottom + (z.story || z.later || z.below ? 0 : extra) > z.top + (z.later ? z.fxT || 0 : 0) + 0.03 && top < z.bottom - 0.03;
    // The room a line from `top` to `bottom` (through `extra` below, its space after) meets: the first shape's beside it,
    // and any other's beside it that overlaps that one across, as one (wrap-word4 sq-thick: Text 16's rectangle, saved at
    // page 4's foot, over Text 21's own; Word's lines there go round both, the words before the room ending at the one's
    // left, those after it starting at the other's right). (A few times over, for a chain of them; the shapes looked at
    // once each a time.)
    const roomAt = (top, bottom, extra) => {
      const at = near.filter((zz) => beside(top, bottom, zz, extra));
      if (at.length < 2) return at[0] || null;
      let z = at[0];
      for (let round = 0; round < 3; round++) {
        let grew = false;
        for (const o of at) {
          if (o === z || o.left >= z.right || o.right <= z.left) continue;
          if (o.left < z.left || o.right > z.right) { z = { ...z, left: Math.min(z.left, o.left), right: Math.max(z.right, o.right), toEnd: z.toEnd || o.toEnd }; grew = true; }
        }
        if (!grew) break;
      }
      return z;
    };
    // A spacer no longer ending at the room's right side on a line beside it (the text before it changed, or the lines
    // moved): it and those after it in the paragraph go, to be put back as the lines settle. (Its word may have gone on
    // to the next line, not fitting beside the room: the spacer still ends there. Its line by where it sits; one holding
    // nothing else, a line's height above it.)
    const lh = parseFloat(cs.lineHeight) || parseFloat(cs.fontSize) * 1.2 || 16;
    // (A `tail` spacer, below, keeps its line's space after and the line it added: the room it met, a later paragraph's,
    // went down that line with it.)
    const pitch = Math.max(lh, lines.length > 1 ? lines[lines.length - 1].top - lines[lines.length - 2].top : 0);
    const els = [...el.querySelectorAll('.odx-wb')];
    const stale = own.find((s, i) => {
      const r = els[i]?.getBoundingClientRect();
      if (!r) return true;
      const line = lines.find((l) => r.top >= l.top - 1 && r.top <= l.bottom + (l.after || 0) + 1);
      const [top, bottom] = line ? [line.boxTop, line.bottom] : [r.bottom - lh, r.bottom];
      const z = roomAt(top, bottom, s.tail ? after + pitch : line?.after);
      if (!z) return true;
      const off = r.right - z.right - (s.flush || 0);
      // (A centred or right-aligned line moves back by half or all of a spacer put in it, so the spacer ends that far short
      // of the room's side. Taken for stale it went and came back until the passes ran out, the pages undrawn all that time:
      // a centred line of 28 characters held a 32-page document for 81 passes, and a print taken meanwhile had no pages.)
      const back = justified ? 0 : cs.textAlign === 'center' ? s.w / 2 : cs.textAlign === 'right' ? s.w : 0;
      return Math.min(Math.abs(off), Math.abs(off + back)) > 1;
    });
    if (stale) {
      list = list.filter((s) => !(s.pos >= stale.pos && s.pos < to));
      changed = true;
      return;
    }
    for (const line of lines) {
      const z = roomAt(line.boxTop, line.bottom, line.after);
      if (!z) continue;
      // (Beside it through its space after only, `tail`: split, its words go on to a new last line and it isn't the last
      // any more, but Word, laying the lines out in turn, has split it already; so it keeps that space after while it
      // stays (the stale test above), and a paragraph takes one such. Otherwise the spacer came and went until the passes
      // ran out, the pages never laid out: wrap-word3 vt-none, Text 19's last line above the room of Text 20's rectangle,
      // placed 4.5pt above its paragraph.)
      const tail = !!line.after && !beside(line.boxTop, line.bottom, z);
      if (tail && own.some((s) => s.tail)) continue;
      // Justified: each side spread on its own (splitJustified).
      const sp = justified ? splitJustified(line, chars, z, end, view.odx?.session?.compatMode >= 15) : null;
      if (sp) {
        const c = line.chars[sp.s];
        const w = z.right - sp.start + (sp.flush || 0);
        if (w <= 0.5 || list.some((x) => x.pos === c.pos)) continue;
        // (next: where the line after it starts, for justify.js; the word pushed on, when nothing follows the room.)
        list.push({ pos: c.pos, w, l: sp.l, r: sp.r, next: z.toEnd ? c.pos : sp.r ? sp.r.to : null, tail, flush: sp.flush || 0 });
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
      list.push({ pos: c.pos, w, tail });
      list.sort((a, b) => a.pos - b.pos);
      changed = true;
      return; // (one a pass: the lines after it move)
    }
  });
  if (!changed) return false;
  // Spacers come back to a set they've had: two paragraphs undoing each other, each one's spacer moving the other's line
  // (a 41-page document went between two sets for 81 passes, 12 seconds, and a print took whichever it found). The set of
  // that round with the most spacers, the same one whichever pass got there first, and no more passes for this document.
  const sigOf = (l) => l.map((s) => `${s.pos}:${s.w.toFixed(2)}`).join(',');
  const sig = sigOf(list);
  const again = seen.states.findIndex((x) => x.sig === sig);
  if (again >= 0) {
    const round = seen.states.slice(again).sort((a, b) => b.list.length - a.list.length || (a.sig < b.sig ? -1 : 1));
    list = round[0].list;
    seen.done = true;
    if (sigOf(st.list.slice().sort((a, b) => a.pos - b.pos)) === round[0].sig) return false;
  } else seen.states.push({ sig, list });
  // (Word's squeezing of justified lines, justify.js, worked out for lines all the column's width: gone from those split
  // now, to be worked out again for the lines above and below the split ones.)
  view.dispatch(dropSqueeze(view.state.tr, fresh).setMeta(key, list).setMeta('addToHistory', false));
  return true;
}
