// Line pitch as Word lays it. Word stacks lines at the font's line height, unrounded (Times New Roman 10pt:
// 11.5pt = 15.333px a line). Edge rounds the font's ascent and descent to whole px and truncates each half of
// the leading to 1/64px, so the same line comes out shorter (15.3125px), by an amount that depends on the font,
// size and line height: up to ~1px over a page, enough to fit a line Word moves to the next page.
// For each paragraph style in use, a hidden probe finds the line height (just above the intended one) whose laid-out
// pitch comes closest, and the paragraph gets that value as a node decoration: a number for "auto" spacing (runs
// in other sizes still scale from it), a length for exact and at-least spacing (as they were).
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { Mark } from 'prosemirror-model';
import { schema } from './schema.js';
import { lineHeightOf } from './util.js';
import { pictureFix, pictureAlone } from './picfix.js';
import { mixedLines } from './mixlines.js';

const key = new PluginKey('odx-lhfix');
export const lineFixPlugin = new Plugin({
  key,
  state: {
    init: () => DecorationSet.empty,
    apply(tr, set) {
      const next = tr.getMeta(key);
      return next ? DecorationSet.create(tr.doc, next) : set.map(tr.mapping, tr.doc);
    },
  },
  props: { decorations: (state) => key.getState(state) },
});

const cache = new Map(); // paragraph and run styles, line height -> chosen line height in px (or null: leave it)
const OURS = /;?\s*--odx-lhv:\s*[\d.]+(px)?;?\s*$/; // the decoration's part of a paragraph's style attribute

/**
 * Edge's distance between two lines' tops (px) for a paragraph laid out with line height `lh` (px), measured on a
 * copy of it (same classes and style, inside a copy of the editor's content box) holding text in a copy of its
 * first run's elements: a line is as tall as the runs on it, each rounded on its own.
 */
function pitch(box, el, chain, lh, length) {
  const q = el.cloneNode(false);
  q.removeAttribute('data-odx-lh');
  q.setAttribute('style', (el.getAttribute('style') || '').replace(OURS, ''));
  q.style.setProperty('--odx-lhv', length ? `${lh}px` : String(lh / parseFloat(getComputedStyle(el).fontSize)));
  let inner = q;
  for (const c of chain) {
    const k = c.cloneNode(false);
    // (Not a run box of Word's from an earlier pass, mixlines.js: the probe measures the stylesheet's line height.)
    if (el.dataset.odxMix === '1') for (const p of ['line-height', 'vertical-align', 'position', 'top']) k.style.removeProperty(p);
    inner.append(k);
    inner = k;
  }
  box.append(q);
  inner.textContent = 'x';
  const one = q.getBoundingClientRect().height;
  inner.textContent = 'x\nx\nx\nx\nx';
  const five = q.getBoundingClientRect().height;
  q.remove();
  return (five - one) / 4;
}

/**
 * The line height (px) whose pitch is nearest `want`, or null when Edge already lays it within 0.002px. `where`: the
 * kind of box probed in, part of what's remembered (a header's paragraph styles aren't the body's).
 */
function bestHeight(box, el, chain, want, length, where = '') {
  const k = [where, el.className, (el.getAttribute('style') || '').replace(OURS, ''), ...chain.map((c) => `${c.tagName}.${c.className}:${c.getAttribute('style') || ''}`), want, length].join('|');
  if (cache.has(k)) return cache.get(k);
  let best = null;
  let err = Math.abs(pitch(box, el, chain, want, length) - want);
  if (err > 0.002) {
    for (let i = 1; i <= 4; i++) {
      const c = want + i / 128;
      const e = Math.abs(pitch(box, el, chain, c, length) - want);
      if (e < err - 1e-6) { err = e; best = c; }
    }
  }
  cache.set(k, best);
  return best;
}

/**
 * A paragraph's picture-line fix (picfix.js) with its pictures by offset in the paragraph, or null: for one with a picture
 * on a line with text (pararules.js odx-pictext), measured with the line height chosen (`value`, or its own).
 */
function picture(view, node, pos, el, box, value) {
  if (!el.querySelector('img:not(.ProseMirror-separator)')) return null;
  const fix = el.classList.contains('odx-pictext') ? pictureFix(el, box, value ?? '') : pictureAlone(el, box);
  if (!fix) return null;
  const mts = [];
  // (Looked up by picture, not searched for each: a paragraph of thousands of pictures.)
  const byImg = new Map(fix.mts);
  node.forEach((c, off) => {
    const dom = c.type.name === 'image' && view.nodeDOM(pos + 1 + off);
    if (!dom) return;
    let mt = byImg.get(dom);
    if (mt == null) for (const img of dom.querySelectorAll?.('img') || []) if ((mt = byImg.get(img)) != null) break;
    if (mt != null) mts.push([off, mt]);
  });
  return { d: fix.d, mts };
}

/**
 * Word's boxes for the runs of a paragraph of mixed sizes, fonts or raises (mixlines.js), as [from, to, css] by offset
 * from the paragraph's position, or null.
 */
function mixed(view, node, pos, el, box) {
  // (Text all in one formatting has one box: nothing to read from the page. Most paragraphs.)
  let marks = null, one = true;
  node.descendants((c) => { if (c.isText) { if (!marks) marks = c.marks; else if (one && !Mark.sameSet(marks, c.marks)) one = false; } return one; });
  if (one && /[^ \u200b]/.test(node.textContent)) return null; // (nothing but spaces in one size: the mark's line, mixlines.js)
  const got = mixedLines(el, box);
  if (!got) return null;
  const out = [];
  try {
    for (const [span, css] of got.spans) {
      const from = view.posAtDOM(span.firstChild, 0) - pos;
      const to = from + span.firstChild.data.length;
      if (!(from > 0) || to > view.state.doc.nodeAt(pos).nodeSize - 1) return null;
      out.push([from, to, css]);
    }
  } catch { return null; }
  out.mark = !!got.mark;
  return out;
}

/** The pitch (px) of the document's grid of lines (w:docGrid type lines or linesAndChars), or 0. */
export function lineGrid(doc) {
  const g = /<w:docGrid\b[^>]*>/.exec(doc.attrs.sectPr || '')?.[0] || '';
  return /w:type="(lines|linesAndChars)"/.test(g) ? Number(/w:linePitch="(\d+)"/.exec(g)?.[1] || 0) / 15 : 0;
}

let last = '';
let gridPitch = 0;
// Results by paragraph node ({ value, want }, null for none), until fonts load: probing every paragraph again on each
// layout pass was a long document's single largest cost after the table measuring.
let memo = new WeakMap();
let memoFonts = -1;
let fontLoads = 0;
if (typeof document !== 'undefined' && document.fonts) document.fonts.addEventListener('loadingdone', () => { fontLoads++; });
/** Recomputes the paragraphs' line heights; true when the decorations changed (measure again once drawn). */
export function applyLineFix(view) {
  const decos = [];
  // Probes go in a copy of the content box (its classes give the document's styles), off to the side.
  const box = document.createElement('div');
  box.className = view.dom.className;
  // (Laid out though off screen: a large document's paragraphs skip layout there, editor.css, and every probe measured 0,
  // leaving a 2,000-paragraph document's lines up to 0.02pt short each, 1px a page.)
  box.classList.add('odx-probes');
  box.setAttribute('aria-hidden', 'true');
  box.style.cssText = `position:absolute;left:-99999px;top:0;width:${view.dom.clientWidth}px;visibility:hidden;contain:layout`;
  view.dom.parentElement.append(box);
  if (memoFonts !== fontLoads) { memo = new WeakMap(); memoFonts = fontLoads; }
  // A document grid of lines (w:docGrid type lines, linesAndChars): Word sets each line of a paragraph spaced "auto" on
  // the grid, its pitch (w:linePitch) times the spacing, or the whole number of pitches its text needs, with the text
  // centred in it (Word's PDFs of the generated documents, tools/doc-generator: 11pt at 1.08 lines on an 18pt grid
  // 25.9px a line, 16pt headings 48px, double spacing 48px).
  const pitch = lineGrid(view.state.doc);
  if (pitch !== gridPitch) { memo = new WeakMap(); gridPitch = pitch; }
  // A paragraph's decorations from its result: its line height, and a picture line's fix (picfix.js) on it and its pictures.
  const put = (pos, node, r) => {
    if (!r) return;
    if (r.value != null) decos.push(Decoration.node(pos, pos + node.nodeSize, { style: `--odx-lhv:${r.value}`, 'data-odx-lh': String(r.want), ...(r.grid ? { 'data-odx-grid': '1' } : {}) }, { key: `lh:${r.value}${r.grid ? 'g' : ''}` }));
    if (r.mix) {
      decos.push(Decoration.node(pos, pos + node.nodeSize, { 'data-odx-mix': r.mix.mark ? 'mark' : '1' }, { key: r.mix.mark ? 'mxm' : 'mx' }));
      for (const [from, to, css] of r.mix) decos.push(Decoration.inline(pos + from, pos + to, { style: css }, { key: `mx:${css}` }));
    }
    if (!r.pic) return;
    decos.push(Decoration.node(pos, pos + node.nodeSize, { style: `--odx-picb:${r.pic.d}px` }, { key: `pb:${r.pic.d}` }));
    for (const [off, mt] of r.pic.mts) decos.push(Decoration.node(pos + 1 + off, pos + 2 + off, { style: `--odx-picmt:${mt}px${mt < 0 ? `;--odx-picclip:${-mt}px` : ''}` }, { key: `pm:${mt}` }));
  };
  view.state.doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.paragraph) return true;
    // An unchanged paragraph keeps its result (ProseMirror keeps the node): no probing again.
    if (memo.has(node)) {
      put(pos, node, memo.get(node));
      return false;
    }
    const el = view.nodeDOM(pos);
    if (!el || el.nodeType !== 1) return false;
    memo.set(node, null); // (none, unless found below)
    const cs = getComputedStyle(el);
    const fs = parseFloat(cs.fontSize);
    // The value Edge used before any fix of ours (a decoration from an earlier pass shows in the data attribute).
    const want = el.dataset.odxLh ? Number(el.dataset.odxLh) : parseFloat(lineHeightOf(cs));
    if (!want || !fs) return false;
    // The text's own elements (Word documents often set the font and size on runs, not the paragraph). Not an equation's:
    // its parts (a script with one of its two children copied) aren't a line of text, and Firefox reports them as errors.
    const chain = chainOf(el);
    const length = cs.getPropertyValue('--odx-lhlen').trim() === '1';
    // (Not in a table: Word sets a cell's lines as without the grid.)
    if (pitch && !length && parseFloat(cs.getPropertyValue('--odx-runmult')) !== 0 && !el.closest('td')) {
      const mult = parseFloat(cs.getPropertyValue('--odx-mult')) || 1;
      const on = Math.max(pitch * mult, Math.ceil(want / pitch - 1e-3) * pitch);
      const value = `${on.toFixed(4)}px`;
      const r = { value, want, grid: true, pic: picture(view, node, pos, el, box, value) };
      memo.set(node, r);
      put(pos, node, r);
      return false;
    }
    const best = bestHeight(box, el, chain, want, length);
    const value = best == null ? null : length ? `${best.toFixed(4)}px` : `${(best / fs).toFixed(6)}`;
    const pic = picture(view, node, pos, el, box, value);
    const mix = mixed(view, node, pos, el, box);
    if (value == null && !pic && !mix) return false;
    const r = { value, want, pic, mix };
    memo.set(node, r);
    put(pos, node, r);
    return false;
  });
  box.remove();
  overfullCells(view, decos);
  wideLabels(view, decos);
  const sig = decos.map((d) => `${d.from}:${d.to}:${d.spec.key}`).join(','); // (positions here, not in the keys: pages.js applyPushes)
  if (sig === last) return false;
  last = sig;
  view.dispatch(view.state.tr.setMeta(key, decos).setMeta('addToHistory', false));
  return true;
}

/**
 * Cells of exact-height rows holding more than the row has room for, over several paragraphs: marked (odx-exover), so
 * the row keeps its height and the text starts at the cell's top, cut at the row's foot, as in Word. (One paragraph is
 * held to the row's room by its own max-height, editor.css; a browser's cell is as tall as all it holds, and the row
 * with it: a sample's 78px row holding five lines and an empty paragraph came out 13px taller, its other cells' centred
 * text 6.5px lower.) Measured from the paragraphs' own boxes, which the mark doesn't move.
 */
function overfullCells(view, decos) {
  for (const tr of view.dom.querySelectorAll('tr[data-exact]')) {
    const h = parseFloat(tr.style.getPropertyValue('--odx-trh'));
    if (!(h > 0)) continue;
    for (const td of tr.cells) {
      if (td.childElementCount < 2 || td.rowSpan > 1) continue; // (a cell merged down several rows has their height, not this one's)
      const first = td.firstElementChild.getBoundingClientRect(), end = td.lastElementChild.getBoundingClientRect();
      if (!first.height && !end.height) continue; // not laid out
      const cs = getComputedStyle(td);
      const room = h - (parseFloat(cs.paddingTop) || 0) - (parseFloat(cs.paddingBottom) || 0);
      const held = end.bottom - first.top + Math.max(0, parseFloat(getComputedStyle(td.firstElementChild).marginTop) || 0);
      if (held <= room + 0.5) continue;
      let pos;
      try { pos = view.posAtDOM(td, 0) - 1; } catch { continue; }
      const cell = pos >= 0 ? view.state.doc.nodeAt(pos) : null;
      if (cell && /^table_(cell|header)$/.test(cell.type.name)) decos.push(Decoration.node(pos, pos + cell.nodeSize, { class: 'odx-exover' }, { key: 'exover' }));
    }
  }
}

let labelMemo = new WeakMap(); // paragraph node -> the room its wide label needs (px), 0 for none
let labelFonts = -1;
let labelCtx = null;
/**
 * List labels wider than the room before their text (pararules.js data-ltab): the text starts at the next tab stop after
 * the label's end, as in Word, where it had followed the label with nothing between ("Appendix A." and its title). The
 * label's width from its own font, measured apart from the page (the label's box is as wide as the room it's given).
 */
function wideLabels(view, decos) {
  if (labelFonts !== fontLoads) { labelMemo = new WeakMap(); labelFonts = fontLoads; }
  for (const el of view.dom.querySelectorAll('p.odx-li[data-ltab]')) {
    let pos;
    try { pos = view.posAtDOM(el, 0) - 1; } catch { continue; }
    const node = pos >= 0 ? view.state.doc.nodeAt(pos) : null;
    if (!node || node.type !== schema.nodes.paragraph) continue;
    let room = labelMemo.get(node);
    if (room == null) {
      room = 0;
      const [start, gap, dts, ...own] = el.dataset.ltab.split(',').map(Number);
      const label = el.dataset.label || '';
      const cs = label ? getComputedStyle(el, '::before') : null;
      const size = cs ? parseFloat(cs.fontSize) || 0 : 0;
      // (No wider than its letters could be: most labels, without measuring.)
      // (Measured where the label may be wider than its room, or the paragraph has a stop of its own inside that room:
      // the text starts at such a stop when it comes after the label's end, as after a tab typed there. Word's prints of
      // 12 generated lists, the number at 100 twips and the text's indent at 1540: a stop at 300, 819 or 1200 twips
      // takes the text of "1.", only the one at 1200 that of "Section 1.", which ends past the others; a stop on the
      // indent or past it, or none, leaves the text at the indent.)
      const inside = own.some((t) => t > start && t < start + gap);
      if (size && (inside || label.length * size * 1.2 > gap / 15)) {
        labelCtx = labelCtx || document.createElement('canvas').getContext('2d');
        labelCtx.font = `${cs.fontStyle} ${cs.fontWeight} ${size}px ${cs.fontFamily}`;
        const width = labelCtx.measureText(label).width + label.length * (parseFloat(cs.letterSpacing) || 0);
        // (Wider by anything at all: a sample's heading number 24.09px wide in 24px of room has its text at the next stop.)
        const before = inside ? own.find((t) => t > start + width * 15 + 1 && t < start + gap) : undefined;
        if (before != null && width <= gap / 15 + 0.03) room = Math.round(((before - start) / 15) * 100) / 100;
        else if (width > gap / 15 + 0.03) {
          const end = start + width * 15;
          const mine = own.find((t) => t > end + 1);
          const pitch = dts > 0 ? dts : 720;
          const stop = mine ?? (Math.floor((Math.max(end, own.length ? own[own.length - 1] : 0) + 1) / pitch) + 1) * pitch;
          room = Math.round(((stop - start) / 15) * 100) / 100;
        }
      }
      labelMemo.set(node, room);
    }
    if (room > 0) decos.push(Decoration.node(pos, pos + node.nodeSize, { style: `--odx-hangw:${room}px` }, { key: `lw:${room}` }));
  }
}

/** A paragraph's first text's elements below it (where its line height is probed), or []. */
function chainOf(el) {
  const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.data.trim() && !n.parentElement.closest('math') ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP) });
  const text = walk.nextNode();
  const chain = [];
  for (let n = text?.parentElement; n && n !== el; n = n.parentElement) chain.unshift(n);
  return chain;
}

// Header and footer paragraphs' results by paragraph node ({ value, want, mix: [css by odx-t span], spans }), until
// fonts load: a story is drawn again for every page.
let storyMemo = new WeakMap();
let storyFonts = -1;
/**
 * The same line heights for the header and footer paragraphs under `root` (drawn, in the document): each paragraph's
 * pitch as Word's (as the body's: a 10pt Times New Roman footer line 15.31px, not 15.33), and Word's run boxes in a
 * line of mixed sizes (mixlines.js: sample S078's footer, 10pt text and a 12pt run raised 3pt, drawn 0.7px low).
 * `nodeOf`: a paragraph element's node, for remembering its result (or undefined). Not on a document grid (lines on a
 * grid are the body's: left as they were).
 */
export function storyLineFix(root, nodeOf) {
  if (storyFonts !== fontLoads) { storyMemo = new WeakMap(); storyFonts = fontLoads; }
  let box = null;
  const probeBox = (el) => {
    if (box) return box;
    const hf = el.closest('.odx-hf') || root;
    box = hf.cloneNode(false);
    box.removeAttribute('id');
    box.classList.add('odx-probes');
    box.setAttribute('aria-hidden', 'true');
    box.style.cssText = `position:absolute;left:-99999px;top:0;width:${hf.clientWidth}px;visibility:hidden;contain:layout;white-space:pre-wrap`;
    hf.parentElement.append(box);
    return box;
  };
  try {
    for (const el of root.querySelectorAll('p')) {
      const node = nodeOf?.(el);
      const spans = el.querySelectorAll('span.odx-t');
      let r = node && storyMemo.get(node);
      if (!r || r.spans !== spans.length) {
        const cs = getComputedStyle(el);
        const fs = parseFloat(cs.fontSize);
        const want = parseFloat(lineHeightOf(cs));
        const length = cs.getPropertyValue('--odx-lhlen').trim() === '1';
        r = { value: null, want, mix: null, spans: spans.length };
        if (want && fs) {
          const best = bestHeight(probeBox(el), el, chainOf(el), want, length, 'hf');
          r.value = best == null ? null : length ? `${best.toFixed(4)}px` : `${(best / fs).toFixed(6)}`;
        }
        if (r.value != null) el.style.setProperty('--odx-lhv', r.value);
        const got = mixedLines(el, probeBox(el));
        if (got) {
          const at = new Map([...spans].map((s, i) => [s, i]));
          r.mix = got.spans.map(([s, css]) => [at.get(s), css]);
          r.mark = !!got.mark;
        }
        if (node) storyMemo.set(node, r);
      } else if (r.value != null) el.style.setProperty('--odx-lhv', r.value);
      if (r.value != null) el.dataset.odxLh = String(r.want);
      if (r.mix) {
        for (const [i, css] of r.mix) if (spans[i]) spans[i].style.cssText += `;${css}`;
        el.dataset.odxMix = r.mark ? 'mark' : '1';
      }
    }
  } finally { box?.remove(); }
}
