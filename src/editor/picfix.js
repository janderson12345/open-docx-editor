// A line holding an inline picture beside text, as tall as Word makes it. Word's picture line is the picture (its
// effect extents included) standing on the baseline, then the text font's own descent (hhea: the line gap goes above
// the text) and, under "multiple" spacing, all of the extra (..\picmatrix-word: Aptos, Calibri and Times New Roman at
// single, 1.08, 1.15, 1.16 and 1.5, pictures 9 to 55.3pt tall). The browser puts the text's descent and half its
// leading under the baseline (the line gap's half and the extra's): Calibri's single-spaced picture lines came out
// 1.2px tall, Aptos's at 1.16 0.55px. Under "at least" spacing Word's line is the larger of the length and the picture
// and descent, its baseline the descent above the line's foot; under exact spacing the line stays the length (the
// picture doesn't make it taller).
// Measured per paragraph (lhfix.js, once its line height is chosen): the difference for each line holding a picture
// taller than the text, added (or taken off) under the paragraph by --odx-picb (editor.css), and for at-least and exact
// lines a margin above each picture (--odx-picmt) that puts its line's top where Word's is.
import { ASCENT } from './baseline.js';
import { layoutRect, lineHeightOf } from './util.js';

// A font's descent where Word takes it from the Windows metrics (OS/2 usWinDescent, larger than hhea's): Calibri's 0.2686em,
// not 0.25 (its picture lines 0.2pt short a line with hhea's; Aptos and Times New Roman sit right on hhea's).
const DESCENT = { calibri: 0.2686, 'calibri light': 0.2686, carlito: 0.2686 };

const cache = new Map();

/** The browser's room above and below the baseline (px) of a line of the paragraph's text: [above, below]. */
function strut(box, el, chain, lhv) {
  const k = [el.className, (el.getAttribute('style') || '').replace(/;?\s*--odx-picb:[^;]*/g, ''), lhv,...chain.map((c) => `${c.tagName}.${c.className}:${c.getAttribute('style') || ''}`)].join('|');
  if (cache.has(k)) return cache.get(k);
  const q = el.cloneNode(false);
  q.removeAttribute('data-odx-lh');
  q.classList.remove('odx-pictext-end');
  // (Less this module's and lhfix.js's own parts: the line height chosen goes in as `lhv`.)
  q.setAttribute('style', (el.getAttribute('style') || '').replace(/;?\s*--odx-picb:[^;]*/g, '').replace(/;?\s*--odx-lhv:\s*[\d.]+(px)?;?\s*$/, ''));
  if (lhv) q.style.setProperty('--odx-lhv', lhv);
  q.style.setProperty('--odx-picb', '0px');
  q.style.padding = '0';
  q.style.border = '0';
  q.style.minHeight = '0';
  let inner = q;
  for (const c of chain) { const n = c.cloneNode(false); inner.append(n); inner = n; }
  inner.textContent = 'x';
  box.append(q);
  const one = q.getBoundingClientRect().height;
  const tall = document.createElement('i');
  tall.style.cssText = 'display:inline-block;width:0;height:400px;vertical-align:baseline';
  q.append(tall);
  const below = q.getBoundingClientRect().height - 400;
  q.remove();
  const v = [one - below, below];
  cache.set(k, v);
  return v;
}

/**
 * For a paragraph with a picture on a line with text (pararules.js odx-pictext): { d, mts: [[img, px]] } — d the px to
 * add under the paragraph (negative: to take off), mts the margins above pictures — or null where the font isn't known.
 * `box`: a copy of the content box to probe in; `lhv`: the paragraph's --odx-lhv as lhfix.js sets it (or '').
 */
export function pictureFix(el, box, lhv) {
  const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.data.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP) });
  const texts = [];
  for (let t = walk.nextNode(); t; t = walk.nextNode()) texts.push(t);
  if (!texts.length) return null;
  const rcs = getComputedStyle(texts[0].parentElement);
  const family = rcs.fontFamily.split(',').map((f) => f.replace(/["']/g, '').trim().toLowerCase()).find((f) => ASCENT[f]);
  if (!family) return null;
  const pcs = getComputedStyle(el);
  const fs = parseFloat(rcs.fontSize);
  const lf = parseFloat(rcs.getPropertyValue('--odx-lh')) || 1.17;
  const asc = ASCENT[family] * fs;
  const desc = (DESCENT[family] ?? Math.max(0, lf - ASCENT[family])) * fs;
  const exact = parseFloat(pcs.getPropertyValue('--odx-runmult')) === 0;
  const atLeast = !exact && pcs.getPropertyValue('--odx-lhlen').trim() === '1';
  const mult = exact || atLeast ? 1 : parseFloat(pcs.getPropertyValue('--odx-mult')) || 1;
  const chain = [];
  for (let n = texts[0].parentElement; n && n !== el; n = n.parentElement) chain.unshift(n);
  const [up, down] = strut(box, el, chain, lhv);
  // The line's length under exact and at-least spacing (px): a text line's, as the browser has it (at least the text's own).
  const len = up + down;
  // The text's lines: the vertical extent of each text box, to find the pictures that share one.
  const boxes = [];
  for (const t of texts) {
    const r = document.createRange();
    r.selectNodeContents(t);
    for (const b of r.getClientRects()) if (b.height) boxes.push(b);
  }
  // Pictures in the line (not floating or set apart), by line: their baseline (the margin box's foot) and height above it.
  const lines = [];
  for (const img of el.querySelectorAll('img')) {
    if (img.classList.contains('ProseMirror-separator') || /position:|float:/.test(img.getAttribute('style') || '')) continue;
    const ics = getComputedStyle(img);
    if (ics.display === 'block' || ics.position !== 'static' || ics.float !== 'none') continue;
    const fx = parseFloat(ics.getPropertyValue('--odx-fx')) || 0;
    const fxt = parseFloat(ics.getPropertyValue('--odx-fxt')) || 0;
    const h = (parseFloat(ics.height) || img.offsetHeight) + fx + fxt;
    // (Its box as laid out: a turned picture's rect is its turned outline.)
    const base = layoutRect(img).bottom + fx;
    if (!boxes.some((b) => b.top < base && b.bottom > base - 1)) continue; // (not on a line with text)
    const line = lines.find((l) => Math.abs(l.base - base) < 1);
    if (line) { line.imgs.push(img); line.h = Math.max(line.h, h); }
    else lines.push({ base, h, imgs: [img] });
  }
  let d = 0;
  const mts = [];
  for (const { h, imgs } of lines) {
    if (exact) {
      // The line stays its length: the picture's top goes into the line above (by a margin as far up as it's tall over the text).
      if (h > up) for (const img of imgs) mts.push([img, -(h - up)]);
      continue;
    }
    if (atLeast) {
      if (h <= up && h + desc <= len) continue; // (a line of text, in Word and here)
      const top = Math.max(len, h + desc) - desc; // Word's: the baseline the descent above the line's foot
      if (top > h) for (const img of imgs) mts.push([img, top - h]);
      d += top + desc - (Math.max(top, up) + down);
    } else {
      if (h <= up && h <= asc) continue; // (no taller than the text: a line of text)
      d += Math.max(h, asc) + desc + (mult - 1) * lf * fs - (Math.max(h, up) + down);
    }
  }
  return { d: Math.round(d * 1000) / 1000, mts: mts.map(([img, v]) => [img, Math.round(v * 1000) / 1000]) };
}

/** A picture's height above its line's foot (px, effect extents included), or 0 for one not in the line. */
function standing(img) {
  if (img.classList.contains('ProseMirror-separator') || /position:|float:/.test(img.getAttribute('style') || '')) return 0;
  const ics = getComputedStyle(img);
  if (ics.display === 'block' || ics.position !== 'static' || ics.float !== 'none') return 0;
  return (parseFloat(ics.height) || img.offsetHeight) + (parseFloat(ics.getPropertyValue('--odx-fx')) || 0) + (parseFloat(ics.getPropertyValue('--odx-fxt')) || 0);
}

/**
 * For a paragraph of one line holding only pictures (no text): { d, mts } as pictureFix gives, or null. Word's line is the
 * picture and the extra of "multiple" spacing, but no lower than a line of the paragraph mark's font (its picture then on
 * that line's baseline); at least the length under at-least spacing; the length under exact spacing, the picture cut to
 * it (..\picmatrix-word picalone: 9pt pictures 0.8pt short at single spacing and 5.7pt at 1.5, exact lines 11-31pt off).
 */
export function pictureAlone(el, box) {
  // (Not a horizontal rule's line: a picture set on the baseline, its line the paragraph's own, editor.css; doc 25's
  // rules 2px tall were raised to the font's ascent, 0.3px a rule over its 128 pages.)
  // (Nor one with spaces beside the picture: Word's is then a line of text, the picture on its baseline: CnESymbols' page 5.)
  if (el.textContent.replace(/\u200b/g, '') || el.querySelector('img[style*="vertical-align"]')) return null;
  const imgs = [...el.querySelectorAll('img')].filter((i) => standing(i) > 0);
  if (!imgs.length || el.querySelector('br:not(.ProseMirror-trailingBreak)')) return null;
  const bottoms = imgs.map((i) => layoutRect(i).bottom);
  if (Math.max(...bottoms) - Math.min(...bottoms) > 1) return null; // (on one line)
  const pcs = getComputedStyle(el);
  const family = pcs.fontFamily.split(',').map((f) => f.replace(/["']/g, '').trim().toLowerCase()).find((f) => ASCENT[f]);
  if (!family) return null;
  const fs = parseFloat(pcs.fontSize);
  const lf = parseFloat(pcs.getPropertyValue('--odx-lh')) || 1.17;
  const exact = parseFloat(pcs.getPropertyValue('--odx-runmult')) === 0;
  const atLeast = !exact && pcs.getPropertyValue('--odx-lhlen').trim() === '1';
  const mult = exact || atLeast ? 1 : parseFloat(pcs.getPropertyValue('--odx-mult')) || 1;
  // (The length as written, in pt or px: the paragraph's own line height is 0 in quirks mode.)
  const px = (v) => { const m = /^([\d.]+)(pt|px)$/.exec(v.trim()); return m ? Number(m[1]) * (m[2] === 'pt' ? 4 / 3 : 1) : 0; };
  const len = px(pcs.getPropertyValue(atLeast ? '--odx-atleast' : '--odx-lhv')) || parseFloat(lineHeightOf(pcs)) || 0;
  const h = Math.max(...imgs.map(standing));
  let word;
  let mt = 0;
  if (exact) { word = len; if (h > len) mt = len - h; }
  else if (atLeast) { word = Math.max(len, h); mt = word - h; }
  else { word = Math.max(h, lf * fs) + (mult - 1) * lf * fs; if (h < ASCENT[family] * fs) mt = ASCENT[family] * fs - h; }
  mt = Math.round(mt * 1000) / 1000;
  // The browser's line, with those margins: measured on a copy (less this module's own parts).
  const q = el.cloneNode(true);
  q.removeAttribute('data-odx-lh');
  q.setAttribute('style', (el.getAttribute('style') || '').replace(/;?\s*--odx-picb:[^;]*/g, ''));
  q.style.setProperty('--odx-picb', '0px');
  q.style.padding = '0';
  q.style.border = '0';
  q.style.minHeight = '0'; // (the line itself: the paragraph's floor, 1.15em, stays on the paragraph)
  box.append(q);
  for (const i of q.querySelectorAll('img')) {
    i.style.removeProperty('--odx-picmt');
    i.style.removeProperty('--odx-picclip');
    if (mt && standing(i)) i.style.setProperty('--odx-picmt', `${mt}px`);
  }
  const have = q.getBoundingClientRect().height;
  q.remove();
  return { d: Math.round((word - have) * 1000) / 1000, mts: mt ? imgs.map((i) => [i, mt]) : [] };
}
