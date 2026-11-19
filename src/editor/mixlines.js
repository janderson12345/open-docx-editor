// Lines holding runs of different sizes or fonts, or raised and lowered runs. Word makes such a line as tall as the
// highest run's top above the baseline (its font's ascent, plus its raise) and the lowest run's foot below it (its
// descent, less its raise), each from its own font: max(ascent + raise) + max(descent - raise); "multiple" spacing adds
// (multiple - 1) times the tallest font's single line below the text (Word's PDFs of 136 generated cases: Times New
// Roman, Calibri, Arial and Aptos, 8 to 16pt runs raised and lowered 3 to 8pt in 10 to 12pt text, single, 1.15 and
// double spacing, within 0.02pt). The ascent is the font's Windows ascent plus its external leading, the descent its
// Windows descent (baseline.js wordAscent).
//
// The browser centres each run's leading on its own text instead, so the split above and below differs by font: a 12pt
// run raised 3pt in 10pt Times New Roman made a line 0.48px short, in Calibri 0.47px tall, at double spacing 1.5px off.
// Runs of only spaces or tabs don't count (Word's PDFs: a 16pt tab or space in 10pt text leaves a 10pt line).
//
// So in such a paragraph each text span gets Word's box: a line height of its top plus its foot, and a vertical-align
// that puts that box where Word's is, its text moved back by the same amount (position: relative), so the text stays
// where it was drawn and only the line's extent changes.
import { knownFamily, raiseOf, wordAscent } from './baseline.js';
import { lineFactor } from '../ooxml/props.js';

const tops = new Map(); // font, line height -> the browser's top of a span's box above its baseline (px)
let probe = null;

/** The browser's top (px) of the box of a span in `cs`'s font with line height `lh` (px), above its baseline. */
function boxTop(cs, lh, host) {
  const font = `${cs.fontStyle}|${cs.fontWeight}|${cs.fontSize}|${cs.fontFamily}`;
  const key = `${font}|${lh}`;
  if (tops.has(key)) return tops.get(key);
  if (!probe) {
    probe = document.createElement('div');
    probe.setAttribute('aria-hidden', 'true');
    probe.style.cssText = 'position:absolute;left:-9999px;top:0;visibility:hidden;white-space:nowrap;line-height:0;font-size:0';
    probe.innerHTML = '<span>X<i style="display:inline-block;width:0;height:0"></i></span>';
  }
  host.append(probe);
  const s = probe.firstChild;
  s.style.fontFamily = cs.fontFamily;
  s.style.fontSize = cs.fontSize;
  s.style.fontWeight = cs.fontWeight;
  s.style.fontStyle = cs.fontStyle;
  s.style.lineHeight = `${lh}px`;
  const v = probe.querySelector('i').getBoundingClientRect().top - probe.getBoundingClientRect().top;
  probe.remove();
  if (tops.size > 4000) tops.clear();
  tops.set(key, v);
  return v;
}

/** A span's line height (px) as the stylesheet gives it (--odx-lhv: a number of its font size, or a length), or 0. */
function cssLine(cs, fs) {
  const v = cs.getPropertyValue('--odx-lhv').trim();
  if (v.endsWith('px')) return parseFloat(v) || 0;
  const n = Number(v);
  return Number.isFinite(n) && n > 0 ? n * fs : 0;
}

const MAX_TEXTS = 2000; // (a paragraph of more runs than this is left as it is)

/**
 * A paragraph of nothing but spaces (no text, no tabs or breaks): Word makes its line the paragraph mark's alone, as an
 * empty paragraph's, whatever size the spaces are in (Word: an 18pt space, an 18pt and a 12pt one, or an 18pt lowered
 * one, over a 12pt mark: a 12pt line; a 12pt space over an 18pt mark: an 18pt line). So its spans get no box and the
 * paragraph its own line, as an empty one has (editor.css, data-odx-mix="mark"). undefined: not such a paragraph; null:
 * one already laid out as Word does (all its spaces in the mark's size, none raised or lowered).
 */
function blankOnly(el, pcs) {
  if (el.textContent.trim() || !el.textContent || el.querySelector('.odx-tab, .odx-opaque-inline')) return undefined;
  const pfs = parseFloat(pcs.fontSize) || 0;
  const spans = [];
  let differs = false;
  const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  for (let t = walk.nextNode(); t; t = walk.nextNode()) {
    if (!t.data) continue;
    if (spans.length >= MAX_TEXTS) return null;
    const span = t.parentElement;
    if (span.closest('sup, sub')) differs = true;
    if (!span.classList.contains('odx-t') || span.childNodes.length !== 1) { if (span !== el) differs = true; continue; }
    if (Math.abs(parseFloat(getComputedStyle(span).fontSize) - pfs) > 0.01) differs = true;
    spans.push([span, 'line-height:0']);
  }
  return differs ? { spans, mark: true } : null;
}

/**
 * Word's boxes for the text spans of a paragraph `el` (drawn, in the document) whose runs differ in font, size or raise,
 * where the browser's lines would come out more than 0.05px off Word's: { spans: [[span, css]] }, or null to leave it.
 * Only "auto" spacing of 1 or more (exact and at-least lines aren't sized from their runs this way), and not a paragraph
 * with pictures, equations, breaks, fields drawn as text, a padded list label or a document grid.
 */
export function mixedLines(el, host) {
  if (el.tagName !== 'P' || el.dataset.odxGrid === '1') return null;
  const pcs = getComputedStyle(el);
  const rm = pcs.getPropertyValue('--odx-runmult').trim();
  const mult = rm ? parseFloat(rm) : 1; // (no spacing of its own: single, sample S078's footer)
  if (!(mult >= 1) || pcs.getPropertyValue('--odx-lhlen').trim() === '1') return null;
  if (el.querySelector('img:not(.ProseMirror-separator), math, svg, br:not(.ProseMirror-trailingBreak), .odx-opaque-inline:not(:empty), .odx-strut')) return null;
  if (/--odx-lbl-/.test(el.getAttribute('style') || '')) return null;
  const blankLine = blankOnly(el, pcs);
  if (blankLine !== undefined) return blankLine;
  const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  const groups = new Map();
  const spans = [];
  let n = 0;
  for (let t = walk.nextNode(); t; t = walk.nextNode()) {
    if (!t.data) continue;
    if (++n > MAX_TEXTS) return null;
    const span = t.parentElement;
    if (span.closest('sup, sub')) continue; // (raised and lowered text in Word's smaller size doesn't size its line)
    // A run of nothing but spaces or tabs doesn't size its line in Word (16pt tabs or spaces in a line of 10pt text: a 10pt
    // line; sample S078's footer, its tabs in the Footer style's 12pt between 10pt runs).
    const blank = !t.data.trim();
    if (blank && span.classList.contains('odx-tab')) continue; // (a tab in the body: no box of its own, line-height 0)
    if (!span.classList.contains('odx-t') || span.childNodes.length !== 1) return null;
    const cs = getComputedStyle(span);
    const family = knownFamily(cs);
    if (!family && !blank) return null;
    const fs = parseFloat(cs.fontSize);
    const raise = raiseOf(span, el);
    if (!(fs > 0) || raise == null) return null;
    const key = `${blank ? 'blank|' : ''}${cs.fontFamily}|${cs.fontSize}|${raise}|${cs.fontWeight}|${cs.fontStyle}`;
    let g = groups.get(key);
    if (!g) {
      const lb = cssLine(cs, fs);
      if (!(lb > 0)) return null;
      g = { cs, fs, raise, blank, chars: 0, lb };
      if (!blank) {
        const a = wordAscent(family);
        const f = lineFactor(family);
        if (!(f > a)) return null;
        Object.assign(g, { T: a * fs + raise, B: (f - a) * fs - raise, L: f * fs });
      }
      g.word = `${family}|${cs.fontSize}|${raise}`;
      groups.set(key, g);
    }
    g.chars += t.data.length;
    spans.push([span, g]);
  }
  const all = [...groups.values()];
  const gs = all.filter((g) => !g.blank);
  // (Runs differing only in weight or slant are one size to Word: sample 24's DejaVu Sans paragraphs with bold and
  // italic runs, whose stand-in's bold face the browser gives other metrics, came out 0.75px a line short as mixed.)
  if (!gs.length || (all.length === gs.length && new Set(gs.map((g) => g.word)).size < 2)) return null;
  const max = (k) => Math.max(...gs.map((g) => g[k]));
  const word = max('T') + max('B') + (mult - 1) * max('L');
  // The browser's line as drawn now, all of the runs on one line: each span's box centred on its own text.
  let bt = -Infinity, bb = -Infinity;
  for (const g of all) {
    const top = boxTop(g.cs, g.lb, host);
    bt = Math.max(bt, top + g.raise);
    bb = Math.max(bb, g.lb - top - g.raise);
  }
  if (Math.abs(bt + bb - word) <= 0.05) return null;
  // Under "multiple" spacing the extra goes below the line's lowest foot, from its tallest font: each run's box reaches
  // its own extra below the paragraph's main text's foot (the runs most of its text is in), which the runs' own lines
  // nearly always share, so a line's foot comes out as Word's.
  const base = gs.reduce((x, y) => (y.chars > x.chars ? y : x));
  for (const g of all) {
    // (A blank run's box: nothing, a point on its line.)
    if (g.blank) { g.css = 'line-height:0'; continue; }
    const foot = (mult > 1 ? Math.max(g.B, base.B) : g.B) + (mult - 1) * g.L;
    const lh = Math.round((g.T + foot) * 1000) / 1000;
    if (!(lh > 0)) return null;
    const v = Math.round((g.T - g.raise - boxTop(g.cs, lh, host)) * 1000) / 1000;
    g.css = `line-height:${lh}px;vertical-align:${v}px;position:relative;top:${v}px`;
  }
  return { spans: spans.map(([s, g]) => [s, g.css]) };
}
