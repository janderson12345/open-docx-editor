// Line numbers (a section's w:lnNumType), drawn in the left margin as Word draws them: each line of the body's
// paragraphs counted (not tables, notes, headers or text boxes; not a paragraph with w:suppressLineNumbers, nor an
// empty one that only ends a section), every countBy-th number shown, right-aligned w:distance left of the text
// (0.25in when auto), on the line's baseline; restarting each page, each section, or going on from the section before
// (continuous), from w:start + 1. Drawn after each layout pass into a layer of the page, so they print and go into PDF
// export, and never move the text. (All checked with Word's PDF of test/fixtures/line-numbers.docx.) Not yet in a
// section of several columns: there the numbers sit left of the page's text, not of each column.
import { sectionXmls, sectionEnd } from './sections.js';

/** A section's line numbering, or null: { countBy, start, distance (px), restart }. */
export function lineNumbering(sectPr) {
  const el = /<w:lnNumType\b[^>]*\/?>/.exec(sectPr || '')?.[0];
  if (!el) return null;
  const a = (n) => new RegExp(`w:${n}="([^"]*)"`).exec(el)?.[1];
  const countBy = Number(a('countBy') || 0);
  if (!(countBy > 0)) return null; // (Word numbers lines only with a count to show them by)
  const distance = a('distance') && a('distance') !== 'auto' ? Number(a('distance')) / 15 : 24;
  return { countBy, start: Number(a('start') || 0), distance, restart: a('restart') || 'newPage' };
}

const ascents = new Map();
let ctx = null;
/** The font's ascent above the baseline (CSS px), as the browser lays it out. */
function ascent(cs) {
  const font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
  if (!ascents.has(font)) {
    ctx ||= document.createElement('canvas').getContext('2d');
    ctx.font = font;
    ascents.set(font, ctx.measureText('1').fontBoundingBoxAscent ?? parseFloat(cs.fontSize) * 0.8);
  }
  return ascents.get(font);
}

/**
 * Draws the numbers into `layer` (a child of the page element). `pageOf(y)` is the page at a height in the flow,
 * measured from the top of `view.dom`.
 */
export function drawLineNumbers(view, layer, pageOf) {
  const doc = view.state.doc;
  const rules = sectionXmls(doc).map(lineNumbering);
  if (!rules.some(Boolean)) { layer.replaceChildren(); return; }
  // In the document's own font and size (Word's Line Number style is the default paragraph font's).
  const docCs = getComputedStyle(view.dom);
  layer.style.fontFamily = docCs.fontFamily;
  layer.style.fontSize = docCs.fontSize;
  const origin = layer.getBoundingClientRect();
  const flowTop = view.dom.getBoundingClientRect().top;
  const left = view.dom.getBoundingClientRect().left + (parseFloat(getComputedStyle(view.dom).paddingLeft) || 0);
  const numCs = getComputedStyle(layer);
  const numAscent = ascent(numCs);
  const range = document.createRange();
  const frag = document.createDocumentFragment();
  let sect = 0, n = rules[0]?.start || 0, page = -1; // (the first section counts from its w:start + 1 too)
  doc.forEach((node, pos) => {
    const rule = rules[sect];
    const ends = sectionEnd(node) != null;
    const el = view.nodeDOM(pos);
    const suppressed = (node.attrs.extra || []).some((e) => e[0] === 'suppressLineNumbers' && !/w:val="(?:0|false|off)"/.test(e[1]));
    // (Nor the empty paragraph that only ends a section: Word doesn't count it.)
    if (rule && node.type.name === 'paragraph' && el && el.nodeType === 1 && !suppressed && !(ends && !node.childCount) && !el.classList.contains('odx-all-del')) {
      for (const baseline of lineBaselines(el, range)) {
        const p = pageOf(baseline - flowTop);
        if (p !== page && rule.restart === 'newPage') n = rule.start;
        page = p;
        n++;
        if (n % rule.countBy) continue;
        const s = document.createElement('span');
        s.className = 'odx-lnnum';
        s.textContent = String(n);
        s.style.cssText = `left:${left - rule.distance - origin.left}px;top:${baseline - numAscent - origin.top}px`;
        frag.append(s);
      }
    }
    if (ends) {
      sect++;
      const next = rules[sect];
      if (next && next.restart !== 'continuous') { n = next.start; page = -1; }
      else if (next && !rules[sect - 1]) n = next.start;
    }
  });
  layer.replaceChildren(frag);
}

/** The baselines (client y) of a paragraph's lines: from its text's boxes, a line where they move down. */
function lineBaselines(p, range) {
  const out = [];
  let lineBottom = -Infinity;
  const walk = document.createTreeWalker(p, NodeFilter.SHOW_TEXT, {
    acceptNode: (t) => (t.parentElement.closest('[contenteditable="false"]') && t.parentElement.closest('[contenteditable="false"]') !== p ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  for (let t = walk.nextNode(); t; t = walk.nextNode()) {
    if (!t.data.length) continue;
    range.selectNodeContents(t);
    const cs = getComputedStyle(t.parentElement);
    for (const r of range.getClientRects()) {
      if (!r.height) continue;
      // (A box starting at or below the line's foot, 2px allowed for runs in other sizes, starts a new line.)
      if (r.top >= lineBottom - 2) { out.push(r.top + ascent(cs)); lineBottom = r.bottom; }
      else lineBottom = Math.max(lineBottom, r.bottom);
    }
  }
  // An empty paragraph is a line too.
  if (!out.length) {
    const r = p.getClientRects()[0];
    if (r) out.push(r.top + (parseFloat(getComputedStyle(p).paddingTop) || 0) + ascent(getComputedStyle(p)));
  }
  return out;
}
