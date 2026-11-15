// Line numbers (a section's w:lnNumType), drawn in the left margin as Word draws them: each line of the body's
// paragraphs counted (not tables, notes, headers or text boxes; not a paragraph with w:suppressLineNumbers, nor an
// empty one that only ends a section), every countBy-th number shown, right-aligned w:distance left of the text
// (0.25in when auto), on the line's baseline; restarting each page, each section, or going on from the section before
// (continuous), from w:start + 1. Drawn after each layout pass into a layer of the page, so they print and go into PDF
// export, and never move the text. (All checked with Word's PDF of test/fixtures/line-numbers.docx.) In a section of
// several columns, left of each column's text (right of it in a right-to-left section), a paragraph's lines carried
// to the next column numbered there, in reading order (checked with Word's prints, odx-test/mklinenumcols.ps1).
import { sectionXmls, sectionEnd } from './sections.js';

/** A section's line numbering, or null: { countBy, start, distance (px), restart, rtl }. */
export function lineNumbering(sectPr) {
  const el = /<w:lnNumType\b[^>]*\/?>/.exec(sectPr || '')?.[0];
  if (!el) return null;
  const a = (n) => new RegExp(`w:${n}="([^"]*)"`).exec(el)?.[1];
  const countBy = Number(a('countBy') || 0);
  if (!(countBy > 0)) return null; // (Word numbers lines only with a count to show them by)
  // (Auto: 0.25in, or 0.125in in a section of columns: Word's PDF of two columns, the generated documents.)
  const cols = Number(/<w:cols\b[^>]*\sw:num="(\d+)"/.exec(sectPr || '')?.[1] || 1);
  const distance = a('distance') && a('distance') !== 'auto' ? Number(a('distance')) / 15 : cols > 1 ? 12 : 24;
  // (A right-to-left section numbers its lines right of each column: Word's print of two RTL columns.)
  const rtl = /<w:bidi(?=[\s/>])(?![^>]*w:val="(?:0|false|off)")/.test(sectPr || '');
  return { countBy, start: Number(a('start') || 0), distance, restart: a('restart') || 'newPage', rtl };
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
 * measured from the top of `view.dom`. `carriedOf(el)` is a paragraph's parts carried on to the next columns, as drawn
 * there ([{ copy, cut, stop, rewrap, col }]: their lines are numbered there, where they're seen, not where they're
 * clipped off).
 * `columnOf(el)` is a paragraph's column in a section of columns ({ x: from the text's left edge, w }): the numbers go
 * by the column's text edges (a paragraph's own box there can be narrowed on either side).
 */
export function drawLineNumbers(view, layer, pageOf, carriedOf = () => null, columnOf = () => null) {
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
      // The column's text edge (on the right in a right-to-left section): its column's, or the paragraph's box less its
      // start indent.
      const pr = el.getBoundingClientRect();
      const ecs = getComputedStyle(el);
      const indent = parseFloat(rule.rtl ? ecs.marginRight : ecs.marginLeft) || 0;
      const lines = [];
      const parts = carriedOf(el) || [];
      const carry = parts[0];
      const colEdge = (c) => (c ? (rule.rtl ? left + c.x + c.w : left + c.x) : null);
      // (A hanging indent's first line starts left of the box: the column's edge is still the box's less its indent.)
      const edge = colEdge(columnOf(el)) ?? (rule.rtl ? pr.right + indent : pr.left - indent);
      for (const [baseline, , top] of lineBaselines(el, range)) {
        if (carry && top >= pr.top + carry.cut - 1) continue; // clipped off: numbered on the copy below
        lines.push([baseline, edge]);
      }
      for (const part of parts) {
        // The carried lines, at each next column's top: a clipped copy of the whole paragraph (its lines from the cut
        // down to the next part's) or the text from the cut laid out again at that column's width (all of it).
        // (Either copy sits where the paragraph's box would in that column, its indents those of the paragraph.)
        const cr = part.copy.getBoundingClientRect();
        const cEdge = colEdge(part.col) ?? (rule.rtl ? cr.right + indent : cr.left - indent);
        for (const [baseline, , top] of lineBaselines(part.copy, range)) {
          if (!part.rewrap && (top < cr.top + part.cut - 1 || (part.stop != null && top >= cr.top + part.stop - 1))) continue;
          lines.push([baseline, cEdge]);
        }
      }
      for (const [baseline, edge] of lines) {
        const at = !rule.rtl && Math.abs(edge - left) < 2 ? left : edge;
        const p = pageOf(baseline - flowTop);
        if (p !== page && rule.restart === 'newPage') n = rule.start;
        page = p;
        n++;
        if (n % rule.countBy) continue;
        const s = document.createElement('span');
        s.className = rule.rtl ? 'odx-lnnum odx-lnnum-rtl' : 'odx-lnnum';
        s.textContent = String(n);
        s.style.cssText = `left:${(rule.rtl ? at + rule.distance : at - rule.distance) - origin.left}px;top:${baseline - numAscent - origin.top}px`;
        frag.append(s);
      }
    }
    if (ends) {
      sect++;
      const next = rules[sect];
      // A section restarting each section starts again; one restarting each page goes on counting down the page it
      // starts on (Word, after a continuous break: its first line follows the section before's last), its new page
      // starting again as any page does.
      if (next && next.restart === 'newSection') { n = next.start; page = -1; }
      else if (next && !rules[sect - 1]) n = next.start;
    }
  });
  layer.replaceChildren(frag);
}

/** A paragraph's lines: [baseline, left, top] (client px), from its text's boxes, a line where they move down. */
function lineBaselines(p, range) {
  const out = [];
  let lineBottom = -Infinity;
  const boxes = []; // [rect, style]
  const walk = document.createTreeWalker(p, NodeFilter.SHOW_TEXT, {
    acceptNode: (t) => (t.parentElement.closest('[contenteditable="false"]') && t.parentElement.closest('[contenteditable="false"]') !== p ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  for (let t = walk.nextNode(); t; t = walk.nextNode()) {
    if (!t.data.length) continue;
    range.selectNodeContents(t);
    const cs = getComputedStyle(t.parentElement);
    for (const r of range.getClientRects()) if (r.height) boxes.push([r, cs]);
  }
  // (A line holding only a break, such as a column break at the paragraph's start, is a line too, as Word counts it.)
  for (const br of p.querySelectorAll('br:not(.ProseMirror-trailingBreak)')) {
    const r = br.getClientRects()[0];
    if (r && r.height) boxes.push([r, getComputedStyle(br.parentElement)]);
  }
  boxes.sort((a, b) => a[0].top - b[0].top);
  for (const [r, cs] of boxes) {
    // (A box starting at or below the line's foot, 2px allowed for runs in other sizes, starts a new line.)
    if (r.top >= lineBottom - 2) { out.push([r.top + ascent(cs), r.left, r.top]); lineBottom = r.bottom; }
    else lineBottom = Math.max(lineBottom, r.bottom);
  }
  // An empty paragraph is a line too.
  if (!out.length) {
    const r = p.getClientRects()[0];
    const top = r && r.top + (parseFloat(getComputedStyle(p).paddingTop) || 0);
    if (r) out.push([top + ascent(getComputedStyle(p)), r.left, top]);
  }
  return out;
}
