// Tab stops: each tab is sized after layout so the text after it lands on the
// paragraph's next stop (left, right, center or decimal), with dot, hyphen or
// underscore leaders, falling back to the document's default tab grid.
import { mergeTabs } from '../ooxml/props.js';

const PX = 96 / 1440; // twips to CSS px

const LEADER = {
  dot: 'radial-gradient(circle, currentColor 0.6px, transparent 1px) left bottom 0.2em / 0.3em 0.3em repeat-x',
  middleDot: 'radial-gradient(circle, currentColor 0.6px, transparent 1px) left center / 0.3em 0.3em repeat-x',
  hyphen: 'linear-gradient(90deg, currentColor 60%, transparent 60%) left 60% / 0.4em 1px repeat-x',
  underscore: 'linear-gradient(currentColor, currentColor) left bottom 0.1em / 100% 1px no-repeat',
  heavy: 'linear-gradient(currentColor, currentColor) left bottom 0.1em / 100% 2px no-repeat',
};

/** Width of the text from `start` (after a tab) up to the next tab or the paragraph end; up to '.' for decimal stops. */
function runWidth(p, tab, decimal) {
  const r = document.createRange();
  r.setStartAfter(tab);
  const next = [...p.querySelectorAll('.odx-tab')].find((t) => tab.compareDocumentPosition(t) & Node.DOCUMENT_POSITION_FOLLOWING);
  if (next) r.setEndBefore(next);
  else r.setEnd(p, p.childNodes.length);
  if (decimal) {
    const w = document.createTreeWalker(r.commonAncestorContainer, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      if (!r.intersectsNode(n) || !(tab.compareDocumentPosition(n) & Node.DOCUMENT_POSITION_FOLLOWING)) continue;
      const i = n.data.indexOf('.');
      if (i >= 0) { r.setEnd(n, i); break; }
    }
  }
  // (Empty boxes left out: a field's hidden code can sit at the end of a line the text after it wrapped from.)
  const rects = [...r.getClientRects()].filter((x) => x.width > 0);
  if (!rects.length) return 0;
  const top = rects[0].top;
  const line = rects.filter((x) => Math.abs(x.top - top) < 2);
  return Math.max(...line.map((x) => x.right)) - Math.min(...line.map((x) => x.left));
}

export function layoutTabs(view, session) {
  layoutTabsIn(view.dom, session, (p) => view.state.doc.resolve(view.posAtDOM(p, 0)).parent);
}

/** Sizes the tabs of the paragraphs under `root`; `nodeOf(p)` gives a paragraph's node (its stops), or throws. */
export function layoutTabsIn(root, session, nodeOf) {
  const spans = root.querySelectorAll('.odx-tab');
  if (!spans.length) return;
  const byPara = new Map();
  for (const s of spans) {
    const p = s.closest('p');
    if (!p) continue;
    if (!byPara.has(p)) byPara.set(p, []);
    byPara.get(p).push(s);
    s.style.cssText = 'display:inline-block;tab-size:0;width:0';
  }
  const grid = (session.defaultTabStop || 720) * PX;
  for (const [p, tabs] of byPara) {
    let stops = [];
    try {
      const node = nodeOf(p);
      stops = mergeTabs(session.styleProp(node.attrs.styleId, 'tabs'), node.attrs.view?.tabs)
        .filter((t) => t.val !== 'bar').map((t) => ({ ...t, x: t.pos * PX }));
    } catch { /* not a document paragraph */ }
    const cs = getComputedStyle(p);
    const pr = p.getBoundingClientRect();
    // Stops are measured from the column's start edge: the paragraph box less its start indent. In a right-to-left
    // paragraph that is the right edge, and positions run leftwards.
    const rtl = cs.direction === 'rtl';
    const startMargin = parseFloat(rtl ? cs.marginRight : cs.marginLeft) || 0;
    const edge = rtl ? pr.right + startMargin : pr.left - startMargin;
    const from = (r) => (rtl ? edge - r.right : r.left - edge); // a box's start, from the edge
    // A hanging indent is a stop too: a tab on the first line, before the indent, goes to it first.
    const indent = (parseFloat(cs.textIndent) || 0) < 0 ? startMargin : null;
    for (const s of tabs) {
      const x = from(s.getBoundingClientRect());
      let stop = stops.find((t) => t.x > x + 0.5);
      if (indent != null && indent > x + 0.5 && (!stop || stop.x > indent)) stop = { val: 'left', x: indent };
      let w;
      // (The line's end, from the edge: "right" in a left-to-right paragraph.)
      const right = rtl ? edge - (pr.left + (parseFloat(cs.paddingLeft) || 0)) : pr.right - (parseFloat(cs.paddingRight) || 0) - edge;
      // A tab with no stop left that starts at the right margin goes to the next line in Word (sample S097's
      // footer: "Revision Date: …" ends at the right stop, and the tab after it makes an empty line).
      const wraps = !stop && x >= right - 1;
      if (wraps) w = right - x + 1;
      else if (!stop) w = (Math.floor(x / grid) + 1) * grid - x;
      else if (stop.val === 'right' || stop.val === 'end' || stop.val === 'decimal' || stop.val === 'center') {
        const rw = runWidth(p, s, stop.val === 'decimal');
        const end = Math.min(stop.x, right);
        const used = stop.val === 'center' ? rw / 2 : rw;
        w = end - x - used - 0.5; // sub-pixel slack, or a stop at the margin wraps
        // Text that no longer fits before the stop takes its tab to the next line, as Word lays out a long TOC entry:
        // the line's text stays, and the tab starts the next line at the paragraph's indent with its leader up to
        // the right-aligned text (a tab wider than what is left of the line wraps before it).
        if (w < 0 && x + rw > right + 0.5) w = end - startMargin - used - 0.5;
      } else w = stop.x - x;
      // Dot, middle-dot and hyphen leaders are characters, as Word draws them (editor.css); rules stay lines.
      const chars = stop?.leader && ['dot', 'middleDot', 'hyphen'].includes(stop.leader);
      const lead = !chars && stop?.leader && LEADER[stop.leader];
      if (chars) s.dataset.leader = stop.leader; else delete s.dataset.leader;
      s.style.cssText = `display:inline-block;tab-size:0;width:${Math.max(0, w)}px${lead ? `;background:${lead}` : ''}${wraps ? ';white-space:pre' : ''}`; // (its tab character gives that line its height)
    }
  }
}
