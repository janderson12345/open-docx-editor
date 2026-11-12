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

const sized = new WeakMap(); // paragraph element -> what its tabs were sized for
// (Sized again once a font loads: a web font arriving after the first layout, as Carlito does where Calibri isn't
// installed, changes every width measured in its stand-in.)
let fontLoads = 0;
if (typeof document !== 'undefined' && document.fonts) document.fonts.addEventListener('loadingdone', () => { fontLoads++; });

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
  }
  // A paragraph whose text, formatting and place are as when its tabs were sized keeps them (sizing every tab again
  // each pass restyled them all: a keystroke in a long document laid out each of its tabbed lines again).
  // (Where its tabs and its text's pieces sit is part of it, as they were once sized: the layout corrects a hinted
  // font's widths after the tabs, in the text's own spans, and a stop measured before that lands short.)
  const sigOf = (p, tabs) => `${fontLoads}|${p.textContent}|${p.getAttribute('style')}|${p.className}|${p.offsetWidth}|${p.offsetLeft}|${tabs.map((t) => t.offsetLeft).join(',')}|${[...p.querySelectorAll('.odx-t')].map((t) => t.offsetWidth).join(',')}`;
  const todo = [...byPara].filter(([p, tabs]) => sized.get(p) !== sigOf(p, tabs));
  for (const [, tabs] of todo) for (const s of tabs) s.style.cssText = 'display:inline-block;tab-size:0;width:0';
  const grid = (session.defaultTabStop || 720) * PX;
  for (const [p, tabs] of todo) {
    let stops = [];
    try {
      const node = nodeOf(p);
      stops = mergeTabs(session.styleProp(node.attrs.styleId, 'tabs'), node.attrs.view?.tabs)
        .filter((t) => t.val !== 'bar').map((t) => ({ ...t, x: t.pos * PX }));
    } catch { /* not a document paragraph */ }
    const cs = getComputedStyle(p);
    // Stops are measured from the column's start edge: the paragraph box less its start indent. In a right-to-left
    // paragraph that is the right edge, and positions run leftwards.
    const rtl = cs.direction === 'rtl';
    const startMargin = parseFloat(rtl ? cs.marginRight : cs.marginLeft) || 0;
    // A hanging indent is a stop too: a tab on the first line, before the indent, goes to it first.
    const indent = (parseFloat(cs.textIndent) || 0) < 0 ? startMargin : null;
    let aligned = false; // the tab before was a right, centre or decimal one
    for (const s of tabs) {
      // (The paragraph's box read again for each tab: sizing the one before can move the page, as a scrollbar
      // appearing does, and a stop measured from where the paragraph was would land short.)
      const pr = p.getBoundingClientRect();
      const edge = rtl ? pr.right + startMargin : pr.left - startMargin;
      const from = (r) => (rtl ? edge - r.right : r.left - edge); // a box's start, from the edge
      const x = from(s.getBoundingClientRect());
      let stop = stops.find((t) => t.x > x + 0.5);
      // (Text ending a fraction of a pixel past the last stop, a right one, is still before it in Word's units: a long
      // table of contents title that fills its line, in a font drawn a shade wider than Word's, with its page number
      // after the tab. A tab with nothing after it there starts an empty line, as below.)
      const last = stops[stops.length - 1];
      if (!stop && last && (last.val === 'right' || last.val === 'end') && x - last.x < 1 && runWidth(p, s, false) > 0) stop = last;
      if (indent != null && indent > x + 0.5 && (!stop || stop.x > indent)) stop = { val: 'left', x: indent };
      let w;
      // (The line's end, from the edge: "right" in a left-to-right paragraph.)
      const right = rtl ? edge - (pr.left + (parseFloat(cs.paddingLeft) || 0)) : pr.right - (parseFloat(cs.paddingRight) || 0) - edge;
      // A tab with no stop left that starts at the right margin goes to the next line in Word (sample S097's
      // footer: "Revision Date: …" ends at the right stop, and the tab after it makes an empty line).
      const wraps = !stop && x >= right - 1;
      if (wraps) w = right - x + 1;
      // (After text aligned at a stop, which ends the half pixel short of it that the slack below leaves, the next default
      // stop is past that one: a default stop there too gave the next tab no width, where Word's goes on to the next,
      // tools/doc-generator's centre and decimal stops.)
      else if (!stop) {
        const next = (v) => (Math.floor(v / grid) + 1) * grid;
        // (Not at the line's end: a tab after text right-aligned at the right margin stays on the line there, sample
        // Attachment D's two-part headings.)
        w = (aligned && next(x + 0.5) < right - 0.5 ? next(x + 0.5) : next(x)) - x;
      }
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
      aligned = !!stop && ['right', 'end', 'decimal', 'center'].includes(stop.val);
      s.style.cssText = `display:inline-block;tab-size:0;width:${Math.max(0, w)}px${lead ? `;background:${lead}` : ''}${wraps ? ';white-space:pre' : ''}`; // (its tab character gives that line its height)
    }
  }
  for (const [p, tabs] of todo) sized.set(p, sigOf(p, tabs));
}
