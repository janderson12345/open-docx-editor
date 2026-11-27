// Headers and footers as drawn on the sheets: a story's content serialized, sized as Word sizes its lines, its
// tables styled and its fields filled in for the page.
import { DOMSerializer } from 'prosemirror-model';
import { TableMap } from 'prosemirror-tables';
import { schema } from './schema.js';
import { tableCellCss, tablePlacementCss, firstCellBorders } from '../ooxml/styles.js';
import { parseXml } from '../ooxml/xml.js';
import { twipsToPx } from '../ooxml/units.js';
import { fontStack, lineFactor, borderCss, themeFace } from '../ooxml/props.js';
import { deletedWhole, markSize } from './pararules.js';
import { wrapText } from './textspans.js';
import { uniformRun } from './runs.js';
import { docProps, fillFields } from './pagefields.js';
import { previewStory } from './mailmerge.js';

/** Header and footer paragraphs -> their nodes, for laying out their tabs once drawn. */
export const hfNodes = new WeakMap();

/** A story's tables (a header's, a footer's, a note's) given the same placement, column widths, borders and shading the
 *  body's decorations give its tables. */
export function styleStoryTables(session, doc, el, { note = false } = {}) {
  const tables = [];
  doc.descendants((n) => { if (n.type === schema.nodes.table) tables.push(n); });
  const doms = el.querySelectorAll('table');
  tables.forEach((t, k) => {
    const dom = doms[k];
    if (!dom) return;
    const place = tablePlacementCss(session.styles, t.attrs.tblPr, session.compatMode, parseXml, t.attrs.tblGrid, 0, firstCellBorders(t));
    // (A note's tables don't float: Word has no text wrapping in notes, and puts one that says so in the note's flow.)
    if (place && !(note && /(?:^|;)float:/.test(place))) dom.style.cssText += ';' + place;
    // (Floating, placed from the text above its own place: up too, as Word draws it in a header (its PDF: a table at
    // tblpY -200 starts 10pt above the paragraph after it; it had been drawn at that paragraph's top). The body's own
    // floating tables keep theirs at 0 or below, pages.js.)
    const up = !note && /(?:^|;)float:/.test(place) && !/\bw:vertAnchor="page"/.test(t.attrs.tblPr || '') ? /\bw:tblpY="(-\d{1,6})"/.exec(t.attrs.tblPr || '') : null;
    if (up) {
      dom.style.marginTop = ''; // (the body's lift, styles.js: here drawn up only)
      dom.style.top = `${Math.max(-1440, twipsToPx(Number(up[1]))).toFixed(2)}px`;
    }
    const map = TableMap.get(t);
    const cells = [];
    const seen = new Set();
    for (const rel of map.map) {
      if (seen.has(rel)) continue;
      seen.add(rel);
      const node = t.nodeAt(rel);
      if (!node || !node.type.spec.tableRole?.includes('cell')) continue;
      const r = map.findCell(rel);
      cells.push({ node, row: r.top, col: r.left, rowspan: r.bottom - r.top, colspan: r.right - r.left, tcPr: node.attrs.tcPr });
    }
    const res = tableCellCss(session.styles, t.attrs.tblPr, map.height, map.width, cells, parseXml, session.theme, session.compatMode);
    const firstRow = dom.rows[0];
    const ws = firstRow ? [...firstRow.cells].map((td) => (td.getAttribute('data-colwidth') || '').split(',').map(Number)).flat() : [];
    if (ws.length && ws.every((w) => w > 0)) {
      const cg = document.createElement('colgroup');
      for (const w of ws) cg.appendChild(document.createElement('col')).style.width = w + 'px';
      dom.prepend(cg);
      dom.style.width = ws.reduce((a, b) => a + b, 0) + 'px';
      dom.style.tableLayout = 'fixed';
    }
    const tds = [...dom.rows].flatMap((tr) => [...tr.cells]);
    cells.forEach((c, i) => { if (res.css[i] && tds[i]) tds[i].style.cssText += ';' + res.css[i]; });
  });
}

/**
 * A header's or footer's height as Word sets the body clear of it: down to its last block in the flow, through that
 * block's space after. What floats in it (a picture wrapped square or tight, a floating table) reaches below that without
 * moving the body, which wraps round it instead (Word's PDFs: a header's 60pt logo wrapped square beside its paragraph,
 * reaching 26pt into the body's top margin, and a floating table 6 rows tall under an empty paragraph, the body's first
 * line where it is with no header; both had pushed the body down under them).
 */
export function flowHeight(el) {
  const ecs = getComputedStyle(el);
  const top = el.getBoundingClientRect().top + (parseFloat(ecs.borderTopWidth) || 0);
  let bottom = top + (parseFloat(ecs.paddingTop) || 0);
  for (const c of el.children) {
    const cs = getComputedStyle(c);
    if (cs.float !== 'none' || cs.position === 'absolute' || cs.position === 'fixed' || cs.display === 'none') continue;
    bottom = Math.max(bottom, c.getBoundingClientRect().bottom + Math.max(0, parseFloat(cs.marginBottom) || 0));
  }
  return bottom - top + (parseFloat(ecs.paddingBottom) || 0) + (parseFloat(ecs.borderBottomWidth) || 0);
}

/**
 * A drawn header's shapes that the body's text wraps round (reader.js anchorCss, --odx-hfwrap), and a footer's placed
 * below their paragraph: out of the story's flow, where they're drawn. Word wraps none of a header's own text round
 * them, only the body's: the header's lines run on across a picture wrapped square beside them, under it or through it
 * (hdrht-word e8-e13: at the left edge, in the middle, past the right margin, the next paragraphs too; a sample's header
 * line beside its logo past the margin had gone under the logo, 50.7px low, the body with it). As floats they pushed
 * the header's own text aside, and placed below their paragraph their room reached from its top down to them, pushing
 * the header's foot (and the body under it) down past them, where Word keeps both as without the shape (footwrap-word
 * fw-header, a rectangle 250pt below its header paragraph). `el` must be in the document, to be measured.
 */
export function settleStoryShapes(el) {
  placeStoryShapes(el);
  const header = el.classList.contains('odx-header') && el.dataset.odxStory !== 'footer'; // (pages.js storyHeight's footers)
  const below = [];
  for (const img of el.querySelectorAll('img[style*="--odx-hfwrap"]')) {
    const p = img.closest('p');
    if (!p || getComputedStyle(img).float === 'none') continue;
    // (A footer's beside its paragraph stays a float: not checked in Word, and taken out of the flow four samples' footers
    // with a text box or picture wrapped beside their text lost 6 to 166 lines.)
    if (!header && (parseFloat(getComputedStyle(img).marginTop) || 0) < p.getBoundingClientRect().height) continue;
    below.push({ img, at: img.getBoundingClientRect() });
  }
  for (const b of below) { b.img.style.float = 'none'; b.img.style.position = 'absolute'; }
  for (const { img, at } of below) {
    const now = img.getBoundingClientRect();
    img.style.translate = `${(at.left - now.left).toFixed(2)}px ${(at.top - now.top).toFixed(2)}px`;
  }
  if (header) dropBelowLogo(el);
}

/**
 * A header's or footer's positional tabs (w:ptab, schema.js data-ptab): each as wide as puts what follows it, up to the
 * next positional tab or line break, at its place: its left edge at the margin's (or the paragraph indent's) left, its
 * middle at their middle, its right edge at their right, as Word sets them (hfbody-word p10, q*). They had no width at
 * all, the text after them run on from where the text before ended. Laid out in order along each paragraph, after its
 * ordinary tabs; one with no room (what follows runs past its place) takes none. `el` must be in the document.
 */
export function layoutPtabs(el) {
  const tabs = [...el.querySelectorAll('span.odx-opaque-inline[data-kind="ptab"]')];
  if (!tabs.length) return;
  const eb = el.getBoundingClientRect();
  const scale = el.offsetWidth ? eb.width / el.offsetWidth || 1 : 1; // (from its width: its height is whole pixels)
  const box = (e) => {
    const r = e.getBoundingClientRect(), cs = getComputedStyle(e);
    const l = r.left + ((parseFloat(cs.borderLeftWidth) || 0) + (parseFloat(cs.paddingLeft) || 0)) * scale;
    const rr = r.right - ((parseFloat(cs.borderRightWidth) || 0) + (parseFloat(cs.paddingRight) || 0)) * scale;
    return [l, rr];
  };
  for (const t of tabs) {
    t.style.display = 'inline-block'; t.style.width = '0px';
    // (A line with one goes from the left in Word, whatever the paragraph's alignment: q5, centred, "Left words" at the
    // margin and "Right words" at the right one.)
    const p = t.closest('p');
    if (p) p.style.textAlign = 'left';
  }
  for (const t of tabs) {
    const p = t.closest('p');
    if (!p) continue;
    // (Against the paragraph's indents, against the margins or not: q8, an indented paragraph's tab "relative to the
    // margin" ends its text at the right indent, as q4's relative to the indent.)
    const [l, r] = box(p);
    // What follows it on its line: up to the next positional tab or line break, or the paragraph's end.
    const range = document.createRange();
    range.setStartAfter(t);
    let end = null;
    const walk = document.createTreeWalker(p, NodeFilter.SHOW_ELEMENT);
    walk.currentNode = t;
    for (let n = walk.nextNode(); n; n = walk.nextNode()) {
      if (t.contains(n)) continue;
      if (n.tagName === 'BR' || n.matches('span[data-kind="ptab"], .odx-hfdrop')) { end = n; break; }
    }
    if (end) range.setEndBefore(end); else range.setEnd(p, p.childNodes.length);
    const rects = [...range.getClientRects()].filter((q) => q.width > 0);
    let at = t.getBoundingClientRect().left;
    const w = rects.length ? Math.max(...rects.map((q) => q.right)) - Math.min(at, ...rects.map((q) => q.left)) : 0;
    const align = t.dataset.ptab;
    // (A left one the text has already passed: what follows goes on at the next line's start, q7.)
    if (align === 'left') {
      if (l < at - 0.5) { t.before(document.createElement('br')); at = t.getBoundingClientRect().left; }
      t.style.width = `${Math.max(0, (l - at) / scale).toFixed(2)}px`;
      continue;
    }
    // Right or centred: the tab reaches the place, and what follows it, gathered in a box of no width, is drawn back by
    // its own width (or half): it stays there whatever width its text comes to once its fonts and Word's widths are in
    // (a width measured here had come out 2px short of the text's, its last word wrapped to the next line). Only where
    // it fits before the place, and its parts lie side by side in the paragraph; else the tab takes the room it leaves.
    const goal = align === 'right' ? r : (l + r) / 2;
    // (Laid out before: its box is there already.)
    if (t.nextSibling?.classList?.contains('odx-ptseg')) { t.style.width = `${Math.max(0, (goal - at) / scale - 0.5).toFixed(2)}px`; continue; }
    const parts = [];
    for (let n = t.nextSibling; n && n !== end; n = n.nextSibling) parts.push(n);
    const flat = t.parentElement === p && (!end || end.parentElement === p);
    const fits = at + (align === 'right' ? w : w / 2) <= goal + 4; // (give or take the widths still to come)
    if (flat && fits && parts.length) {
      const seg = document.createElement('span');
      seg.className = 'odx-ptseg';
      seg.style.cssText = 'display:inline-block;width:0;white-space:pre';
      const inner = document.createElement('span');
      inner.style.cssText = `display:inline-block;translate:calc(${align === 'right' ? '-100%' : '-50%'} + 0.5px) 0`;
      t.after(seg);
      seg.append(inner);
      inner.append(...parts);
      // (Half a pixel short, made up by the box's shift: at the very edge it had gone to the next line.)
      t.style.width = `${Math.max(0, (goal - at) / scale - 0.5).toFixed(2)}px`;
    } else t.style.width = `${Math.max(0, (goal - (align === 'right' ? w : w / 2) - at) / scale - 0.5).toFixed(2)}px`; // (nothing after it: up to just short of the margin, a line of its own otherwise)
  }
}

/**
 * A header line of text, a right positional tab (w:ptab), then a logo placed against the page and wrapped round (square,
 * tight or through) reaching across that line, then more positional tabs: Word sets what follows the logo below it, one
 * line for what comes before the next positional tab and one for each positional tab after (hfbody-word p1, p5, p6, p9,
 * p10-p12: wrapped square at the logo's foot, tight 12px lower, --odx-lbelow; the body under those lines). Without the tab
 * before or after the logo, with a logo wrapped in front or behind (wrapNone) or one starting below the line, or with
 * plain tabs, the line stays as it is (p2-p4, p7, p8, p13). Two samples' headers had their body 59px high on every page.
 * `el` must be in the document at its place on the page (its offsetTop: the header's distance from the page's top).
 */
function dropBelowLogo(el) {
  const imgs = [...el.querySelectorAll('p > img[style*="--odx-lbelow"]')];
  if (!imgs.length) return;
  const eb = el.getBoundingClientRect();
  const scale = el.offsetWidth ? eb.width / el.offsetWidth || 1 : 1; // (from its width: its height is whole pixels)
  for (const img of imgs) {
    const p = img.parentElement;
    const st = img.style;
    if (!['page', 'topMargin'].includes(st.getPropertyValue('--odx-lrel').trim()) || st.getPropertyValue('--odx-lyalign').trim() !== 'none') continue;
    const tabs = [...p.querySelectorAll('span[data-kind="ptab"]')];
    const after = tabs.filter((t) => img.compareDocumentPosition(t) & Node.DOCUMENT_POSITION_FOLLOWING);
    if (after.length === tabs.length || !after.length) continue;
    const top = (parseFloat(st.getPropertyValue('--odx-ly')) || 0) - el.offsetTop;
    const foot = top + (parseFloat(st.height) || img.offsetHeight) + (parseFloat(st.getPropertyValue('--odx-lbelow')) || 0);
    const drop = document.createElement('span');
    drop.className = 'odx-hfdrop';
    drop.style.cssText = 'display:block;height:0';
    img.after(drop);
    const line = (drop.getBoundingClientRect().top - eb.top) / scale; // (the first line's foot)
    const first = (p.getBoundingClientRect().top - eb.top) / scale;
    if (!(top < line && foot > first && foot > line)) { drop.remove(); continue; }
    drop.style.height = `${(foot - line).toFixed(2)}px`;
    // (Each of those lines a full line of the paragraph's font, empty or not.)
    // (In the paragraph style's size, sizeStoryLines' --odx-ssz, not the text's given to the paragraph: Word's lines of
    // spaces and positional tabs with no size of their own under a 12pt first line are 11pt lines, p1, p5.)
    const ssz = getComputedStyle(p).getPropertyValue('--odx-ssz').trim();
    const strut = () => { const s = document.createElement('span'); s.className = 'odx-t odx-hfstrut'; s.textContent = '​'; if (ssz) s.style.fontSize = ssz; return s; };
    drop.after(strut());
    for (const t of after) t.before(document.createElement('br'), strut());
    // (The positional tab before the logo places what follows the logo, on the first line under it: p10, "MID" set
    // right at the margin there. Nothing it places is left on the line above.)
    const lead = tabs.filter((t) => !after.includes(t)).pop();
    const between = lead && document.createRange();
    if (between) { between.setStartAfter(lead); between.setEndBefore(img); }
    if (lead && !between.toString().trim()) drop.nextSibling.after(lead);
  }
}

/**
 * A header's or footer's shapes placed against the page or its margins (reader.js anchorCss: --odx-lrel wrapped round,
 * --odx-arel not), set there out of the story's flow, as Word draws them: from the page's top, or from the margin's,
 * which is where the body starts on that page (pushed down by a header taller than its margin), aligned in that if they
 * say so. They had stayed where their paragraph is: a header's page-placed banner (wrapped square) drawn at the header's
 * top, half an inch below where Word puts it, its paragraph pushed down under it and the body with it (a third of an
 * inch in the samples), and a logo placed above the margin pulled down with the body.
 * The page's geometry comes from `el.dataset` (odxPh: the page's height, odxMtop and odxMbot: where its body starts and
 * how far from the foot it ends), set where the page is drawn; without it (a story measured for its height) the shapes
 * only leave the flow. Their side to side position stays.
 */
function placeStoryShapes(el) {
  const imgs = [...el.querySelectorAll('img[style*="--odx-lrel"], img[style*="--odx-arel"]')];
  if (!imgs.length) return;
  const d = el.dataset;
  const geo = d.odxPh != null ? { ph: Number(d.odxPh), mtop: Number(d.odxMtop), mbot: Number(d.odxMbot), top: el.offsetTop } : null;
  const lefts = imgs.map((img) => img.offsetLeft);
  imgs.forEach((img, i) => {
    const st = img.style;
    const v = (k) => st.getPropertyValue(k).trim();
    const lifted = !!v('--odx-lrel');
    const rel = lifted ? v('--odx-lrel') : v('--odx-arel');
    const off = parseFloat(lifted ? v('--odx-ly') : v('--odx-ay')) || 0;
    const align = lifted ? v('--odx-lyalign') : v('--odx-ayalign');
    st.float = 'none';
    st.position = 'absolute';
    st.marginTop = '0';
    // (One set across the page from the page's edge keeps its own place, which goes by where the story's box starts.)
    if (!st.left.includes('--odx-hf-left')) st.left = `${lefts[i]}px`;
    st.marginLeft = '0';
    if (!geo || !Object.values(geo).every(Number.isFinite)) return;
    // Its area on the page (px from the page's top): the page, its margin box, or the margin above or below it.
    const [from, span] = rel === 'margin' ? [geo.mtop, geo.ph - geo.mtop - geo.mbot] : rel === 'topMargin' ? [0, geo.mtop]
      : rel === 'bottomMargin' ? [geo.ph - geo.mbot, geo.mbot] : [0, geo.ph];
    const h = img.offsetHeight;
    const y = align === 'top' ? from : align === 'center' ? from + (span - h) / 2 : align === 'bottom' ? from + span - h : from + off;
    // (Its top in its containing block, which may be its paragraph: from where it is now, at the story's zoom.)
    st.top = '0px';
    const er = el.getBoundingClientRect();
    // (The zoom from the page the story is set in, where there is one: a one-line header's own height, 15.33px read
    // as a whole 15, gave a zoom of 1.02 and its shapes a place a pixel low.)
    const box = el.offsetParent;
    const scale = box && box.offsetHeight > 100 ? box.getBoundingClientRect().height / box.offsetHeight : el.offsetHeight ? er.height / el.offsetHeight : 1;
    const now = img.getBoundingClientRect().top;
    // (The story's top on its page to the fraction, where the box it is set in gives the same within a pixel:
    // offsetTop is a whole number of pixels, and a footer starts where its lines' heights put it, up to half a pixel
    // off.)
    const host = el.offsetParent;
    const exact = host ? (er.top - host.getBoundingClientRect().top) / (scale || 1) - host.clientTop : NaN;
    const top = Math.abs(exact - geo.top) < 1 ? exact : geo.top;
    st.top = `${((er.top + (y - top) * scale - now) / (scale || 1)).toFixed(2)}px`;
  });
}

/**
 * A story's paragraphs (a header's, a footer's, a note's) sized as in the body: an empty one by its mark's font and size,
 * one whose runs all share a size or font by that, and spaces with none of their own in the style's. `nodes`: a map to
 * fill, element -> node.
 */
export function sizeStoryLines(session, doc, el, nodes = null) {
  // An empty line is as tall as its paragraph mark's font, as in the body.
  const ps = el.querySelectorAll('p');
  let n = 0;
  doc.descendants((node) => {
    if (node.type !== schema.nodes.paragraph) return true;
    const p = ps[n++];
    if (p) nodes?.set(p, node);
    // A paragraph deleted whole (text and mark) disappears unless all markup is shown, as in the body.
    if (p && deletedWhole(node)) p.classList.add('odx-all-del');
    // (Anchored pictures take no room on the line, so a paragraph holding only those counts as empty.)
    const blankText = !node.textContent.trim();
    let pics = false;
    if (blankText) node.descendants((c) => { if (c.type === schema.nodes.image && !/position:\s*absolute/.test(c.attrs.place || '')) pics = true; });
    // Pictures with nothing but spaces beside them: the spaces size no line, before, after or between the pictures and
    // in any size, and the mark's line is the floor, as with no spaces (hdrht-word a2-a4, a8-a11: the next line where it
    // is under the picture alone, or under a 20pt mark's line; a sample's header logo with 26 spaces after it had made its
    // line 4.9px taller, every later page that much low).
    // Tabs between or beside them are the same: two logos with tabs between, a line as tall as the taller logo in Word;
    // the tabs' text had added its descent under the pictures, 4.4px (a sample template's body that much low on every
    // page).
    if (p && pics) for (const b of p.querySelectorAll('span.odx-blank, span.odx-tab > span.odx-t')) b.style.lineHeight = '0';
    // (The paragraph style's size, for lines dropBelowLogo adds.)
    if (p && p.querySelector('img[style*="--odx-lbelow"]')) p.style.setProperty('--odx-ssz', `${Number(session.styleRunProp(node.attrs.styleId, 'size') || session.styles?.defaults?.rPr?.size || 20) / 2}pt`);
    // Ending in a line break: its empty last line holds only the paragraph mark, as tall as the mark's size (its own,
    // else the style's), as in the body (pararules.js lineSize, editor.css odx-endbreak). A trailing <br> alone draws no
    // line: a header's last paragraph closing with a break over a 6pt mark had lost that line, 9.4px, the body as much
    // high on every page.
    const end = node.lastChild;
    if (p && end?.type === schema.nodes.hard_break && (!end.attrs.type || end.attrs.type === 'textWrapping')) {
      p.classList.add('odx-endbreak');
      const size = markSize(node) || Number(session.styleRunProp(node.attrs.styleId, 'size') || session.styles?.defaults?.rPr?.size || 20);
      p.style.setProperty('--odx-endmark', `${size / 2}pt`);
    }
    const mark = ((node.attrs.extra || []).find((e) => e[0] === 'rPr') || [])[1] || '';
    const sz = blankText && /<w:sz w:val="(\d+)"/.exec(mark);
    if (p && sz) p.style.fontSize = `${sz[1] / 2}pt`;
    const font = blankText && /<w:rFonts\b[^>]*w:(?:ascii|hAnsi)="([^"]+)"/.exec(mark);
    if (p && font) { p.style.fontFamily = fontStack(font[1]); p.style.setProperty('--odx-lh', lineFactor(font[1])); }
    // (Or the theme's font, w:asciiTheme, as in the body, runs.js: a footer's empty last paragraph, its mark the theme's
    // Calibri under a Times style, is Calibri's 17.9px line at 11pt, not Times's 16.9px; the line above it 1pt higher.)
    const themed = blankText && !font && session.theme && /<w:rFonts\b[^>]*\bw:(?:ascii|hAnsi)Theme="((?:major|minor)\w*)/.exec(mark);
    const tf = themed && themeFace(session.theme, themed[1]);
    if (p && tf) { p.style.fontFamily = fontStack(tf); p.style.setProperty('--odx-lh', lineFactor(tf)); }
    // With text, as in the body: when every run shares a size or font, the line is sized from it.
    const run = !blankText && p && uniformRun(node, (id) => session.charStyle?.(id), session.theme);
    if (run && run.size) p.style.fontSize = `${run.size / 2}pt`;
    else if (run && run.strut && run.raised && node.attrs.numId == null && !node.content.content.some((c) => c.type.name === 'hard_break')) p.style.fontSize = '0';
    else if (run && run.strut && Number(run.strut) < Number(session.styleRunProp(node.attrs.styleId, 'size') || 20)) p.style.fontSize = `${run.strut / 2}pt`;
    if (run && run.font) { p.style.fontFamily = fontStack(run.font); p.style.setProperty('--odx-lh', lineFactor(run.font)); }
    // Spaces with no size or font of their own keep the paragraph style's, not the text's given to the paragraph
    // above: 28 spaces before 20pt text in an 11pt style are 11pt spaces in Word, 93px wide; at the text's 20pt they
    // were 169px, the line 69px too wide for its header and wrapped to two, the whole body 33px low.
    // (They size no line, as in Word: 10pt spaces beside 9pt text had made a header line 0.6px taller.)
    if (run && (run.size || run.font) && node.textContent.trim()) {
      const size = Number(session.styleRunProp(node.attrs.styleId, 'size') || 20);
      const tfn = session.styleRunProp(node.attrs.styleId, 'themeFont');
      const fname = session.styleRunProp(node.attrs.styleId, 'font') || (tfn && themeFace(session.theme, tfn)) || null;
      for (const b of p.querySelectorAll('span.odx-blank')) {
        let sized = false;
        let fonted = false;
        for (let a = b.parentElement; a && a !== p; a = a.parentElement) {
          if (a.style.fontSize || a.dataset.cstyle) sized = true;
          if (a.style.fontFamily || a.dataset.cstyle) fonted = true;
        }
        if (run.size && !sized && size !== Number(run.size)) { b.style.fontSize = `${size / 2}pt`; b.style.lineHeight = '0'; }
        if (run.font && !fonted && fname && fname !== run.font) b.style.fontFamily = fontStack(fname);
      }
    }
    return false;
  });
}

/**
 * A header's or footer's paragraph holding a picture wrapped top and bottom, placed from the paragraph's top (reader.js
 * fromParagraphTop): in Word 2013 and later (compatibility mode 15) the paragraph's space before comes again under the
 * picture, then its own line, even with no text of its own (its mark's line), before what follows (hdrtb-word,
 * odx-test\mkhdrtb.mjs: a 1pt mark spaced 6pt before under a picture reaching 7.5pt past the top margin, the body
 * 7.2pt lower than the picture's foot; spaced none, 1.2pt; an 11pt mark, 13.3pt more; header text, the same). The
 * picture itself is where it was. Mode 14 and older: the body right under the picture, as before. (In the body Word
 * does otherwise: TODO.md.)
 */
function underTopAndBottom(el) {
  for (const p of el.querySelectorAll('p')) {
    if (!p.querySelector(':scope > img.odx-wrap-block[style*="--odx-sb"]')) continue;
    p.style.setProperty('--odx-sbtb', 'var(--odx-sb, 0px)');
    if (!p.textContent.trim()) p.classList.add('odx-tbline');
  }
}

/**
 * A function drawing one story of `session`: `(doc, cls, pageNo, pages)` -> its element, classed `odx-hf <cls>`, its
 * PAGE fields showing `pageNo` (a number, or { n, fmt }) and NUMPAGES `pages`.
 */
export function storyDrawer(session) {
  const serializer = DOMSerializer.fromSchema(schema);
  let props = null; // the file's properties, for header and footer fields
  const styleTables = (doc, el) => styleStoryTables(session, doc, el);
  return (story, cls, pageNo, pages) => {
    // (Preview Results: the merge fields filled with the record's values, mailmerge.js.)
    const doc = previewStory(session, story);
    const el = document.createElement('div');
    el.className = `odx-hf ${cls}`;
    el.append(serializer.serializeFragment(doc.content));
    wrapText(el); // (lines as tall as their text, as in the body)
    // The spaces that end a paragraph are marked, as in the body (textspans.js): under a centred or right-aligned
    // paragraph they take no room (a header set right, its text ending in a space, a space's width left of Word's).
    for (const p of el.querySelectorAll('p')) {
      const spans = [...p.querySelectorAll('span.odx-t')].filter((x) => x.childNodes.length === 1 && x.firstChild.nodeType === 3);
      const last = spans[spans.length - 1];
      if (!last || last !== [...p.querySelectorAll('span.odx-t')].pop()) continue;
      const t = last.firstChild;
      let from = t.data.length;
      while (from > 0 && t.data.charCodeAt(from - 1) === 32) from--;
      if (from === t.data.length || !(from > 0 || spans.some((x) => x !== last && x.textContent.trim()))) continue;
      const trail = document.createElement('span');
      trail.className = 'odx-trail';
      if (from === 0) { last.replaceChildren(trail); trail.append(t); } else { const rest = t.splitText(from); trail.append(rest); last.append(trail); }
    }
    sizeStoryLines(session, doc, el, hfNodes);
    if (session.compatMode >= 15) underTopAndBottom(el);
    // "Don't add space between paragraphs of the same style" (contextualSpacing), as in the body (editor.js): each such
    // paragraph loses its own space next to one of its style (Title paragraphs in a header 6px apart otherwise).
    const top = [];
    doc.forEach((node) => top.push(node));
    const tps = [...el.children];
    const csOf = (node) => node.attrs.view?.contextualSpacing ?? session.styleProp?.(node.attrs.styleId, 'contextualSpacing');
    for (let i = 1; i < top.length && i < tps.length; i++) {
      const a = top[i - 1], b = top[i];
      if (a.type !== schema.nodes.paragraph || b.type !== schema.nodes.paragraph || a.attrs.styleId !== b.attrs.styleId || tps[i - 1].tagName !== 'P' || tps[i].tagName !== 'P') continue;
      if (csOf(b)) tps[i].style.marginTop = '0';
      if (csOf(a)) tps[i - 1].style.marginBottom = '0';
    }
    // Paragraphs in a row with the same borders (their own or their style's) are one bordered group in Word, as in the
    // body (editor.js): its top border above the first, its bottom border under the last, a "between" line if it has one
    // between them. (Only borders written on the paragraphs had been grouped: three Header-style paragraphs under the
    // style's 0.5pt rule took a rule and its space each, the second and third lines 1 and 3px low in hdrht-word d2, a
    // sample's second header line 1px low on every page.)
    const own = (node, k) => node.attrs.view?.[k];
    const prop = (node, k) => own(node, k) ?? session.styleProp?.(node.attrs.styleId, k);
    const borderKey = (node) => {
      if (node.type !== schema.nodes.paragraph) return null;
      const b = prop(node, 'borders');
      return b && (b.top || b.bottom || b.left || b.right) ? JSON.stringify([b, prop(node, 'indLeft') ?? 0, prop(node, 'indRight') ?? 0]) : null;
    };
    const keys = top.map(borderKey);
    for (let i = 1; i < top.length && i < tps.length; i++) {
      const a = top[i - 1], b = top[i];
      if (!keys[i] || keys[i] !== keys[i - 1] || tps[i - 1].tagName !== 'P' || tps[i].tagName !== 'P') continue;
      // (Of one style, or both with borders of their own, as the paragraphs grouped before.)
      if (a.attrs.styleId !== b.attrs.styleId && !(own(a, 'borders') && own(b, 'borders'))) continue;
      const between = prop(b, 'borders')?.between;
      const line = between && between.val && between.val !== 'nil' && between.val !== 'none';
      tps[i].style.borderTop = 'none';
      tps[i].style.paddingTop = `${line ? between.space || 0 : 0}pt`;
      tps[i - 1].style.borderBottom = line ? borderCss(between) : 'none';
      tps[i - 1].style.paddingBottom = `${line ? between.space || 0 : 0}pt`;
    }
    styleTables(doc, el);
    fillFields(el, { PAGE: pageNo?.n ?? pageNo, PAGEFMT: pageNo?.fmt, NUMPAGES: pages, SECTIONPAGES: pages, FILENAME: session.fileName || null, PROPS: props || (props = docProps(session.pkg)), RESOLVE: pageNo?.resolve });
    return el;
  };
}
