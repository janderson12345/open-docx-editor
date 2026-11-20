// Headers and footers as drawn on the sheets: a story's content serialized, sized as Word sizes its lines, its
// tables styled and its fields filled in for the page.
import { DOMSerializer } from 'prosemirror-model';
import { TableMap } from 'prosemirror-tables';
import { schema } from './schema.js';
import { tableCellCss, tablePlacementCss } from '../ooxml/styles.js';
import { parseXml } from '../ooxml/xml.js';
import { twipsToPx } from '../ooxml/units.js';
import { fontStack, lineFactor } from '../ooxml/props.js';
import { deletedWhole } from './pararules.js';
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
    const place = tablePlacementCss(session.styles, t.attrs.tblPr, session.compatMode, parseXml, t.attrs.tblGrid);
    // (A note's tables don't float: Word has no text wrapping in notes, and puts one that says so in the note's flow.)
    if (place && !(note && /(?:^|;)float:/.test(place))) dom.style.cssText += ';' + place;
    // (Floating, placed from the text above its own place: up too, as Word draws it in a header (its PDF: a table at
    // tblpY -200 starts 10pt above the paragraph after it; it had been drawn at that paragraph's top). The body's own
    // floating tables keep theirs at 0 or below, pages.js.)
    const up = !note && /(?:^|;)float:/.test(place) && !/\bw:vertAnchor="page"/.test(t.attrs.tblPr || '') ? /\bw:tblpY="(-\d{1,6})"/.exec(t.attrs.tblPr || '') : null;
    if (up) dom.style.top = `${Math.max(-1440, twipsToPx(Number(up[1]))).toFixed(2)}px`;
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
    const res = tableCellCss(session.styles, t.attrs.tblPr, map.height, map.width, cells, parseXml, session.theme);
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
 * A drawn header's or footer's shapes that the body's text wraps round (reader.js anchorCss, --odx-hfwrap), placed below
 * their own paragraph: out of the story's flow, where they're drawn. As floats their room reached from the paragraph's
 * top down to them, pushing the header's own text aside and the header's foot (and the body under it) down past them,
 * where Word keeps both as without the shape (footwrap-word fw-header, a rectangle 250pt below its header paragraph).
 * `el` must be in the document, to be measured.
 */
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

export function settleStoryShapes(el) {
  placeStoryShapes(el);
  const below = [];
  for (const img of el.querySelectorAll('img[style*="--odx-hfwrap"]')) {
    const p = img.closest('p');
    const cs = getComputedStyle(img);
    if (!p || cs.float === 'none' || (parseFloat(cs.marginTop) || 0) < p.getBoundingClientRect().height) continue;
    below.push({ img, at: img.getBoundingClientRect() });
  }
  for (const b of below) { b.img.style.float = 'none'; b.img.style.position = 'absolute'; }
  for (const { img, at } of below) {
    const now = img.getBoundingClientRect();
    img.style.translate = `${(at.left - now.left).toFixed(2)}px ${(at.top - now.top).toFixed(2)}px`;
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
    st.left = `${lefts[i]}px`;
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
    const scale = el.offsetHeight ? er.height / el.offsetHeight : 1;
    const now = img.getBoundingClientRect().top;
    st.top = `${((er.top + (y - geo.top) * scale - now) / (scale || 1)).toFixed(2)}px`;
  });
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
    // An empty line is as tall as its paragraph mark's font, as in the body.
    const ps = el.querySelectorAll('p');
    let n = 0;
    doc.descendants((node) => {
      if (node.type !== schema.nodes.paragraph) return true;
      const p = ps[n++];
      if (p) hfNodes.set(p, node);
      // A paragraph deleted whole (text and mark) disappears unless all markup is shown, as in the body.
      if (p && deletedWhole(node)) p.classList.add('odx-all-del');
      // (Anchored pictures take no room on the line, so a paragraph holding only those counts as empty.)
      let empty = !node.textContent.trim();
      if (empty) node.descendants((c) => { if (c.type === schema.nodes.image && !/position:\s*absolute/.test(c.attrs.place || '')) empty = false; });
      const mark = ((node.attrs.extra || []).find((e) => e[0] === 'rPr') || [])[1] || '';
      const sz = empty && /<w:sz w:val="(\d+)"/.exec(mark);
      if (p && sz) p.style.fontSize = `${sz[1] / 2}pt`;
      const font = empty && /<w:rFonts\b[^>]*w:(?:ascii|hAnsi)="([^"]+)"/.exec(mark);
      if (p && font) { p.style.fontFamily = fontStack(font[1]); p.style.setProperty('--odx-lh', lineFactor(font[1])); }
      // (Or the theme's font, w:asciiTheme, as in the body, runs.js: a footer's empty last paragraph, its mark the theme's
      // Calibri under a Times style, is Calibri's 17.9px line at 11pt, not Times's 16.9px; the line above it 1pt higher.)
      const themed = empty && !font && session.theme && /<w:rFonts\b[^>]*\bw:(?:ascii|hAnsi)Theme="(major|minor)/.exec(mark);
      const tf = themed && session.theme[themed[1]];
      if (p && tf) { p.style.fontFamily = fontStack(tf); p.style.setProperty('--odx-lh', lineFactor(tf)); }
      // With text, as in the body: when every run shares a size or font, the line is sized from it.
      const run = !empty && p && uniformRun(node, (id) => session.charStyle?.(id), session.theme);
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
        const fname = session.styleRunProp(node.attrs.styleId, 'font') || (tfn && session.theme?.[tfn.startsWith('major') ? 'major' : 'minor']) || null;
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
    // Paragraphs in a row with the same borders are one bordered group in Word: one bottom border, under the last.
    const sides = (p) => ['Top', 'Bottom', 'Left', 'Right'].map((s) => p.style[`border${s}`] + '|' + p.style[`padding${s}`]).join(';');
    const kids = [...el.children];
    const sig = kids.map((p) => (p.tagName === 'P' && (p.style.borderBottom || p.style.borderTop) ? sides(p) : null));
    for (let i = 1; i < kids.length; i++) {
      const a = kids[i - 1];
      const b = kids[i];
      if (!sig[i] || sig[i] !== sig[i - 1]) continue;
      a.style.borderBottom = a.style.paddingBottom = '';
      b.style.borderTop = b.style.paddingTop = '';
    }
    styleTables(doc, el);
    fillFields(el, { PAGE: pageNo?.n ?? pageNo, PAGEFMT: pageNo?.fmt, NUMPAGES: pages, SECTIONPAGES: pages, FILENAME: session.fileName || null, PROPS: props || (props = docProps(session.pkg)), RESOLVE: pageNo?.resolve });
    return el;
  };
}
