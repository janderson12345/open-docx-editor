// Headers and footers as drawn on the sheets: a story's content serialized, sized as Word sizes its lines, its
// tables styled and its fields filled in for the page.
import { DOMSerializer } from 'prosemirror-model';
import { TableMap } from 'prosemirror-tables';
import { schema } from './schema.js';
import { tableCellCss, tablePlacementCss } from '../ooxml/styles.js';
import { parseXml } from '../ooxml/xml.js';
import { fontStack, lineFactor } from '../ooxml/props.js';
import { deletedWhole } from './pararules.js';
import { wrapText } from './textspans.js';
import { uniformRun } from './runs.js';
import { docProps, fillFields } from './pagefields.js';

/** Header and footer paragraphs -> their nodes, for laying out their tabs once drawn. */
export const hfNodes = new WeakMap();

/**
 * A function drawing one story of `session`: `(doc, cls, pageNo, pages)` -> its element, classed `odx-hf <cls>`, its
 * PAGE fields showing `pageNo` (a number, or { n, fmt }) and NUMPAGES `pages`.
 */
export function storyDrawer(session) {
  const serializer = DOMSerializer.fromSchema(schema);
  let props = null; // the file's properties, for header and footer fields
  // Header/footer tables get the same placement, borders and shading the body's decorations give.
  const styleTables = (doc, el) => {
    const tables = [];
    doc.descendants((n) => { if (n.type === schema.nodes.table) tables.push(n); });
    const doms = el.querySelectorAll('table');
    tables.forEach((t, k) => {
      const dom = doms[k];
      if (!dom) return;
      const place = tablePlacementCss(session.styles, t.attrs.tblPr, session.compatMode, parseXml);
      if (place) dom.style.cssText += ';' + place;
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
  };
  return (doc, cls, pageNo, pages) => {
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
      // With text, as in the body: when every run shares a size or font, the line is sized from it.
      const run = !empty && p && uniformRun(node, (id) => session.charStyle?.(id), session.theme);
      if (run && run.size) p.style.fontSize = `${run.size / 2}pt`;
      else if (run && run.strut && run.raised && node.attrs.numId == null && !node.content.content.some((c) => c.type.name === 'hard_break')) p.style.fontSize = '0';
      else if (run && run.strut && Number(run.strut) < Number(session.styleRunProp(node.attrs.styleId, 'size') || 20)) p.style.fontSize = `${run.strut / 2}pt`;
      if (run && run.font) { p.style.fontFamily = fontStack(run.font); p.style.setProperty('--odx-lh', lineFactor(run.font)); }
      return false;
    });
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
    fillFields(el, { PAGE: pageNo?.n ?? pageNo, PAGEFMT: pageNo?.fmt, NUMPAGES: pages, SECTIONPAGES: pages, PROPS: props || (props = docProps(session.pkg)) });
    return el;
  };
}
