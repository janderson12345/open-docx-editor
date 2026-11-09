// ProseMirror document -> word/document.xml.
// Writes modelled nodes/marks as WordprocessingML and splices raw XML back in
// for everything that was carried through untouched.
import { TableMap } from 'prosemirror-tables';
import { esc } from './xml.js';
import { EMU_PER_PX, pxToEmu } from './units.js';

// Child order required by the WordprocessingML schema (ECMA-376 Part 1, §17).
const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'];
const PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
const TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark', 'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange'];

function ordered(order, items) {
  // Extension elements (w14:shadow and the like) go after all of w:'s, as the schema wants: Word ignores
  // them anywhere else. Their local names can clash with w: ones (w14:shadow vs w:shadow).
  const rank = ([n, xml]) => {
    if (/^<(?!w:)[\w]+:/.test(xml)) return order.length + 1;
    const i = order.indexOf(n);
    return i < 0 ? order.length : i;
  };
  return items
    .map((it, i) => [it, i])
    .sort((a, b) => rank(a[0]) - rank(b[0]) || a[1] - b[1])
    .map(([it]) => it[1])
    .join('');
}

const val = (name, v) => `<w:${name} w:val="${esc(v)}"/>`;
const onOff = (name, on) => (on ? `<w:${name}/>` : `<w:${name} w:val="0"/>`);

/**
 * @param doc ProseMirror Node (type doc)
 * @param originalXml the document.xml text being replaced (its root and namespaces are kept)
 * @param ctx { linkRel(href) -> rId } allocates relationships for new hyperlinks
 */
export function writeDocument(doc, originalXml, ctx = {}) {
  const parts = [writeBlocks(doc, ctx)];
  if (doc.lastChild && doc.lastChild.type.name === 'table') parts.push('<w:p/>');
  if (doc.attrs.sectPr) parts.push(doc.attrs.sectPr);
  const body = parts.join('');
  const m = /<w:body(\s[^>]*)?>[\s\S]*<\/w:body>|<w:body(\s[^>]*)?\/>/.exec(originalXml);
  if (!m) throw new Error('document.xml has no w:body');
  return originalXml.slice(0, m.index) + `<w:body>${body}</w:body>` + originalXml.slice(m.index + m[0].length);
}

/** A container's blocks, with runs of blocks from one content control wrapped in its w:sdt again. */
export function writeBlocks(parent, ctx) {
  let out = '';
  let open = null;
  parent.forEach((b) => {
    const sdt = b.attrs.sdt || null;
    if (open && (!sdt || sdt.g !== open.g)) { out += open.close; open = null; }
    if (sdt && !open) { out += sdt.open; open = sdt; }
    out += writeBlock(b, ctx);
  });
  if (open) out += open.close;
  return out;
}

function writeBlock(node, ctx) {
  switch (node.type.name) {
    case 'paragraph': return writeParagraph(node, ctx);
    case 'table': return writeTable(node, ctx);
    case 'opaque_block': return node.attrs.xml;
    default: throw new Error('Unknown block ' + node.type.name);
  }
}

function writeParagraph(p, ctx) {
  const a = p.attrs;
  return `<w:p>${pPrXml(a)}${writeInlines(p, ctx)}</w:p>`;
}

/** A paragraph's w:pPr from its attrs ('' when empty). `skip` leaves out raw children by name. */
export function pPrXml(a, skip = []) {
  const items = a.extra.filter((e) => !skip.includes(e[0]));
  if (a.styleId) items.push(['pStyle', val('pStyle', a.styleId)]);
  if (a.numId != null) items.push(['numPr', `<w:numPr>${val('ilvl', a.ilvl || 0)}${val('numId', a.numId)}</w:numPr>`]);
  if (a.align) items.push(['jc', val('jc', a.align)]);
  return items.length ? `<w:pPr>${ordered(PPR_ORDER, items)}</w:pPr>` : '';
}

// --- Inline content -------------------------------------------------------

function writeInlines(parent, ctx) {
  const items = [];
  parent.forEach((child) => items.push(child));
  return wrapGroups(items, ['sdt', 'ins', 'del', 'link'], ctx, false);
}

// Group runs into wrappers, outermost first: tracked insertions (<w:ins>), tracked
// deletions (<w:del>, which may sit inside an insertion), then hyperlinks.
function wrapGroups(nodes, levels, ctx, deleted) {
  if (!levels.length) return nodes.map((n) => writeInline(n, deleted)).join('');
  const [name, ...rest] = levels;
  const markOf = (n) => n.marks.find((m) => m.type.name === name) || null;
  let out = '';
  for (let i = 0; i < nodes.length; ) {
    const mark = markOf(nodes[i]);
    let j = i;
    while (j < nodes.length && sameMark(markOf(nodes[j]), mark)) j++;
    // (Deleted text is w:delText; moved-from text keeps w:t, or Word calls the file corrupt.)
    const inner = wrapGroups(nodes.slice(i, j), rest, ctx, deleted || (mark && name === 'del' && !mark.attrs.move));
    if (!mark) out += inner;
    else if (name === 'link') out += wrapLink(mark, inner, ctx);
    else if (name === 'sdt') out += wrapField(mark, nodes.slice(i, j), inner);
    else {
      const tag = mark.attrs.move ? (name === 'del' ? 'moveFrom' : 'moveTo') : name; // (moved text, Word's own elements)
      out += `<w:${tag}${attrsXml({ 'w:id': revisionId(ctx, mark.attrs.id), 'w:author': mark.attrs.author, 'w:date': mark.attrs.date })}>${inner}</w:${tag}>`;
    }
    i = j;
  }
  return out;
}

// Each <w:ins>/<w:del> needs its own w:id; a change split in two (around a link, say) gets a new one.
function revisionId(ctx, id) {
  const used = (ctx.revIds ||= new Set());
  let v = id ?? '0';
  if (used.has(v)) v = String(Math.max(1e6, ...[...used].map(Number).filter(Number.isFinite)) + 1);
  used.add(v);
  return v;
}

const sameMark = (a, b) => (a === null || b === null ? a === b : a.eq(b));

function attrsXml(o) {
  return Object.entries(o)
    .filter(([, v]) => v != null)
    .map(([k, v]) => ` ${k}="${esc(v)}"`)
    .join('');
}

function wrapField(mark, nodes, inner) {
  let { pr, endPr, phText } = mark.attrs;
  // Text typed over the placeholder is a real value now.
  if (phText != null && nodes.map((n) => n.textContent).join('') !== phText) pr = pr.replace(/<w:showingPlcHdr\b[^>]*\/>/g, '');
  return `<w:sdt>${pr}${endPr}<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
}

function wrapLink(link, runs, ctx) {
  let { rId, href, anchor, extra } = link.attrs;
  if (!rId && href && ctx.linkRel) rId = ctx.linkRel(href);
  return `<w:hyperlink${attrsXml({ 'r:id': rId, 'w:anchor': anchor, ...(extra || {}) })}>${runs}</w:hyperlink>`;
}

function writeInline(node, deleted) {
  switch (node.type.name) {
    case 'text': return `<w:r>${rPr(node.marks)}${textContent(node.text, deleted)}</w:r>`;
    case 'tab': return `<w:r>${rPr(node.marks)}<w:tab/></w:r>`;
    case 'hard_break': return `<w:r>${rPr(node.marks)}<w:br${attrsXml({ 'w:type': node.attrs.type, 'w:clear': node.attrs.clear })}/></w:r>`;
    case 'image': return imageXml(node);
    case 'opaque_inline': return node.attrs.xml;
    default: throw new Error('Unknown inline ' + node.type.name);
  }
}

/**
 * A picture's drawing XML, its size updated when it was resized in the editor (the editor keeps its size to a
 * hundredth of a pixel; an unchanged picture keeps its exact EMU size).
 */
function imageXml(node) {
  const { xml, width, height } = node.attrs;
  const m = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(xml || '');
  if (!m || width == null || height == null) return xml;
  const cx = Number(m[1]);
  const cy = Number(m[2]);
  const same = (e, px) => Math.round(e / EMU_PER_PX) === px || Math.abs(e / EMU_PER_PX - px) < 0.01; // (or whole pixels, as before)
  if (same(cx, width) && same(cy, height)) return xml;
  const nx = pxToEmu(width);
  const ny = pxToEmu(height);
  // The frame (wp:extent) and the picture's own transform (a:ext inside pic:spPr), scaled alike.
  return xml.replace(m[0], `<wp:extent cx="${nx}" cy="${ny}"`)
    .replace(/(<pic:spPr\b[\s\S]*?<a:ext )cx="(\d+)" cy="(\d+)"/, (all, head, ax, ay) => `${head}cx="${Math.round(Number(ax) * nx / cx)}" cy="${Math.round(Number(ay) * ny / cy)}"`);
}

function textContent(text, deleted) {
  // Characters with their own WordprocessingML elements.
  let out = '';
  let buf = '';
  const flush = () => {
    if (buf) out += deleted ? `<w:delText xml:space="preserve">${esc(buf)}</w:delText>` : `<w:t xml:space="preserve">${esc(buf)}</w:t>`;
    buf = '';
  };
  for (const ch of text) {
    const el = { '\t': '<w:tab/>', '\n': '<w:br/>', '‑': '<w:noBreakHyphen/>', '­': '<w:softHyphen/>' }[ch];
    if (el) {
      flush();
      out += el;
    } else buf += ch;
  }
  flush();
  return out;
}

/** w:rPr XML for a run's marks ('' when none). */
export function rPr(marks) {
  const items = [];
  for (const m of marks) {
    const a = m.attrs;
    switch (m.type.name) {
      case 'cstyle': items.push(['rStyle', val('rStyle', a.val)]); break;
      case 'font': {
        const keep = a.xml && new RegExp(`w:(ascii|hAnsi)="${a.val.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`).test(a.xml);
        items.push(['rFonts', keep ? a.xml : `<w:rFonts w:ascii="${esc(a.val)}" w:hAnsi="${esc(a.val)}" w:cs="${esc(a.val)}"/>`]);
        break;
      }
      case 'bold': items.push(['b', onOff('b', a.on)]); break;
      case 'italic': items.push(['i', onOff('i', a.on)]); break;
      case 'strike': items.push(['strike', onOff('strike', a.on)]); break;
      case 'color': items.push(['color', val('color', a.val)]); break;
      case 'size': items.push(['sz', val('sz', a.val)]); break;
      case 'highlight': items.push(['highlight', val('highlight', a.val)]); break;
      case 'underline': items.push(['u', val('u', a.val)]); break;
      case 'vertAlign': items.push(['vertAlign', val('vertAlign', a.val)]); break;
      case 'rextra': items.push(...a.items); break;
    }
  }
  return items.length ? `<w:rPr>${ordered(RPR_ORDER, items)}</w:rPr>` : '';
}

// --- Tables ---------------------------------------------------------------

const DEFAULT_TBLPR = '<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:tblLook w:val="04A0"/></w:tblPr>';
const sameWidths = (a, b) => JSON.stringify(a || null) === JSON.stringify(b || null);

function writeTable(table, ctx) {
  const map = TableMap.get(table);
  let resized = false;
  table.descendants((n) => {
    if (n.type.name === 'table_cell' && !sameWidths(n.attrs.colwidth, n.attrs.colwidth0)) resized = true;
    return n.type.name !== 'table_cell';
  });
  let grid = table.attrs.tblGrid;
  if (!grid || resized) {
    const widths = columnWidths(table, map);
    grid = `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
  }
  let rows = '';
  let openSdt = null; // row-level content control being written
  for (let r = 0; r < map.height; r++) {
    const row = table.child(r);
    const sdt = row.attrs.sdt || null;
    if (openSdt && (!sdt || sdt.g !== openSdt.g)) { rows += openSdt.close; openSdt = null; }
    if (sdt && !openSdt) { rows += sdt.open; openSdt = sdt; }
    let cells = '';
    let openCellSdt = null; // cell-level content control being written
    for (let c = 0; c < map.width; c++) {
      const pos = map.map[r * map.width + c];
      const cell = table.nodeAt(pos);
      if (!cell || !cell.type.spec.tableRole?.includes('cell')) continue; // a short row (fewer cells than the grid)
      const rect = map.findCell(pos);
      if (rect.left !== c) continue; // covered by a horizontal span
      if (cell.attrs.phantom) continue; // gridBefore/gridAfter, kept in trPr
      // (A merged cell's control around the row it starts in only: its continuing rows are the same node, and the
      // control written around each of them would repeat one w:id.)
      const cs = rect.top === r ? cell.attrs.sdt || null : null;
      if (openCellSdt && (!cs || cs.g !== openCellSdt.g)) { cells += openCellSdt.close; openCellSdt = null; }
      if (cs && !openCellSdt) { cells += cs.open; openCellSdt = cs; }
      cells += writeCell(cell, rect.top === r ? (cell.attrs.rowspan > 1 ? 'restart' : null) : 'continue', ctx);
    }
    if (openCellSdt) cells += openCellSdt.close;
    rows += `<w:tr>${row.attrs.tblPrEx || ''}${row.attrs.trPr || ''}${cells}</w:tr>`;
  }
  if (openSdt) rows += openSdt.close;
  return `<w:tbl>${table.attrs.tblPr || DEFAULT_TBLPR}${grid}${rows}</w:tbl>`;
}

function columnWidths(table, map) {
  const widths = new Array(map.width).fill(null);
  for (let i = 0; i < map.map.length; i++) {
    const pos = map.map[i];
    const col = i % map.width;
    const cell = table.nodeAt(pos);
    if (!cell || !cell.type.spec.tableRole?.includes('cell')) continue;
    const rect = map.findCell(pos);
    const cw = cell.attrs.colwidth;
    if (cw && cw[col - rect.left] && widths[col] == null) widths[col] = Math.round(cw[col - rect.left] * 15);
  }
  const fallback = Math.round(9000 / map.width);
  return widths.map((w) => w ?? fallback);
}

function writeCell(cell, vMerge, ctx) {
  const a = cell.attrs;
  let items = a.tcPr.slice();
  if (!sameWidths(a.colwidth, a.colwidth0) && a.colwidth) {
    items = items.filter((it) => it[0] !== 'tcW');
    items.push(['tcW', `<w:tcW w:w="${Math.round(a.colwidth.reduce((s, w) => s + w, 0) * 15)}" w:type="dxa"/>`]);
  }
  if (a.colspan > 1) items.push(['gridSpan', val('gridSpan', a.colspan)]);
  if (vMerge === 'restart') items.push(['vMerge', val('vMerge', 'restart')]);
  if (vMerge === 'continue') items.push(['vMerge', '<w:vMerge/>']);
  const tcPr = items.length ? `<w:tcPr>${ordered(TCPR_ORDER, items)}</w:tcPr>` : '';
  if (vMerge === 'continue') return `<w:tc>${tcPr}<w:p/></w:tc>`;
  let body = '';
  body += writeBlocks(cell, ctx);
  if (!cell.lastChild || cell.lastChild.type.name !== 'paragraph') body += '<w:p/>';
  return `<w:tc>${tcPr}${body}</w:tc>`;
}

/** A header or footer part (w:hdr / w:ftr) with its content replaced by `doc`'s; the root and namespaces are kept. */
export function writeStory(doc, originalXml, ctx = {}) {
  let body = writeBlocks(doc, ctx);
  if (doc.lastChild && doc.lastChild.type.name === 'table') body += '<w:p/>';
  const open = /<w:(hdr|ftr)\b[^>]*?(\/?)>/.exec(originalXml);
  if (!open) throw new Error('not a header or footer part');
  const start = open.index + open[0].length;
  const end = open[2] ? start : originalXml.lastIndexOf(`</w:${open[1]}>`);
  const head = open[2] ? originalXml.slice(0, open.index) + open[0].replace(/\/>$/, '>') : originalXml.slice(0, start);
  return head + body + (open[2] ? `</w:${open[1]}>` : originalXml.slice(end));
}
