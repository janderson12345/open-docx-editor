// word/document.xml -> ProseMirror JSON.
// Recognised structure becomes editable nodes and marks; everything else is
// kept as raw XML (opaque nodes, `extra` attrs) so it can be written back.
import { NS, kids, kid, attr, onOff, outerXml, nsDecls, textOf, parseXml, childNamed, elementKids } from './xml.js';
import { readPPr, hexColor } from './props.js';
import { EMU_PER_PX, TWIPS_PER_PX, emuToPx, twipsToPx } from './units.js';

// w:pPr children we model as attrs; all others go to `extra`.
const PPR_MODELLED = new Set(['pStyle', 'numPr', 'jc']);
// w:rPr children we model as marks; all others go to the `rextra` mark.
const RPR_MODELLED = new Set(['rStyle', 'rFonts', 'b', 'i', 'strike', 'color', 'sz', 'highlight', 'u', 'vertAlign']);
// Paragraph-level markers with no visible content.
export const ZERO_WIDTH = new Set(['bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd', 'permStart', 'permEnd', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd']);

export function readDocument(docXml, { resolveImage = () => null, resolveLink = () => null, resolveChart = () => null, resolveShape = () => null } = {}) {
  sdtGroup = 0; // field keys are per document, so reading is repeatable
  const root = docXml.documentElement;
  const rootNs = nsDecls(root);
  const body = kid(root, 'body');
  if (!body) throw new Error('document.xml has no w:body');
  const ctx = { rootNs, resolveImage, resolveLink, resolveChart, resolveShape, xml: (el) => outerXml(el, rootNs) };

  let sectPr = null;
  const bodyKids = kids(body);
  // Each section's page, for drawings sized as a percentage of it (wp14:sizeRelH / sizeRelV).
  const sects = bodyKids.map((el) => (el.localName === 'sectPr' ? el : el.localName === 'p' ? kid(kid(el, 'pPr'), 'sectPr') : null)).filter(Boolean);
  ctx.pages = sects.map(pageBox);
  ctx.sect = 0;
  if (bodyKids.length && bodyKids[bodyKids.length - 1].localName === 'sectPr') sectPr = ctx.xml(bodyKids.pop());
  const content = readBlocks(bodyKids, ctx);
  if (!content.length) content.push({ type: 'paragraph' });
  stackDrawings(content);
  return { type: 'doc', attrs: { sectPr }, content };
}

/**
 * Floating drawings out of the flow (no wrapping) stacked as in Word: by their drawing order (wp:anchor relativeHeight,
 * higher on top), in front of the text or behind it. The browser stacks ones of the same z-index in document order,
 * which is Word's too when the drawing order follows the document; otherwise each gets a z-index by its rank, a picture
 * then staying under a text box drawn over it later in the document's history but earlier in its text.
 */
function stackDrawings(content) {
  const front = [], behind = [];
  const walk = (nodes) => {
    for (const n of nodes || []) {
      const z = n.type === 'image' && /z-index:(-?1)\b/.exec(n.attrs?.place || '');
      const rh = z && /<wp:anchor\b[^>]*\brelativeHeight="(\d+)"/.exec(n.attrs.xml || '');
      if (rh) (z[1] === '1' ? front : behind).push([Number(rh[1]), n]);
      walk(n.content);
    }
  };
  walk(content);
  for (const [list, sign] of [[front, 1], [behind, -1]]) {
    const sorted = [...list].sort((a, b) => a[0] - b[0]);
    if (sorted.every((e, i) => e === list[i])) continue; // already in drawing order
    // Front: 1 (lowest) upwards; behind: -1 (highest, nearest the text) downwards.
    sorted.forEach(([, n], i) => {
      const z = sign > 0 ? 1 + i : -(sorted.length - i);
      n.attrs.place = n.attrs.place.replace(/z-index:-?1\b/, `z-index:${z}`);
    });
  }
}

/** A header or footer part (w:hdr / w:ftr) as a doc. */
export function readStory(xmlDoc, { resolveImage = () => null, resolveLink = () => null, resolveChart = () => null, resolveShape = () => null } = {}, sectPrXml = null) {
  const root = xmlDoc.documentElement;
  const rootNs = nsDecls(root);
  const ctx = { rootNs, resolveImage, resolveLink, resolveChart, resolveShape, xml: (el) => outerXml(el, rootNs) };
  // (A story read for the body (a shape moved or resized there): its drawings placed and sized on the body's page.)
  if (sectPrXml) { ctx.pages = [pageBox(parseXml(sectPrXml).documentElement)]; ctx.sect = 0; }
  const content = readBlocks(kids(root), ctx);
  stackDrawings(content);
  return { type: 'doc', content: content.length ? content : [{ type: 'paragraph' }] };
}

// Fragments saved inside tracked formatting changes (w:rPrChange, w:pPrChange), read back
// when such a change is rejected.
const FRAGMENT_NS = `xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`;
function fragment(xml) {
  const root = parseXml(`<w:root ${FRAGMENT_NS}>${xml}</w:root>`).documentElement;
  const ctx = { rootNs: nsDecls(root), resolveImage: () => null, resolveLink: () => null, resolveChart: () => null, resolveShape: () => null };
  ctx.xml = (el) => outerXml(el, ctx.rootNs);
  return { root, ctx };
}

/** Run marks (as JSON) for a w:rPr XML fragment. */
export function marksFromRPr(xml) {
  const { root, ctx } = fragment(xml || '');
  return rPrMarks(kid(root, 'rPr'), ctx);
}

/** Paragraph attrs (styleId, align, numId, ilvl, extra, view) for a w:pPr XML fragment. */
export function paragraphAttrsFromPPr(xml) {
  const { root, ctx } = fragment(`<w:p>${xml || ''}</w:p>`);
  return readParagraph(kid(root, 'p'), ctx).attrs;
}

let sdtGroup = 0;
/**
 * A block-level content control (w:sdt) around paragraphs and tables: its blocks are
 * read as usual and each carries the wrapper in `sdt`, so the writer can put it back.
 * Returns false (keep it opaque) when it holds anything else or another control.
 */
function sdtBlocks(el, ctx, out) {
  const content = kid(el, 'sdtContent');
  if (!content) return false;
  const inner = kids(content);
  // A control that only wraps another (a page-number gallery around a footer's table): one wrapper of both, written back as it was.
  if (inner.length === 1 && inner[0].namespaceURI === NS.w && inner[0].localName === 'sdt') {
    const sub = [];
    if (!sdtBlocks(inner[0], ctx, sub)) return false;
    const head = kids(el).filter((c) => c !== content).map((c) => ctx.xml(c)).join('');
    for (const b of sub) {
      const s = b.attrs.sdt;
      out.push({ ...b, attrs: { ...b.attrs, sdt: { open: `<w:sdt>${head}<w:sdtContent>${s.open}`, close: `${s.close}</w:sdtContent></w:sdt>`, g: s.g } } });
    }
    return true;
  }
  if (!inner.length || !inner.every((c) => c.namespaceURI === NS.w && (c.localName === 'p' || c.localName === 'tbl'))) return false;
  const blocks = readBlocks(inner, ctx);
  if (blocks.some((b) => b.attrs?.sdt)) return false;
  const head = kids(el).filter((c) => c !== content).map((c) => ctx.xml(c)).join('');
  const sdt = { open: `<w:sdt>${head}<w:sdtContent>`, close: '</w:sdtContent></w:sdt>', g: ++sdtGroup };
  for (const b of blocks) out.push({ ...b, attrs: { ...(b.attrs || {}), sdt } });
  return true;
}

function readBlocks(els, ctx) {
  const out = [];
  for (const el of els) {
    if (el.namespaceURI === NS.w && el.localName === 'p') {
      out.push(readParagraph(el, ctx));
      if (ctx.pages && kid(kid(el, 'pPr'), 'sectPr')) ctx.sect++; // the next paragraphs are in the next section
    }
    else if (el.namespaceURI === NS.w && el.localName === 'tbl') out.push(readTable(el, ctx));
    else if (el.namespaceURI === NS.w && el.localName === 'sdt' && sdtBlocks(el, ctx, out)) continue;
    else out.push({ type: 'opaque_block', attrs: { xml: ctx.xml(el), text: previewText(el), kind: el.localName } });
  }
  return out;
}

function previewText(el) {
  // Visible text of an unmodelled block, paragraph breaks kept.
  const ps = el.getElementsByTagNameNS(NS.w, 'p');
  if (!ps.length) return textOf(el).trim();
  return Array.from(ps).map((p) => Array.from(p.getElementsByTagNameNS(NS.w, 't')).map(textOf).join('')).join('\n');
}

// --- Paragraphs -----------------------------------------------------------

function readParagraph(p, ctx) {
  const pPr = kid(p, 'pPr');
  const attrs = { styleId: null, align: null, numId: null, ilvl: 0, extra: [], view: {} };
  if (pPr) {
    const props = readPPr(pPr);
    attrs.styleId = props.styleId ?? null;
    attrs.align = props.align ?? null;
    if (props.numId != null) {
      attrs.numId = props.numId;
      attrs.ilvl = props.ilvl;
    }
    for (const c of kids(pPr)) if (!(c.namespaceURI === NS.w && PPR_MODELLED.has(c.localName))) attrs.extra.push([c.localName, ctx.xml(c)]);
    for (const k of ['before', 'beforeAuto', 'after', 'afterAuto', 'line', 'lineRule', 'indLeft', 'indRight', 'indFirst', 'contextualSpacing', 'keepNext', 'keepLines', 'widowControl', 'pageBreakBefore', 'borders', 'shading', 'tabs', 'bidi']) if (props[k] != null) attrs.view[k] = props[k];
  }
  // Word's paragraph identities (w14:paraId, w14:textId) and revision-save ids (w:rsid*), written back as they were.
  const own = [...p.attributes].filter((a) => a.prefix !== 'xmlns' && a.name !== 'xmlns').map((a) => [a.name, a.value]);
  if (own.length) attrs.pattrs = own;
  const content = [];
  readInline(kids(p).filter((c) => c !== pPr), [], ctx, content);
  bandFloats(content);
  return content.length ? { type: 'paragraph', attrs, content } : { type: 'paragraph', attrs };
}

function readInline(els, marks, ctx, out) {
  for (const el of els) {
    const name = el.namespaceURI === NS.w ? el.localName : null;
    switch (name) {
      case 'r':
        readRun(el, marks, ctx, out);
        break;
      case 'sdt': {
        // An inline content control (form field): its runs are read normally, marked as belonging to it.
        const content = kid(el, 'sdtContent');
        if (!content || marks.some((m) => m.type === 'sdt')) { out.push(opaque(ctx.xml(el), textOf(el), 'sdt', marks)); break; }
        const pr = kid(el, 'sdtPr');
        const endPr = kid(el, 'sdtEndPr');
        const prXml = pr ? ctx.xml(pr) : '';
        const before = out.length;
        const field = { type: 'sdt', attrs: { key: ++sdtGroup, pr: prXml, endPr: endPr ? ctx.xml(endPr) : '', ...fieldInfo(prXml), phText: /<w:showingPlcHdr(?![^>]*w:val="(?:0|false)")/.test(prXml) ? textOf(content) : null } };
        readInline(kids(content), [...marks, field], ctx, out);
        if (out.length === before) { out.length = before; out.push(opaque(ctx.xml(el), '', 'sdt', marks)); }
        break;
      }
      case 'hyperlink': {
        const extra = {};
        for (const a of Array.from(el.attributes)) if (!/^(r:id|w:anchor|xmlns)/.test(a.name)) extra[a.name] = a.value;
        const rId = el.getAttributeNS(NS.r, 'id') || null;
        const link = { type: 'link', attrs: { rId, href: rId ? ctx.resolveLink(rId) : null, anchor: attr(el, 'anchor'), extra: Object.keys(extra).length ? extra : null } };
        readInline(kids(el), [...marks, link], ctx, out);
        break;
      }
      case 'ins':
      case 'del':
      case 'moveTo':
      case 'moveFrom': {
        // Moved text is a deletion where it was and an insertion where it went, flagged as a move.
        const move = name === 'moveTo' || name === 'moveFrom';
        const type = name === 'moveTo' ? 'ins' : name === 'moveFrom' ? 'del' : name;
        const rev = { type, attrs: { id: attr(el, 'id'), author: attr(el, 'author'), date: attr(el, 'date'), ...(move ? { move: true } : {}) } };
        readInline(kids(el), [...marks, rev], ctx, out);
        break;
      }
      case 'proofErr':
        break; // spell-check markers; Word regenerates them
      case 'smartTag':
      case 'customXml':
        readInline(kids(el).filter((c) => c.localName !== 'smartTagPr' && c.localName !== 'customXmlPr'), marks, ctx, out);
        break;
      default: {
        const kind = el.localName;
        const text = ZERO_WIDTH.has(kind) ? '' : textOf(el);
        out.push(opaque(ctx.xml(el), text, kind, marks));
      }
    }
  }
}

/** Name and kind of a content control from its w:sdtPr XML. */
export function fieldInfo(prXml) {
  const a = (re) => (re.exec(prXml) || [])[1] ?? null;
  const kind = /<w14:checkbox\b/.test(prXml) ? 'checkbox' : /<w:dropDownList\b/.test(prXml) ? 'dropdown' : /<w:comboBox\b/.test(prXml) ? 'combobox'
    : /<w:date\b/.test(prXml) ? 'date' : /<w:picture\b/.test(prXml) ? 'picture' : /<w:text\b/.test(prXml) ? 'text' : 'richText';
  return { tag: a(/<w:tag\b[^>]*w:val="([^"]*)"/), alias: a(/<w:alias\b[^>]*w:val="([^"]*)"/), kind };
}

function opaque(xml, text, kind, marks) {
  const node = { type: 'opaque_inline', attrs: { xml, text, kind } };
  const keep = marks.filter((m) => m.type === 'link' || m.type === 'ins' || m.type === 'del' || m.type === 'sdt');
  if (keep.length) node.marks = keep;
  return node;
}

function readRun(r, outer, ctx, out) {
  const rPr = kid(r, 'rPr');
  const marks = [...outer, ...rPrMarks(rPr, ctx)];
  const rPrXml = rPr ? ctx.xml(rPr) : '';
  const push = (node) => {
    if (marks.length) node.marks = marks;
    out.push(node);
  };
  for (const c of kids(r)) {
    if (c === rPr) continue;
    const name = c.namespaceURI === NS.w || c.localName === 'AlternateContent' ? c.localName : null;
    switch (name) {
      case 't':
      case 'delText': {
        // A literal tab inside w:t displays as a tab in Word.
        textOf(c).split('\t').forEach((t, i) => {
          if (i) push({ type: 'tab' });
          if (t) push({ type: 'text', text: t });
        });
        break;
      }
      case 'tab':
        push({ type: 'tab' });
        break;
      case 'br':
      case 'cr':
        push({ type: 'hard_break', attrs: { type: attr(c, 'type'), clear: attr(c, 'clear') } });
        break;
      case 'noBreakHyphen':
        push({ type: 'text', text: '‑' });
        break;
      case 'softHyphen':
        push({ type: 'text', text: '­' });
        break;
      case 'lastRenderedPageBreak':
        break; // layout cache written by Word; safe to drop
      case 'drawing':
      case 'pict':
      case 'object':
      case 'AlternateContent': {
        const img = readGraphic(c, ctx, rPrXml);
        if (img) {
          if (img.type === 'opaque_inline') img.attrs.xml = `<w:r>${rPrXml}${ctx.xml(c)}</w:r>`;
          const keep = marks.filter((m) => m.type === 'link' || m.type === 'ins' || m.type === 'del' || m.type === 'sdt');
          if (keep.length) img.marks = keep;
          out.push(img);
          break;
        }
      }
      // fall through: anything else stays opaque, showing its text
      default: {
        const xml = `<w:r>${rPrXml}${ctx.xml(c)}</w:r>`;
        // Graphics show only their text boxes' text (once: the fallback repeats it), never offsets or other values.
        const graphicText = (g) => g.getElementsByTagNameNS(NS.w, 't').length ? previewText(g) : '';
        const text = c.localName === 'instrText' || c.localName === 'delInstrText' || c.localName === 'fldChar' ? ''
          : c.localName === 'sym' ? symChar(c)
          : c.localName === 'AlternateContent' ? graphicText(byLocal(c, 'Choice')[0] || c)
          : ['drawing', 'pict', 'object'].includes(c.localName) ? graphicText(c) : textOf(c);
        out.push(opaque(xml, text, c.localName, outer));
      }
    }
  }
}

const byLocal = (el, name) => [...el.getElementsByTagName('*')].filter((e) => e.localName === name);

/**
 * A w:drawing, w:pict or mc:AlternateContent run child as an image node when we can draw it,
 * else as a blank box of its size (inline graphics) so the page lays out like Word's.
 * Returns null to fall back to a plain opaque node (floating text boxes keep showing their text).
 */
function readGraphic(el, ctx, rPrXml) {
  let drawing = el.localName === 'drawing' ? el : null;
  // An embedded OLE object (w:object) shows its VML preview picture, like w:pict.
  let pict = el.localName === 'pict' || el.localName === 'object' ? el : null;
  if (el.localName === 'AlternateContent') {
    const choice = byLocal(el, 'Choice')[0];
    const fallback = byLocal(el, 'Fallback')[0];
    drawing = (choice && byLocal(choice, 'drawing')[0]) || (fallback && byLocal(fallback, 'drawing')[0]) || null;
    pict = (fallback && byLocal(fallback, 'pict')[0]) || null;
    if (!drawing && !pict) return null;
  }
  if (drawing && ctx.pages?.length) relativeSize(drawing, ctx.pages[Math.min(ctx.sect, ctx.pages.length - 1)]);
  const xml = `<w:r>${rPrXml}${ctx.xml(el)}</w:r>`;
  if (drawing) {
    const img = readDrawing(drawing, ctx, rPrXml);
    if (img) return (img.attrs.xml = xml), img;
  }
  if (pict) {
    // A horizontal rule (Insert > Horizontal Line, and in documents exported from other editors): a filled bar across the column.
    const hr = byLocal(pict, 'rect').find((r) => r.getAttributeNS('urn:schemas-microsoft-com:office:office', 'hr') === 't');
    if (hr) {
      const h = /height:\s*([\d.]+)pt/.exec(hr.getAttribute('style') || '');
      const fill = /^#?[0-9a-f]{6}$/i.test(hr.getAttribute('fillcolor') || '') ? hr.getAttribute('fillcolor').replace('#', '') : 'a0a0a0';
      const src = `data:image/svg+xml,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><rect width="1" height="1" fill="#${fill}"/></svg>`)}`;
      // Wider than any column: the image's max-width trims it to the column, the height stays.
      return { type: 'image', attrs: { xml, src, width: 4000, height: Math.max(1, Math.round((h ? Number(h[1]) : 1.5) * 4 / 3)), alt: '', place: 'vertical-align:baseline' } };
    }
    const wm = watermark(pict, ctx);
    if (wm) return { type: 'image', attrs: { xml, ...wm } };
    const data = byLocal(pict, 'imagedata').find((d) => d.getAttributeNS(NS.r, 'id'));
    const shape = data && data.parentNode;
    const css = (shape && shape.getAttribute && shape.getAttribute('style')) || '';
    const len = (k) => {
      const m = new RegExp(`(?:^|;)\\s*${k}:\\s*([\\d.]+)(pt|in|px)?`).exec(css);
      return m ? Math.round(Number(m[1]) * ({ in: 96, px: 1 }[m[2]] || 4 / 3)) : null;
    };
    const src = data && ctx.resolveImage(data.getAttributeNS(NS.r, 'id'));
    const alt = data && (data.getAttributeNS('urn:schemas-microsoft-com:office:office', 'title') || '');
    if (src && !/position:\s*absolute/.test(css)) return { type: 'image', attrs: { xml, src, width: len('width'), height: len('height'), alt } };
    // A floating VML picture (older Word files): placed from the column (or page) by its margins, like a wrapNone drawing.
    if (src) {
      const off = (k) => { const m = new RegExp(`(?:^|;)\\s*${k}:\\s*(-?[\\d.]+)(pt|in|px)?`).exec(css); return m ? Math.round(Number(m[1]) * ({ in: 96, px: 1 }[m[2]] || 4 / 3)) : 0; };
      const x = off('margin-left');
      const fromPage = /mso-position-horizontal-relative:\s*page/.test(css);
      const z = (/z-index:\s*(-?\d+)/.exec(css) || [])[1];
      const left = fromPage ? `left:${x}px;` : `left:calc(var(--odx-page-left, 96px) + ${x}px);`;
      const place = `position:absolute;${left}margin-top:${off('margin-top')}px;z-index:${z && Number(z) < 0 ? -1 : 1}`;
      return { type: 'image', attrs: { xml, src, width: len('width'), height: len('height'), alt, place } };
    }
  }
  const inline = drawing && byLocal(drawing, 'inline')[0];
  const extent = inline && byLocal(inline, 'extent')[0];
  if (!extent) return null;
  const px = (v) => Math.round(emuToPx(v));
  return { type: 'opaque_inline', attrs: { xml, text: '', kind: 'drawing', width: px(extent.getAttribute('cx')), height: px(extent.getAttribute('cy')) } };
}

/**
 * Word's watermark (Design > Watermark): a VML shape in a header, floating behind the text. Text (WordArt: its string
 * fitted to the shape, in its colour and opacity) is drawn into a picture; a picture is washed out (its gain).
 * Returned as an image's attrs, `--odx-wm` in its style: pages.js puts it on each page, centred on the margins (or
 * where the shape says), behind the text. Null for any other shape.
 */
function watermark(pict, ctx) {
  const shape = byLocal(pict, 'shape').find((s) => /position:\s*absolute/.test(s.getAttribute('style') || '') && (/water\s*mark/i.test(s.getAttribute('id') || '') || byLocal(s, 'textpath').some((t) => t.getAttribute('string'))));
  if (!shape) return null;
  const css = shape.getAttribute('style') || '';
  const len = (k) => { const m = new RegExp(`(?:^|;)\\s*${k}:\\s*(-?[\\d.]+)(pt|in|px)?`).exec(css); return m ? Number(m[1]) * ({ in: 96, px: 1 }[m[2]] || 4 / 3) : null; };
  const w = len('width'), h = len('height');
  if (!(w > 0 && h > 0 && w < 5000 && h < 5000)) return null; // (a watermark is about a page in size: anything far past it isn't one)
  const frac = (v) => (v == null || v === '' ? null : /f$/.test(v) ? parseInt(v, 10) / 65536 : Number(v));
  const at = (axis) => {
    const pos = new RegExp(`mso-position-${axis}:\\s*(\\w+)`).exec(css)?.[1];
    const rel = new RegExp(`mso-position-${axis}-relative:\\s*(\\w+)`).exec(css)?.[1] || 'margin';
    return `--odx-wm-${axis[0]}:${pos === 'center' ? 'center' : len(axis === 'horizontal' ? 'margin-left' : 'margin-top') || 0};--odx-wm-${axis[0]}rel:${rel === 'page' ? 'page' : 'margin'}`;
  };
  const rot = Number(/rotation:\s*(-?[\d.]+)/.exec(css)?.[1] || 0);
  let place = `position:absolute;left:0;top:0;pointer-events:none;--odx-wm:1;${at('horizontal')};${at('vertical')}${rot ? `;transform:rotate(${rot}deg)` : ''}`;
  const tp = byLocal(shape, 'textpath').find((t) => t.getAttribute('string'));
  if (tp) {
    const text = tp.getAttribute('string');
    const fill = byLocal(shape, 'fill')[0];
    const src = textPicture(text, /font-family:\s*"?([^";]+)/.exec(tp.getAttribute('style') || '')?.[1] || 'Calibri', /font-weight:\s*bold/.test(tp.getAttribute('style') || ''), (shape.getAttribute('fillcolor') || 'silver').replace(/ \[\d+\]$/, ''), frac(fill?.getAttribute('opacity')) ?? 1, w, h);
    return src ? { src, width: Math.round(w), height: Math.round(h), alt: '', place } : null;
  }
  const data = byLocal(shape, 'imagedata').find((d) => d.getAttributeNS(NS.r, 'id'));
  const src = data && ctx.resolveImage(data.getAttributeNS(NS.r, 'id'));
  if (!src) return null;
  const gain = frac(data.getAttribute('gain'));
  if (gain != null && gain < 1) place += `;opacity:${+gain.toFixed(3)}`; // (washed out: Word's 30% gain)
  return { src, width: Math.round(w), height: Math.round(h), alt: '', place };
}

/** A watermark's text as a PNG data URL, fitted to w x h (px) and centred in it. */
function textPicture(text, family, bold, color, opacity, w, h) {
  if (typeof document === 'undefined') return null;
  const c = document.createElement('canvas');
  const g = c.getContext('2d');
  if (!g) return null;
  const font = `${bold ? 'bold ' : ''}200px "${family}", sans-serif`;
  g.font = font;
  const m = g.measureText(text);
  const bw = m.actualBoundingBoxLeft + m.actualBoundingBoxRight, bh = m.actualBoundingBoxAscent + m.actualBoundingBoxDescent;
  if (!(bw > 0 && bh > 0)) return null;
  const scale = Math.min(2, 2000 / Math.max(w, h));
  c.width = Math.max(1, Math.round(w * scale));
  c.height = Math.max(1, Math.round(h * scale));
  // (Its advance across the shape's width, the font's line (ascent and descent) down its height, as Word draws it.)
  const fh = m.fontBoundingBoxAscent + m.fontBoundingBoxDescent;
  g.setTransform(c.width / m.width, 0, 0, c.height / fh, 0, 0);
  g.font = font;
  g.fillStyle = color;
  g.globalAlpha = opacity;
  g.fillText(text, 0, m.fontBoundingBoxAscent);
  return c.toDataURL('image/png');
}

function symChar(sym) {
  const code = parseInt(attr(sym, 'char') || '0', 16);
  return code ? String.fromCharCode(code >= 0xf000 ? code - 0xf000 : code) : '';
}

function rPrMarks(rPr, ctx) {
  if (!rPr) return [];
  const m = [];
  const extra = [];
  for (const c of kids(rPr)) {
    const n = c.namespaceURI === NS.w ? c.localName : null;
    if (!RPR_MODELLED.has(n)) {
      extra.push([c.localName, ctx.xml(c)]);
      continue;
    }
    const val = attr(c, 'val');
    switch (n) {
      case 'rStyle': m.push({ type: 'cstyle', attrs: { val } }); break;
      case 'rFonts': {
        const f = attr(c, 'ascii') || attr(c, 'hAnsi');
        // Theme-font references round-trip raw; explicit fonts become an editable mark
        // that remembers its original element (cs/eastAsia fonts, hints).
        if (f && !attr(c, 'asciiTheme') && !attr(c, 'hAnsiTheme')) m.push({ type: 'font', attrs: { val: f, xml: ctx.xml(c) } });
        else extra.push([n, ctx.xml(c)]);
        break;
      }
      case 'b': m.push({ type: 'bold', attrs: { on: onOff(c) } }); break;
      case 'i': m.push({ type: 'italic', attrs: { on: onOff(c) } }); break;
      case 'strike': m.push({ type: 'strike', attrs: { on: onOff(c) } }); break;
      case 'color': if (hexColor(val) && c.attributes.length === 1) m.push({ type: 'color', attrs: { val } }); else extra.push([n, ctx.xml(c)]); break;
      case 'sz': m.push({ type: 'size', attrs: { val: Number(val) } }); break;
      case 'highlight': m.push({ type: 'highlight', attrs: { val } }); break;
      case 'u': if (c.attributes.length <= 1) m.push({ type: 'underline', attrs: { val: val || 'single' } }); else extra.push([n, ctx.xml(c)]); break;
      case 'vertAlign': m.push({ type: 'vertAlign', attrs: { val } }); break;
    }
  }
  if (extra.length) m.push({ type: 'rextra', attrs: { items: extra } });
  return m;
}

// --- Drawings -------------------------------------------------------------

/**
 * CSS placing a floating (wp:anchor) drawing roughly where Word puts it: on its own band for
 * "top and bottom" wrapping, floated for square/tight wrapping, out of the flow for none.
 * Horizontal offsets from the page edge use --odx-page-left (the left margin).
 */
function anchorCss(drawing, page) {
  const anchor = [...drawing.getElementsByTagName('*')].find((e) => e.localName === 'anchor');
  if (!anchor) return '';
  const px = (v) => Math.round(emuToPx(v || 0));
  // (A position given as a percentage (Word 2010's wp14:pctPos*) comes in an mc:AlternateContent whose fallback has it
  // in EMUs for the page Word saved it on: that one.)
  const alt = (k) => {
    for (const ac of elementKids(anchor).filter((n) => n.localName === 'AlternateContent')) {
      const fb = childNamed(ac, 'Fallback');
      const el = fb && childNamed(fb, k);
      if (el) return el;
    }
    return null;
  };
  const posH = childNamed(anchor, 'positionH') || alt('positionH');
  const posV = childNamed(anchor, 'positionV') || alt('positionV');
  const hRel = posH ? posH.getAttribute('relativeFrom') : 'column';
  const hAlign = childNamed(posH, 'align')?.textContent;
  const x = px(childNamed(posH, 'posOffset')?.textContent);
  const vRel = posV ? posV.getAttribute('relativeFrom') : 'paragraph';
  // The offset places the box with its effect extent (room for a shadow or outline): the shape itself sits that much in
  // (sample S078's cover box, 1px lower in Word's PDF than its anchor paragraph's top).
  const fx = childNamed(anchor, 'effectExtent');
  const fxT = emuToPx(fx?.getAttribute('t') || 0);
  const y = ['paragraph', 'line'].includes(vRel) ? px(childNamed(posV, 'posOffset')?.textContent) + Math.round(fxT * 100) / 100 : 0;
  let wrap = [...anchor.childNodes].find((n) => n.localName && n.localName.startsWith('wrap'))?.localName || 'wrapNone';
  // Wrapped round, but as wide as the column: no text fits beside it, so it takes its own band, as top-and-bottom (and,
  // a block, it can overlap another float, as Word lets it: sample S095's full-width picture under the text box beside it).
  const extW = px(childNamed(anchor, 'extent')?.getAttribute('cx'));
  const band = page && /^wrap(Square|Tight|Through)$/.test(wrap) && extW >= page.w - page.left - page.right;
  if (band) wrap = 'wrapTopAndBottom';
  const fromPage = hRel === 'page';
  const inFlowX = fromPage ? `calc(${x}px - var(--odx-page-left, 96px))` : `${x}px`;
  if (wrap === 'wrapTopAndBottom') {
    // A paragraph-relative offset is from the paragraph's top above its space before (ftd-1.35MB's page 35 chart, 32.8px
    // under a paragraph spaced 14.9px before: 15px too low when measured from the text).
    const top = vRel === 'paragraph' && y ? `calc(${y}px - var(--odx-sb, 0pt))` : `${y}px`;
    if (hAlign === 'center') return `display:block;margin:${top} auto ${px(anchor.getAttribute('distB'))}px`;
    if (hAlign === 'right') return `display:block;margin:${top} 0 ${px(anchor.getAttribute('distB'))}px auto`;
    // (A wrapped-round one keeps its paragraph's own line, below it like any line Word wraps round it: sample S095's empty
    // anchor paragraph, a line under the picture.)
    return `display:block;margin-top:${top};margin-left:${inFlowX};margin-bottom:${px(anchor.getAttribute('distB'))}px${band ? ';--odx-band:1' : ''}`;
  }
  if (wrap === 'wrapNone') {
    // Centred or right-aligned: in the column (between the margins), or on the page (sample S078's text box,
    // centred in the column, had sat at the left margin).
    const ext = childNamed(anchor, 'extent');
    const w = ext ? px(ext.getAttribute('cx')) : 0;
    const span = fromPage ? '100%' : 'calc(100% - var(--odx-page-left, 96px) - var(--odx-page-right, 96px))';
    const from = fromPage ? '0px' : 'var(--odx-page-left, 96px)';
    const aligned = hAlign === 'center' ? `left:calc(${from} + (${span} - ${w}px) / 2);` : hAlign === 'right' ? `left:calc(${from} + ${span} - ${w}px);` : hAlign === 'left' && hRel === 'page' ? 'left:0px;' : '';
    const left = hAlign ? (['column', 'margin', 'page'].includes(hRel) ? aligned : '') : fromPage ? `left:${x}px;` : `left:calc(var(--odx-page-left, 96px) + ${x}px);`;
    // Placed against the page or its margins vertically: pagination puts it at that height on its paragraph's page
    // (pages.js placeOnPage).
    const onPage = ['page', 'margin', 'topMargin', 'bottomMargin'].includes(vRel)
      ? `;--odx-arel:${vRel};--odx-ay:${px(childNamed(posV, 'posOffset')?.textContent) + Math.round(fxT * 100) / 100}px;--odx-ayalign:${childNamed(posV, 'align')?.textContent || 'none'}` : '';
    return `position:absolute;${left}margin-top:${y}px;z-index:${anchor.getAttribute('behindDoc') === '1' ? -1 : 1}${onPage}`;
  }
  const right = hAlign === 'right' || hAlign === 'outside';
  // Placed against the page or its margins (not its paragraph): pagination lifts it to that page's first paragraph, at
  // that height (pages.js liftFloats). sample S125's arrows, at the top corners of the margins of the page their
  // paragraph is on, with that page's first lines wrapping between them.
  const lift = ['page', 'margin', 'topMargin'].includes(vRel)
    ? `;--odx-lrel:${vRel};--odx-ly:${px(childNamed(posV, 'posOffset')?.textContent)}px;--odx-lyalign:${childNamed(posV, 'align')?.textContent || 'none'}` : '';
  // (Aligned right, its own edge is at the margin: its wrap distance is on the text's side only. sample S125's arrow 12px
  // in from the margin otherwise.)
  // Tight and through wrapping go round the wrap polygon (in 21600ths of the picture), not the picture's box: the text
  // comes as much nearer on the side it's on (sample S116's letter: a picture 120.3px wide with its polygon to 21193, the
  // text 2.3px further left in Word's PDF, where "would" just fits).
  let trimSide = 0;
  let trimBottom = 0;
  const poly = /^wrap(Tight|Through)$/.test(wrap) ? [...anchor.getElementsByTagName('*')].find((n) => n.localName === 'wrapPolygon') : null;
  if (poly) {
    const pts = [...poly.childNodes].filter((n) => n.localName === 'start' || n.localName === 'lineTo').map((n) => [Number(n.getAttribute('x')), Number(n.getAttribute('y'))]);
    const ext = childNamed(anchor, 'extent');
    const w = emuToPx(ext?.getAttribute('cx') || 0), h = emuToPx(ext?.getAttribute('cy') || 0);
    if (pts.length > 2 && w && h) {
      const xs = pts.map((q) => q[0]), ys = pts.map((q) => q[1]);
      trimSide = right ? (Math.max(0, Math.min(...xs)) / 21600) * w : Math.max(0, (21600 - Math.max(...xs)) / 21600) * w;
      trimBottom = Math.max(0, (21600 - Math.max(...ys)) / 21600) * h;
    }
  }
  const r2 = (v) => +v.toFixed(2);
  const mr = (hAlign === 'right' ? 0 : px(anchor.getAttribute('distR'))) - (right ? 0 : trimSide);
  const ml = right ? `${r2(px(anchor.getAttribute('distL')) - trimSide)}px` : inFlowX;
  return `float:${right ? 'right' : 'left'};margin:${y}px ${r2(mr)}px ${r2(px(anchor.getAttribute('distB')) - trimBottom)}px ${ml}${lift}`;
}

/**
 * A paragraph with a column-wide picture taking its own band (anchorCss): another float of its own lying within that
 * band has no text beside it in Word, which lets the two overlap. The browser can't overlap a float and a picture block,
 * and would put the band below it (sample S095: the speeds table's text box over a full-width picture), so such a float
 * is drawn out of the flow at its place.
 */
function bandFloats(content) {
  const imgs = content.filter((n) => n.type === 'image' && n.attrs?.place);
  const band = imgs.find((n) => /^display:block/.test(n.attrs.place));
  if (!band) return;
  const top = (pl) => parseFloat((/margin(?:-top)?:\s*(?:calc\()?(-?[\d.]+)px/.exec(pl) || [])[1]) || 0;
  const b0 = top(band.attrs.place);
  const b1 = b0 + Number(band.attrs.height || 0);
  for (const n of imgs) {
    if (n === band || !/^float:(left|right)/.test(n.attrs.place)) continue;
    const y = top(n.attrs.place.replace(/^float:\w+;/, ''));
    if (y < b0 || y + Number(n.attrs.height || 0) > b1 + 1) continue;
    const right = /^float:right/.test(n.attrs.place);
    const x = right ? `right:var(--odx-page-right, 96px)` : `left:var(--odx-page-left, 96px)`;
    n.attrs.place = `position:absolute;${x};margin-top:${y}px;z-index:1`;
  }
}

/** A section's page size and margins in px. */
function pageBox(sectPr) {
  const at = (el, k, d) => { const v = el && el.getAttributeNS(NS.w, k); return v ? twipsToPx(v) : d; };
  const sz = kid(sectPr, 'pgSz');
  const mar = kid(sectPr, 'pgMar');
  return { w: at(sz, 'w', 816), h: at(sz, 'h', 1056), left: at(mar, 'left', 96) + at(mar, 'gutter', 0), right: at(mar, 'right', 96), top: Math.abs(at(mar, 'top', 96)), bottom: Math.abs(at(mar, 'bottom', 96)) };
}

/**
 * A floating drawing sized as a percentage of the page or its margins (Word 2010's wp14:sizeRelH / sizeRelV)
 * gets that size written into its frame (and its shape's or group's transform), so it is drawn, laid out
 * and saved at the size Word shows. Word works the size out again from the percentage when it opens the file.
 */
function relativeSize(drawing, page) {
  const anchor = [...drawing.childNodes].find((n) => n.localName === 'anchor');
  const extent = anchor && [...anchor.childNodes].find((n) => n.localName === 'extent');
  if (!extent || !page) return;
  const pct = (name, val) => {
    const rel = [...anchor.childNodes].find((n) => n.localName === name);
    const v = rel && [...rel.childNodes].find((n) => n.localName === val);
    const n = v ? Number(v.textContent) : 0;
    return n > 0 ? { n: n / 100000, from: rel.getAttribute('relativeFrom') } : null;
  };
  const H = { page: page.w, margin: page.w - page.left - page.right, leftMargin: page.left, rightMargin: page.right, insideMargin: page.left, outsideMargin: page.right };
  const V = { page: page.h, margin: page.h - page.top - page.bottom, topMargin: page.top, bottomMargin: page.bottom, insideMargin: page.top, outsideMargin: page.bottom };
  const h = pct('sizeRelH', 'pctWidth');
  const v = pct('sizeRelV', 'pctHeight');
  const cx0 = Number(extent.getAttribute('cx'));
  const cy0 = Number(extent.getAttribute('cy'));
  const cx = h && H[h.from] != null ? Math.round(h.n * H[h.from] * EMU_PER_PX) : cx0;
  const cy = v && V[v.from] != null ? Math.round(v.n * V[v.from] * EMU_PER_PX) : cy0;
  if (cx === cx0 && cy === cy0) return;
  extent.setAttribute('cx', String(cx));
  extent.setAttribute('cy', String(cy));
  // The shape (or the group's frame; its children follow through chOff/chExt) takes the same size.
  const frame = [...drawing.getElementsByTagName('*')].find((n) => (n.localName === 'spPr' || n.localName === 'grpSpPr') && n.parentNode?.parentNode?.localName === 'graphicData');
  const ext = frame && [...([...frame.childNodes].find((n) => n.localName === 'xfrm')?.childNodes || [])].find((n) => n.localName === 'ext');
  if (ext) { ext.setAttribute('cx', String(cx)); ext.setAttribute('cy', String(cy)); }
}

/**
 * A picture's cropping (a:srcRect beside its a:blip, in thousandths of a percent of the picture from each edge; negative
 * adds a margin) as fractions { l, t, r, b }, or null when it is shown whole.
 */
function cropOf(blip) {
  const rect = [...(blip.parentNode?.childNodes || [])].find((n) => n.localName === 'srcRect');
  if (!rect) return null;
  const c = Object.fromEntries(['l', 't', 'r', 'b'].map((k) => [k, (Number(rect.getAttribute(k)) || 0) / 100000]));
  return (c.l || c.t || c.r || c.b) && c.l + c.r < 1 && c.t + c.b < 1 ? c : null;
}

function readDrawing(drawing, ctx, rPrXml) {
  const blip = drawing.getElementsByTagNameNS(NS.a, 'blip')[0];
  const extent = drawing.getElementsByTagNameNS(NS.wp, 'extent')[0];
  const docPr = drawing.getElementsByTagNameNS(NS.wp, 'docPr')[0];
  // (To a hundredth of a pixel: sample4's 688.4px picture, rounded to 688, put page 87's last line on it, where Word
  // has it on the next. +30 lines within 1px over the samples; rounding anchored ones still gave 6 fewer.)
  const emu = (v) => Math.round(emuToPx(v || 0) * 100) / 100;
  const width = extent ? emu(extent.getAttribute('cx')) : null;
  const height = extent ? emu(extent.getAttribute('cy')) : null;
  // An inline picture's effect extent (room for a shadow or glow) also takes space on its line, as in Word.
  const fx = extent?.parentNode?.localName === 'inline' && [...extent.parentNode.childNodes].find((n) => n.localName === 'effectExtent');
  // Rounded as a whole, like the picture itself; its top part above the picture (an outline all round: the picture sits
  // that much lower), the rest below.
  const fv = fx ? ['t', 'b'].reduce((a, k) => a + Math.max(0, Number(fx.getAttribute(k)) || 0), 0) : 0;
  const all = fv ? emu(Number(extent.getAttribute('cy')) + fv) - height : 0;
  const above = fx ? Math.min(all, emu(Math.max(0, Number(fx.getAttribute('t')) || 0))) : 0;
  const below = all - above;
  // (And its sides: an outline's 2px each side puts the picture 2px right of the line's start, doc 16's screenshots.)
  const side = (k) => (fx ? Math.round(Math.max(0, Number(fx.getAttribute(k)) || 0) / EMU_PER_PX * 100) / 100 : 0);
  const left = side('l'), right = side('r');
  const effect = [below > 0 && `--odx-fx:${below}px`, above > 0 && `--odx-fxt:${above}px`, left > 0 && `--odx-fxl:${left}px`, right > 0 && `--odx-fxr:${right}px`].filter(Boolean).join(';');
  let src = null;
  const uri = [...drawing.getElementsByTagName('*')].find((e) => e.localName === 'graphicData')?.getAttribute('uri') || '';
  if (/wordprocessing(Group|Shape)$/.test(uri)) {
    src = ctx.resolveShape ? ctx.resolveShape(drawing, width, height) : null;
    if (!src) return null;
  } else if (blip) {
    // An SVG picture (Word 2016+): the SVG itself (asvg:svgBlip in the blip's extensions), the PNG beside it a fallback.
    const svg = [...blip.getElementsByTagName('*')].find((e) => e.localName === 'svgBlip');
    const rId = blip.getAttributeNS(NS.r, 'embed');
    src = (svg && ctx.resolveImage(svg.getAttributeNS(NS.r, 'embed'), cropOf(blip))) || (rId ? ctx.resolveImage(rId, cropOf(blip)) : null);
  } else {
    const chart = [...drawing.getElementsByTagName('*')].find((e) => e.localName === 'chart' && e.getAttributeNS(NS.r, 'id'));
    src = chart && ctx.resolveChart ? ctx.resolveChart(chart.getAttributeNS(NS.r, 'id'), width, height) : null;
    if (!src) return null;
  }
  return {
    type: 'image',
    attrs: {
      xml: `<w:r>${rPrXml}${ctx.xml(drawing)}</w:r>`,
      src,
      width,
      height,
      place: anchorCss(drawing, ctx.pages?.[Math.min(ctx.sect, ctx.pages.length - 1)]) || effect,
      alt: docPr ? docPr.getAttribute('descr') || docPr.getAttribute('name') || '' : '',
      // A picture that is a link (a:hlinkClick on its wp:docPr: a logo linking to a web site), followed on Ctrl+click.
      ...(() => {
        const h = docPr && elementKids(docPr).find((n) => n.localName === 'hlinkClick');
        const href = h && ctx.resolveLink(h.getAttributeNS(NS.r, 'id'));
        return href ? { href } : {};
      })(),
    },
  };
}

// --- Tables ---------------------------------------------------------------

// Word widens autofit cells to fit pictures, so a picture wider than its cell's w:tcW means the grid is right.
function picturesWiderThan(tbl, ws) {
  for (const tr of kids(tbl, 'tr')) {
    kids(tr, 'tc').forEach((tc, i) => {
      for (const ext of [...tc.getElementsByTagName('*')].filter((e) => e.localName === 'extent' && e.parentNode?.localName === 'inline')) {
        if (emuToPx(ext.getAttribute('cx')) > (ws[i] ?? Infinity)) throw picturesWiderThan;
      }
    });
  }
  return false;
}

function readTable(tbl, ctx) {
  const tblPr = kid(tbl, 'tblPr');
  const tblGrid = kid(tbl, 'tblGrid');
  // In px to the hundredth: a whole-px column can be a third of a px narrower than Word's, enough to wrap a word that just fits.
  // Word reads a width given with a fraction ("2486.911", from some generators) as whole twips: 0.06px narrower can wrap.
  const px = (tw) => Math.round(Math.trunc(Number(tw || 0)) / TWIPS_PER_PX * 100) / 100;
  let grid = kids(tblGrid, 'gridCol').map((g) => px(attr(g, 'w')));
  // An autofit table whose cells all give their own widths (w:tcW) is laid out by those, as Word does,
  // when the grid disagrees (a grid wider than the cells add up to, say).
  // (Not when the table's own width, w:tblW, is the grid's: Word keeps the grid then.)
  const tblW = kid(tblPr, 'tblW');
  const gridSum = kids(tblGrid, 'gridCol').reduce((s, g) => s + Number(attr(g, 'w') || 0), 0);
  const widthIsGrid = attr(tblW, 'type') === 'dxa' && Math.abs(Number(attr(tblW, 'w') || 0) - gridSum) <= 15;
  // (Nor when it's a share of the page, w:type="pct": that sets its width, and the grid its columns; sample S125's 70%
  // table with a 4788-twip cell over a 3488-twip column.)
  const share = attr(tblW, 'type') === 'pct' && Number(attr(tblW, 'w') || 0) > 0;
  if (attr(kid(tblPr, 'tblLayout'), 'type') !== 'fixed' && !widthIsGrid && !share) {
    const widthsOf = (tr) => {
      const tcs = tr ? kids(tr, 'tc') : [];
      const ws = tcs.map((tc) => {
        const pr = kid(tc, 'tcPr');
        const w = kid(pr, 'tcW');
        return !kid(pr, 'gridSpan') && w && attr(w, 'type') === 'dxa' ? px(attr(w, 'w')) : null;
      });
      return tcs.length === grid.length && !kid(kid(tr, 'trPr'), 'gridBefore') && ws.every((w) => w > 0) ? ws : null;
    };
    // The first row's, or (its cells merged across columns: a heading row over the columns) the first row with a cell
    // in each column (generators' placeholder grids, 100-twip columns under a merged title row, laid out by those).
    const trs0 = kids(tbl, 'tr');
    const next = () => {
      const w = trs0.length > 1 && kids(trs0[0], 'tc').some((tc) => kid(kid(tc, 'tcPr'), 'gridSpan')) ? trs0.slice(1).map(widthsOf).find(Boolean) : null;
      // (Only for a grid that is plainly a placeholder, under half what the cells give: a real grid is kept.)
      return w && grid.reduce((a, b) => a + b, 0) * 2 < w.reduce((a, b) => a + b, 0) ? w : null;
    };
    const ws = widthsOf(trs0[0]) || next();
    if (ws && ws.some((w, i) => Math.abs(w - grid[i]) > 2) && !(() => { try { return picturesWiderThan(tbl, ws); } catch { return true; } })()) grid = ws;
  }
  const rows = [];
  // vertical-merge bookkeeping: grid column -> the cell that started the merge
  const open = new Map();
  // Rows, including rows inside row-level content controls (w:sdt), which keep their wrapper in `sdt`.
  const trs = [];
  for (const c of kids(tbl)) {
    if (c.namespaceURI !== NS.w) continue;
    if (c.localName === 'tr') trs.push([c, null]);
    else if (c.localName === 'sdt') {
      const content = kid(c, 'sdtContent');
      const head = kids(c).filter((x) => x !== content).map((x) => ctx.xml(x)).join('');
      const sdt = { open: `<w:sdt>${head}<w:sdtContent>`, close: '</w:sdtContent></w:sdt>', g: ++sdtGroup };
      for (const tr of content ? kids(content, 'tr') : []) trs.push([tr, sdt]);
    }
  }
  for (const [tr, rowSdt] of trs) {
    const trPr = kid(tr, 'trPr');
    const tblPrEx = kid(tr, 'tblPrEx');
    const row = { type: 'table_row', attrs: { trPr: trPr ? ctx.xml(trPr) : null, tblPrEx: tblPrEx ? ctx.xml(tblPrEx) : null, sdt: rowSdt }, content: [] };
    let col = 0;
    const gridBefore = Number(attr(kid(trPr, 'gridBefore'), 'val') || 0);
    const gridAfter = Number(attr(kid(trPr, 'gridAfter'), 'val') || 0);
    const phantom = (at, span) => {
      const widths = grid.length >= at + span ? grid.slice(at, at + span) : null;
      return { type: 'table_cell', attrs: { colspan: span, rowspan: 1, colwidth: widths, colwidth0: widths, phantom: true }, content: [{ type: 'paragraph' }] };
    };
    if (gridBefore) row.content.push(phantom(0, gridBefore));
    col += gridBefore;
    // Cells, including cells inside cell-level content controls (w:sdt), which keep their wrapper in `sdt`.
    const tcs = [];
    for (const c of kids(tr)) {
      if (c.namespaceURI !== NS.w) continue;
      if (c.localName === 'tc') tcs.push([c, null]);
      else if (c.localName === 'sdt') {
        const content = kid(c, 'sdtContent');
        const head = kids(c).filter((x) => x !== content).map((x) => ctx.xml(x)).join('');
        const sdt = { open: `<w:sdt>${head}<w:sdtContent>`, close: '</w:sdtContent></w:sdt>', g: ++sdtGroup };
        for (const tc of content ? kids(content, 'tc') : []) tcs.push([tc, sdt]);
      }
    }
    for (const [tc, cellSdt] of tcs) {
      const tcPr = kid(tc, 'tcPr');
      const span = Number(attr(kid(tcPr, 'gridSpan'), 'val') || 1);
      const vMergeEl = kid(tcPr, 'vMerge');
      const vMerge = vMergeEl ? attr(vMergeEl, 'val') || 'continue' : null;
      if (vMerge === 'continue' && open.has(col)) {
        open.get(col).attrs.rowspan++;
        col += span;
        continue;
      }
      const extra = kids(tcPr).filter((c) => c.localName !== 'gridSpan' && c.localName !== 'vMerge').map((c) => [c.localName, ctx.xml(c)]);
      const widths = grid.length >= col + span ? grid.slice(col, col + span) : null;
      const blocks = readBlocks(kids(tc).filter((c) => c !== tcPr), ctx);
      const cell = {
        type: 'table_cell',
        attrs: { colspan: span, rowspan: 1, colwidth: widths, colwidth0: widths, tcPr: extra, sdt: cellSdt },
        content: blocks.length ? blocks : [{ type: 'paragraph' }],
      };
      if (vMerge === 'restart') for (let i = 0; i < span; i++) open.set(col + i, cell);
      else for (let i = 0; i < span; i++) open.delete(col + i);
      row.content.push(cell);
      col += span;
    }
    if (gridAfter && row.content.length) row.content.push(phantom(col, gridAfter));
    if (!row.content.length) row.content.push({ type: 'table_cell', content: [{ type: 'paragraph' }] });
    rows.push(row);
  }
  return {
    type: 'table',
    attrs: { tblPr: tblPr ? ctx.xml(tblPr) : null, tblGrid: tblGrid ? ctx.xml(tblGrid) : null },
    content: rows,
  };
}

/** Block XML (w:p, w:tbl...) as editor JSON nodes, for content we generate (a table of contents...). */
export function blocksFromXml(xml) {
  const { root, ctx } = fragment(xml);
  return readBlocks(kids(root), ctx);
}

/** Footnotes or endnotes part (w:footnotes / w:endnotes) as [{ id, type, xml, content (JSON blocks) }]. */
export function readNotes(xmlDoc, opts = {}) {
  const root = xmlDoc.documentElement;
  const rootNs = nsDecls(root);
  const ctx = { rootNs, resolveImage: () => null, resolveLink: () => null, resolveChart: () => null, resolveShape: () => null, ...opts };
  ctx.xml = (el) => outerXml(el, rootNs);
  return kids(root).filter((el) => el.localName === 'footnote' || el.localName === 'endnote').map((el) => ({
    id: attr(el, 'id'),
    type: attr(el, 'type') || 'normal',
    xml: ctx.xml(el),
    content: readBlocks(kids(el), ctx),
  }));
}
