// Imported content (w:altChunk): a part of HTML or plain text that Word turns into ordinary paragraphs when it opens
// the document, as mail-merge systems write rich-text fields. Here the part is turned into WordprocessingML (w:p and
// w:tbl elements) the reader then reads like any other, following what Word makes of each (checked against Word's
// PDFs of test/fixtures/altchunk.docx):
//   - text in a <p> (or <pre>, <blockquote>) is in Word's HTML paragraph formatting: Times New Roman 12pt, with automatic
//     spacing of 14pt before and after (which two such paragraphs share);
//   - text outside one (straight in the body, a <div>, a list item, a table cell) is in the document's own Normal
//     style; a body with nothing but white space in it is one empty paragraph;
//   - <h1> to <h6> take the document's Heading 1 to 6 styles; <ul> and <ol> are lists (a bullet, or 1. 2. 3., at 0.25in
//     a level, the text 0.25in further), each <ol> numbered from 1;
//   - bold, italic, underline, strikethrough, superscript, subscript, links, <font> and the inline styles font-family,
//     font-size, color, background-color, font-weight, font-style and text-decoration; text-align on a block;
//   - <br> is a line break, a <table> a table (with borders when it has border="1" or more);
//   - a plain-text part is one paragraph a line, in Courier New 10.5pt with no spacing.
// HTML paragraphs, list items and table cells are single-spaced, whatever the document's own line spacing.
// HTML that isn't well-formed is read the way a browser reads it.
import { esc } from './xml.js';

const BLOCK = new Set(['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'table', 'pre', 'blockquote', 'hr', 'section', 'article', 'header', 'footer', 'main', 'nav', 'aside', 'center', 'address', 'dl', 'dt', 'dd', 'form', 'fieldset', 'figure', 'body', 'html']);
const SKIP = new Set(['head', 'script', 'style', 'title', 'meta', 'link', 'template', 'noscript', 'img', 'object', 'iframe', 'svg', 'input', 'button', 'select', 'textarea']);
const WEB_P = { font: 'Times New Roman', sz: 24 }; // Word's HTML paragraph (its "Normal (Web)" formatting)
const WEB_SPACING = '<w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1" w:line="240" w:lineRule="auto"/>';
const JC = { center: 'center', right: 'right', justify: 'both', left: 'left' };
const LI_SPACING = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
const LI_FIRST = '<w:spacing w:before="100" w:beforeAutospacing="1" w:after="0" w:line="240" w:lineRule="auto"/>';
const LI_LAST = '<w:spacing w:before="0" w:after="100" w:afterAutospacing="1" w:line="240" w:lineRule="auto"/>';
const LI_ONLY = WEB_SPACING;
const CELL_SPACING = '<w:spacing w:before="0" w:after="100" w:afterAutospacing="1" w:line="240" w:lineRule="auto"/>';

/** The chunk's type from its content type (or file name): 'html', 'text', or null for one this can't convert (RTF, a
 * whole .docx, MHT). */
export function chunkKind(contentType, name = '') {
  const t = (contentType || '').toLowerCase();
  if (/x?html/.test(t) || /\.(x?html?|xht)$/i.test(name)) return 'html';
  if (t === 'text/plain' || /\.txt$/i.test(name)) return 'text';
  return null;
}

/**
 * The chunk as WordprocessingML blocks, or null when it can't be converted.
 * @param {string} source the part's text
 * @param {'html'|'text'} kind
 * @param {{ parseHtml: (s: string) => Document, list: (kind: 'bullet'|'decimal') => string, link: (href: string) => string }} env
 *   parseHtml reads HTML as a browser does; list gives a numId for a new list; link a relationship id for a link's address.
 */
export function chunkToWml(source, kind, env) {
  if (kind === 'text') return textBlocks(source);
  if (kind !== 'html') return null;
  let doc;
  // (A fragment, with no <html> or <body>, read as a body: an XML-based parser wants one root.)
  const src = /<(html|body)\b/i.test(source) ? source : `<html><body>${source}</body></html>`;
  try { doc = env.parseHtml(src); } catch { return null; }
  const body = find(doc, 'body') || doc; // (a fragment with no body: all of it)
  const out = new Writer(env);
  out.blockKids(body, { para: null, run: {} });
  out.flush();
  return out.xml || '<w:p/>';
}

function textBlocks(source) {
  const lines = String(source).replace(/^\uFEFF/, '').split(/\r\n|\r|\n/);
  if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop();
  const rPr = '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>';
  return lines.map((l) => `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>${rPr}</w:pPr>${l ? `<w:r>${rPr}<w:t xml:space="preserve">${esc(l)}</w:t></w:r>` : ''}</w:p>`).join('');
}

const name = (n) => (n.localName || n.nodeName || '').toLowerCase();
const isEl = (n) => n.nodeType === 1;
function find(n, tag) {
  if (!n) return null;
  if (isEl(n) && name(n) === tag) return n;
  for (let c = n.firstChild; c; c = c.nextSibling) { const f = find(c, tag); if (f) return f; }
  return null;
}
const styleOf = (el) => {
  const out = {};
  for (const part of (el.getAttribute?.('style') || '').split(';')) {
    const i = part.indexOf(':');
    if (i > 0) out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim();
  }
  return out;
};

/** A CSS or HTML colour as RRGGBB, or null. */
function hex(v) {
  if (!v) return null;
  v = v.trim().toLowerCase();
  const named = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080', silver: 'C0C0C0', maroon: '800000', navy: '000080', purple: '800080', teal: '008080', olive: '808000', orange: 'FFA500', lime: '00FF00', aqua: '00FFFF', fuchsia: 'FF00FF' };
  if (named[v]) return named[v];
  let m = /^#([0-9a-f]{6})$/.exec(v);
  if (m) return m[1].toUpperCase();
  m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(v);
  if (m) return (m[1] + m[1] + m[2] + m[2] + m[3] + m[3]).toUpperCase();
  m = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(v);
  if (m) return [m[1], m[2], m[3]].map((x) => Math.min(255, Number(x)).toString(16).padStart(2, '0')).join('').toUpperCase();
  return null;
}

/** A CSS font size as half-points, or null. */
function halfPoints(v, base) {
  const keywords = { 'xx-small': 15, 'x-small': 20, small: 26, medium: 32, large: 36, 'x-large': 48, 'xx-large': 72 };
  if (!v) return null;
  v = v.trim().toLowerCase();
  if (keywords[v]) return keywords[v];
  const m = /^([\d.]+)(pt|px|em|%)?$/.exec(v);
  if (!m) return null;
  const n = Number(m[1]);
  const pt = m[2] === 'px' ? n * 0.75 : m[2] === 'em' ? n * (base || 24) / 2 : m[2] === '%' ? n / 100 * (base || 24) / 2 : n;
  return pt > 0 ? Math.round(pt * 2) : null;
}

/** <font size="1".."7"> in half-points, as browsers size them. */
const FONT_SIZES = [null, 15, 20, 24, 27, 36, 48, 72];

/** Run formatting from an element: its tag, then its style. */
function runFrom(el, run) {
  const r = { ...run };
  const tag = name(el);
  if (tag === 'b' || tag === 'strong') r.b = true;
  if (tag === 'i' || tag === 'em' || tag === 'cite' || tag === 'var' || tag === 'dfn') r.i = true;
  if (tag === 'u' || tag === 'ins') r.u = true;
  if (tag === 's' || tag === 'strike' || tag === 'del') r.strike = true;
  if (tag === 'sup') r.va = 'superscript';
  if (tag === 'sub') r.va = 'subscript';
  if (tag === 'code' || tag === 'tt' || tag === 'kbd' || tag === 'samp') { r.font = 'Courier New'; r.sz = 20; }
  if (tag === 'th') r.b = true;
  if (tag === 'font') {
    const face = el.getAttribute('face');
    if (face) r.font = face.split(',')[0].trim().replace(/^["']|["']$/g, '');
    const c = hex(el.getAttribute('color'));
    if (c) r.color = c;
    const size = Number(el.getAttribute('size'));
    if (size >= 1 && size <= 7) r.sz = FONT_SIZES[size];
  }
  const st = styleOf(el);
  if (st['font-family']) r.font = st['font-family'].split(',')[0].trim().replace(/^["']|["']$/g, '');
  const sz = halfPoints(st['font-size'], r.sz);
  if (sz) r.sz = sz;
  const c = hex(st.color);
  if (c) r.color = c;
  const bg = hex(st['background-color'] || st.background);
  if (bg) r.shd = bg;
  if (st['font-weight']) r.b = /^(bold|bolder|[6-9]00)$/.test(st['font-weight']);
  if (st['font-style']) r.i = st['font-style'] === 'italic' || st['font-style'] === 'oblique';
  const td = st['text-decoration'] || st['text-decoration-line'];
  if (td) { if (/underline/.test(td)) r.u = true; if (/line-through/.test(td)) r.strike = true; if (/none/.test(td)) { r.u = false; r.strike = false; } }
  if (st['vertical-align'] === 'super') r.va = 'superscript';
  if (st['vertical-align'] === 'sub') r.va = 'subscript';
  return r;
}

function rPrXml(r) {
  const p = [];
  if (r.style) p.push(`<w:rStyle w:val="${esc(r.style)}"/>`);
  if (r.font) p.push(`<w:rFonts w:ascii="${esc(r.font)}" w:hAnsi="${esc(r.font)}" w:cs="${esc(r.font)}"/>`);
  if (r.b) p.push('<w:b/><w:bCs/>');
  if (r.i) p.push('<w:i/><w:iCs/>');
  if (r.strike) p.push('<w:strike/>');
  if (r.color) p.push(`<w:color w:val="${r.color}"/>`);
  if (r.sz) p.push(`<w:sz w:val="${r.sz}"/><w:szCs w:val="${r.sz}"/>`);
  if (r.u) p.push('<w:u w:val="single"/>');
  if (r.shd) p.push(`<w:shd w:val="clear" w:color="auto" w:fill="${r.shd}"/>`);
  if (r.va) p.push(`<w:vertAlign w:val="${r.va}"/>`);
  return p.length ? `<w:rPr>${p.join('')}</w:rPr>` : '';
}

class Writer {
  constructor(env) {
    this.env = env;
    this.xml = '';
    this.para = null; // the paragraph being written: { pPr, runs: [], rPr (its mark's), text so far }
  }

  /** Starts a paragraph with these properties (ends the one open). */
  open(pPr, markRun = {}) {
    this.flush();
    this.para = { pPr, runs: [], mark: markRun, blank: true, space: false };
  }

  flush() {
    const p = this.para;
    if (!p) return;
    this.para = null;
    // (Trailing white space at the end of a line is dropped, as a browser drops it.)
    while (p.runs.length && p.runs[p.runs.length - 1].kind === 'space') p.runs.pop();
    if (p.blank && p.implicit) return; // white space between blocks
    const runs = p.runs.map((r) => {
      const body = r.kind === 'br' ? '<w:br/>' : `<w:t xml:space="preserve">${esc(r.text)}</w:t>`;
      const run = `<w:r>${rPrXml(r.run)}${body}</w:r>`;
      return r.link ? { link: r.link, run } : { run };
    });
    let inner = '';
    for (let i = 0; i < runs.length; i++) {
      if (runs[i].link) {
        let j = i;
        let s = '';
        while (j < runs.length && runs[j].link === runs[i].link) s += runs[j++].run;
        inner += `<w:hyperlink r:id="${esc(this.env.link(runs[i].link))}" w:history="1">${s}</w:hyperlink>`;
        i = j - 1;
      } else inner += runs[i].run;
    }
    const mark = rPrXml(p.mark);
    this.xml += `<w:p>${p.pPr || mark ? `<w:pPr>${p.pPr || ''}${mark}</w:pPr>` : ''}${inner}</w:p>`;
  }

  text(s, ctx) {
    s = s.replace(/[\t\n\r\f ]+/g, ' ');
    if (!s) return;
    if (!this.para) {
      this.open(ctx.para?.pPr ?? '', ctx.para?.mark ?? {});
      this.para.implicit = !ctx.para;
    }
    const p = this.para;
    if (p.blank || p.space) s = s.replace(/^ /, ''); // (white space at a line's start, or after a space, collapses)
    if (!s) return;
    if (s.trim()) { p.blank = false; p.implicit = false; }
    p.space = s.endsWith(' ');
    // A space on its own is kept apart, so trailing white space can be dropped.
    const word = s.trimEnd();
    if (word) p.runs.push({ kind: 'text', text: word, run: ctx.run, link: ctx.link });
    if (s.length > word.length) p.runs.push({ kind: 'space', text: ' ', run: ctx.run, link: ctx.link });
  }

  br(ctx) {
    if (!this.para) this.open(ctx.para?.pPr ?? '', ctx.para?.mark ?? {});
    const p = this.para;
    while (p.runs.length && p.runs[p.runs.length - 1].kind === 'space') p.runs.pop();
    p.runs.push({ kind: 'br', run: ctx.run });
    p.blank = false; p.implicit = false; p.space = true;
  }

  /** An element's children, in a block context. */
  blockKids(el, ctx) {
    for (let c = el.firstChild; c; c = c.nextSibling) this.node(c, ctx);
  }

  node(n, ctx) {
    if (n.nodeType === 3 || n.nodeType === 4) return this.text(n.nodeValue || '', ctx);
    if (!isEl(n)) return;
    const tag = name(n);
    if (SKIP.has(tag)) return;
    if (tag === 'br') return this.br(ctx);
    if (!BLOCK.has(tag)) {
      // Inline: its formatting for its children; a link its address.
      const href = tag === 'a' && n.getAttribute('href');
      const sub = { ...ctx, run: runFrom(n, ctx.run) };
      if (href && !href.startsWith('#') && !/^javascript:/i.test(href)) { sub.link = href.trim(); sub.run = { ...sub.run, color: sub.run.color || '0563C1', u: true }; }
      return this.blockKids(n, sub);
    }
    const st = styleOf(n);
    const align = JC[(st['text-align'] || n.getAttribute('align') || '').toLowerCase()];
    const jc = align ? `<w:jc w:val="${align}"/>` : '';
    const run = runFrom(n, ctx.run);
    if (tag === 'table') return this.table(n, ctx);
    if (tag === 'ul' || tag === 'ol') return this.list(n, ctx, tag);
    if (tag === 'hr') {
      this.flush();
      this.xml += '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>';
      return;
    }
    let para;
    if (/^h[1-6]$/.test(tag)) para = { pPr: `<w:pStyle w:val="Heading${tag[1]}"/>${jc}`, mark: {} };
    else if (tag === 'p' || tag === 'pre' || tag === 'blockquote' || tag === 'address') {
      const ind = tag === 'blockquote' ? '<w:ind w:left="720" w:right="720"/>' : '';
      Object.assign(run, { font: run.font || WEB_P.font, sz: run.sz || WEB_P.sz });
      if (tag === 'pre') Object.assign(run, { font: 'Courier New', sz: 20 });
      para = { pPr: `${WEB_SPACING}${ind}${jc}`, mark: { font: run.font, sz: run.sz } };
    } else {
      // A div and the like: the paragraph formatting around it, its own alignment.
      const around = ctx.para?.pPr ?? '';
      para = { pPr: jc ? around.replace(/<w:jc [^>]*\/>/, '') + jc : around, mark: ctx.para?.mark ?? {} };
    }
    if (tag === 'pre') return this.pre(n, { ...ctx, para, run });
    this.open(para.pPr, para.mark);
    if (!/^(p|h[1-6]|blockquote|address)$/.test(tag)) this.para.implicit = true; // (an empty div is nothing)
    this.blockKids(n, { ...ctx, para, run });
    this.flush();
  }

  pre(n, ctx) {
    this.open(ctx.para.pPr, {});
    const walk = (el, run) => {
      for (let c = el.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 3 || c.nodeType === 4) {
          (c.nodeValue || '').replace(/^\n/, '').split(/\r\n|\r|\n/).forEach((line, i) => {
            if (i) this.br({ run });
            if (line) { this.para.runs.push({ kind: 'text', text: line.replace(/\t/g, '    '), run }); this.para.blank = false; }
          });
        } else if (isEl(c) && name(c) === 'br') this.br({ run });
        else if (isEl(c)) walk(c, runFrom(c, run));
      }
    };
    walk(n, ctx.run);
    this.flush();
  }

  list(el, ctx, tag) {
    this.flush();
    const level = ctx.listLevel ?? -1;
    const numId = tag === 'ul' ? (this.bullets ||= this.env.list('bullet')) : this.env.list('decimal');
    const ilvl = Math.min(8, level + 1);
    const items = [];
    for (let c = el.firstChild; c; c = c.nextSibling) if (isEl(c) && name(c) === 'li') items.push(c);
    const start = this.xml.length;
    items.forEach((li) => {
      // (No numbering to be had, read without a document's session: the items as plain paragraphs.)
      const pPr = `${numId ? `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>` : ''}${LI_SPACING}`;
      const para = { pPr, mark: {} };
      // An item's text, up to a list inside it; that list's items follow at the next level.
      this.open(pPr, {});
      for (let c = li.firstChild; c; c = c.nextSibling) {
        const t = isEl(c) && name(c);
        if (t === 'ul' || t === 'ol') this.list(c, { ...ctx, listLevel: ilvl, para }, t);
        else if (t === 'p' || t === 'div') { if (!this.para) this.open(pPr, {}); this.blockKids(c, { ...ctx, para, run: runFrom(c, ctx.run) }); }
        else this.node(c, { ...ctx, para, run: ctx.run });
      }
      this.flush();
    });
    // A list's items follow one another with no space between; the list as a whole has the automatic spacing above
    // and below it (as Word lays out an HTML list), set on its first and last paragraphs.
    if (level >= 0) return;
    let s = this.xml.slice(start);
    const first = s.indexOf(LI_SPACING);
    if (first < 0) return;
    s = s.slice(0, first) + LI_FIRST + s.slice(first + LI_SPACING.length);
    const last = Math.max(s.lastIndexOf(LI_SPACING), s.lastIndexOf(LI_FIRST));
    s = s.slice(0, last) + (s.startsWith(LI_FIRST, last) ? LI_ONLY : LI_LAST) + s.slice(last + (s.startsWith(LI_FIRST, last) ? LI_FIRST : LI_SPACING).length);
    this.xml = this.xml.slice(0, start) + s;
  }

  table(el, ctx) {
    this.flush();
    const rows = [];
    const collect = (n) => {
      for (let c = n.firstChild; c; c = c.nextSibling) {
        if (!isEl(c)) continue;
        const t = name(c);
        if (t === 'tr') rows.push(c);
        else if (t === 'thead' || t === 'tbody' || t === 'tfoot') collect(c);
      }
    };
    collect(el);
    const cellsOf = (tr) => { const out = []; for (let c = tr.firstChild; c; c = c.nextSibling) if (isEl(c) && (name(c) === 'td' || name(c) === 'th')) out.push(c); return out; };
    const cols = Math.max(1, ...rows.map((r) => cellsOf(r).reduce((n, c) => n + (Number(c.getAttribute('colspan')) || 1), 0)));
    if (!rows.length) return;
    const border = Number(el.getAttribute('border')) >= 1 || /border/.test(el.getAttribute('style') || '');
    const b = border ? `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('')}</w:tblBorders>` : '';
    const width = Math.round(9360 / cols);
    let xml = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${b}<w:tblLayout w:type="autofit"/><w:tblLook w:val="0000"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="' + width + '"/>'.repeat(cols)}</w:tblGrid>`;
    const outer = this.xml;
    for (const tr of rows) {
      xml += '<w:tr>';
      for (const td of cellsOf(tr)) {
        const span = Math.max(1, Number(td.getAttribute('colspan')) || 1);
        this.xml = '';
        const st = styleOf(td);
        const align = JC[(st['text-align'] || td.getAttribute('align') || (name(td) === 'th' ? 'center' : '')).toLowerCase()];
        // (A cell's text: single-spaced, the automatic spacing below it, as Word lays out an HTML table.)
        const cellPara = { pPr: `${CELL_SPACING}${align ? `<w:jc w:val="${align}"/>` : ''}`, mark: {} };
        this.open(cellPara.pPr, {});
        this.para.implicit = true;
        this.blockKids(td, { ...ctx, para: cellPara, run: runFrom(td, ctx.run) });
        this.flush();
        const inner = this.xml && !/<\/w:tbl>$/.test(this.xml) ? this.xml : this.xml + `<w:p>${cellPara.pPr ? `<w:pPr>${cellPara.pPr}</w:pPr>` : ''}</w:p>`;
        const fill = hex(st['background-color'] || td.getAttribute('bgcolor'));
        xml += `<w:tc><w:tcPr><w:tcW w:w="${width * span}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}${fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : ''}</w:tcPr>${inner}</w:tc>`;
      }
      xml += '</w:tr>';
    }
    this.xml = outer + xml + '</w:tbl>';
  }
}
