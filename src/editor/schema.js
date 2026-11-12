// ProseMirror schema mirroring WordprocessingML closely enough that a
// document can go docx -> editor -> docx without loss. Anything we don't model
// is carried in `opaque_*` nodes or `extra` attrs as raw XML and written back
// verbatim.
import { Schema } from 'prosemirror-model';
import { tableNodes } from 'prosemirror-tables';
import { legacyInfo } from '../ooxml/ffdata.js';
import { ommlToMathML } from '../ooxml/omml.js';
import { pPrToCss, scaleCss, sizeCss, highlightColor, hexColor, fontStack, lineFactor } from '../ooxml/props.js';
import { TWIPS_PER_PX } from '../ooxml/units.js';
import { wordDate } from './util.js';

/** A link's address as the page may follow it: web, mail and phone links only (a document's javascript: or data: link
 * would run in the page when clicked in read-only mode). The document keeps its address as it was. */
export const safeHref = (href) => (href && /^(https?:|mailto:|tel:)/i.test(String(href).trim()) ? String(href).trim() : null);

const tables = tableNodes({
  tableGroup: 'block',
  cellContent: 'block+',
  cellAttributes: {
    tcPr: { // [[localName, xml], ...] minus gridSpan/vMerge; a tracked cell insertion or deletion is shown (editor.css)
      default: [],
      // A pasted cell (a table from another program or a web page): its background and vertical alignment, as cell properties.
      getFromDOM(dom) {
        const out = [];
        const fill = cssHex(dom.style?.backgroundColor);
        if (fill) out.push(['shd', `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`]);
        const va = { middle: 'center', bottom: 'bottom', top: 'top' }[dom.style?.verticalAlign];
        if (va) out.push(['vAlign', `<w:vAlign w:val="${va}"/>`]);
        return out;
      },
      setDOMAttr(value, attrs) {
        const rev = (value || []).some((t) => t[0] === 'cellIns') ? 'ins' : (value || []).some((t) => t[0] === 'cellDel') ? 'del' : null;
        if (rev) attrs['data-odx-cell'] = rev;
      },
    },
    colwidth0: { default: null }, // colwidth as read, to detect resizes
    sdt: { default: null, getFromDOM: () => null, setDOMAttr() {} }, // a cell-level content control's wrapper (w:sdt around w:tc), for the writer
    // Stands in for a row's w:gridBefore/w:gridAfter, the grid columns a row leaves empty; drawn blank, never saved.
    phantom: {
      default: false,
      getFromDOM: (dom) => dom.hasAttribute('data-odx-phantom'),
      setDOMAttr(value, attrs) { if (value) { attrs['data-odx-phantom'] = ''; attrs.contenteditable = 'false'; } },
    },
    deco: {
      default: null, // computed borders/shading (editor only, never saved)
      getFromDOM: () => null,
      setDOMAttr(value, attrs) {
        if (value && value.style) attrs.style = (attrs.style ? attrs.style + ';' : '') + value.style;
      },
    },
  },
});
tables.table.attrs = { tblPr: { default: null }, tblGrid: { default: null }, deco: { default: null }, sdt: { default: null }, chunk: { default: null } }; // deco: placement (editor only); chunk: as a paragraph's
tables.table.toDOM = (node) => {
  const m = node.attrs.tblPr && /<w:tblStyle w:val="([^"]+)"/.exec(node.attrs.tblPr);
  return ['table', m ? { 'data-tstyle': m[1] } : {}, ['tbody', 0]];
};
tables.table_row.attrs = { trPr: { default: null }, tblPrEx: { default: null }, sdt: { default: null } };
tables.table_row.toDOM = (node) => {
  // w:trHeight: a minimum row height (also used for "exact", which CSS rows can't enforce). In twips, with a fraction
  // from some generators (doc 25's header row: "312.978515625").
  const m = node.attrs.trPr && /<w:trHeight\b[^>]*w:val="(\d+(?:\.\d+)?)"/.exec(node.attrs.trPr);
  const auto = node.attrs.trPr && /<w:trHeight\b[^>]*w:hRule="auto"/.test(node.attrs.trPr);
  // Word's minimum leaves out the cells' top and bottom margins: the cells' content box gets it (see editor.css).
  const a = m && !auto && +m[1] ? { style: `--odx-trh:${+m[1] / TWIPS_PER_PX}px` } : {};
  // A deleted row (tracked) disappears unless all markup is shown, as in Word.
  if (node.attrs.trPr && /<w:del\b/.test(node.attrs.trPr)) a.class = 'odx-row-del';
  else if (node.attrs.trPr && /<w:ins\b/.test(node.attrs.trPr)) a.class = 'odx-row-ins';
  // "Exact": Word clips what doesn't fit, so the row never grows (see editor.css).
  if (a.style && /<w:trHeight\b[^>]*w:hRule="exact"/.test(node.attrs.trPr)) a['data-exact'] = '';
  return ['tr', a, 0];
};

function revAttrs(d) {
  return { author: d.getAttribute('data-author') };
}

const HEADING_TAGS = { h1: 'Heading1', h2: 'Heading2', h3: 'Heading3', h4: 'Heading4', h5: 'Heading5', h6: 'Heading6' };

/** A CSS colour (rgb(), #rrggbb) as Word's RRGGBB, or null (none, transparent, white). */
function cssHex(v) {
  const m = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?/.exec(v || '');
  const hex = m ? (m[4] != null && Number(m[4]) === 0 ? null : [m[1], m[2], m[3]].map((x) => Number(x).toString(16).padStart(2, '0')).join('')) : (/^#([0-9a-f]{6})$/i.exec(v || '') || [])[1];
  return hex && hex.toUpperCase() !== 'FFFFFF' ? hex.toUpperCase() : null;
}

/** A pasted paragraph's spacing and indents (CSS margins, text-indent, line-height) as w:spacing and w:ind. */
function pastedSpacing(d) {
  const tw = (v) => {
    const m = /^(-?[\d.]+)(pt|px)$/.exec(v || '');
    return m ? Math.round(Number(m[1]) * (m[2] === 'pt' ? 20 : 15)) : null;
  };
  const extra = [], view = {};
  const before = tw(d.style.marginTop), after = tw(d.style.marginBottom);
  const lh = /^[\d.]+$/.test(d.style.lineHeight || '') ? Math.round(Number(d.style.lineHeight) * 240) : null;
  if (before != null || after != null || lh) {
    extra.push(['spacing', `<w:spacing${before != null ? ` w:before="${before}"` : ''}${after != null ? ` w:after="${after}"` : ''}${lh ? ` w:line="${lh}" w:lineRule="auto"` : ''}/>`]);
    if (before != null) view.before = before;
    if (after != null) view.after = after;
    if (lh) Object.assign(view, { line: lh, lineRule: 'auto' });
  }
  const left = tw(d.style.marginLeft), first = tw(d.style.textIndent);
  if (left || first) {
    extra.push(['ind', `<w:ind${left ? ` w:left="${left}"` : ''}${first ? (first > 0 ? ` w:firstLine="${first}"` : ` w:hanging="${-first}"`) : ''}/>`]);
    if (left) view.indLeft = left;
    if (first) view.indFirst = first;
  }
  return extra.length ? { extra, view } : {};
}

/** Pictures as drawn (their address -> their attrs), so a copy of one pasted back in is the same picture. */
export const shownImages = new Map();

export const schema = new Schema({
  nodes: {
    doc: { content: 'block+', attrs: { sectPr: { default: null } } },

    paragraph: {
      group: 'block',
      content: 'inline*',
      attrs: {
        sdt: { default: null }, // block content control (w:sdt) around this block, kept for saving
        chunk: { default: null }, // imported content (w:altChunk) this block was read from, written back while unedited (reader.js)
        styleId: { default: null },
        align: { default: null },
        numId: { default: null },
        ilvl: { default: 0 },
        extra: { default: [] }, // unmodelled w:pPr children, raw
        pattrs: { default: null }, // the w:p element's own attributes ([qname, value]: w14:paraId, w:rsidR...), raw
        view: { default: {} }, // spacing/indent read from extra, for display only
        deco: { default: null }, // computed by the editor (list label, contextual spacing); never saved
      },
      parseDOM: [
        { tag: 'p', getAttrs: (d) => ({ styleId: d.getAttribute('data-style') || null, align: alignFromCss(d.style.textAlign), ...pastedSpacing(d) }) },
        ...Object.entries(HEADING_TAGS).map(([tag, styleId]) => ({ tag, attrs: { styleId } })),
        { tag: 'li', getAttrs: () => ({}) },
      ],
      toDOM(node) {
        const { styleId, align, view, numId } = node.attrs;
        const a = {};
        if (styleId) a['data-style'] = styleId;
        if (numId != null) a['data-num'] = numId;
        const css = pPrToCss({ ...view, align });
        if (css) a.style = css;
        const d = node.attrs.deco;
        if (d) {
          if (d.class) a.class = d.class;
          if (d['data-label']) a['data-label'] = d['data-label'];
          if (d.style) a.style = (a.style ? a.style + ';' : '') + d.style;
          if (d.role) { a.role = d.role; a['aria-level'] = d['aria-level']; } // (a heading, editor.js)
          for (const k of ['data-frame', 'data-fy', 'data-fyalign']) if (d[k] != null) a[k] = d[k]; // (a frame, frames.js)
        }
        if (node.attrs.sdt) a['data-field'] = (/<w:tag w:val="([^"]*)"/.exec(node.attrs.sdt.open) || [])[1] || '';
        return ['p', a, 0];
      },
    },

    opaque_block: {
      group: 'block',
      atom: true,
      selectable: true,
      attrs: { sdt: { default: null }, chunk: { default: null }, xml: { default: '' }, text: { default: '' }, kind: { default: '' } },
      toDOM: (n) => ['div', { class: 'odx-opaque', 'data-kind': n.attrs.kind, contenteditable: 'false', title: `Preserved ${n.attrs.kind} (not editable yet)` }, n.attrs.text || `[${n.attrs.kind}]`],
    },

    ...tables,

    text: { group: 'inline' },

    hard_break: {
      inline: true,
      group: 'inline',
      selectable: false,
      attrs: { type: { default: null }, clear: { default: null } },
      parseDOM: [{ tag: 'br' }],
      // A text-wrapping break that clears floats (w:clear) drops below floating pictures and tables beside it, as a
      // zero-height block with CSS clear; with nothing beside it, it ends the line like a plain break would. (The zero-width
      // space after it gives the next line its height even when nothing else follows.)
      toDOM: (n) => (n.attrs.type === 'page' ? ['span', { class: 'odx-page-break', contenteditable: 'false' }, ['br']]
        : n.attrs.clear && n.attrs.clear !== 'none' ? ['span', { class: 'odx-clear', contenteditable: 'false' }, ['span', { style: `display:block;clear:${n.attrs.clear === 'all' ? 'both' : n.attrs.clear}` }], '​']
          : ['br']),
    },

    tab: {
      inline: true,
      group: 'inline',
      selectable: false,
      toDOM: () => ['span', { class: 'odx-tab' }, '\t'],
      parseDOM: [{ tag: 'span.odx-tab' }],
    },

    image: {
      inline: true,
      group: 'inline',
      atom: true,
      draggable: true,
      marks: 'link ins del sdt',
      attrs: { xml: { default: '' }, src: { default: null }, width: { default: null }, height: { default: null }, alt: { default: '' }, place: { default: '' }, chart: { default: null }, href: { default: null } }, // href: the picture's link (a:hlinkClick); chart: an edited chart's XML (charts.js), written into its part on save
      toDOM: (n) => {
        if (n.attrs.src && n.attrs.xml) shownImages.set(n.attrs.src, n.attrs);
        const link = n.attrs.href && safeHref(n.attrs.href);
        // (A shape drawn with room around its box for its outline: reader.js readDrawing's --odx-pad.)
        const pad = Number(/--odx-pad:([\d.]+)px/.exec(n.attrs.place || '')?.[1]) || 0;
        return ['img', { src: n.attrs.src || '', alt: n.attrs.alt, 'data-href': link || null, title: link ? `${link}\nCtrl+Click to follow link` : null, class: /^display:block/.test(n.attrs.place || '') ? 'odx-wrap-block' : null, style: [n.attrs.width ? `width:${n.attrs.width + 2 * pad}px;height:${n.attrs.height + 2 * pad}px` : '', n.attrs.place].filter(Boolean).join(';') }];
      },
      // Pasted pictures: the document's own (copied here: what it was, by its address) or one given in full (a data: URL,
      // from a web page or another editor), which the paste stores in the package (editor.js pastePictures).
      parseDOM: [{
        tag: 'img[src]',
        getAttrs: (d) => {
          const src = d.getAttribute('src') || '';
          if (shownImages.has(src)) return { ...shownImages.get(src) };
          if (!/^data:image\/(png|jpeg|gif|bmp|svg\+xml)[;,]/i.test(src)) return false;
          const size = (k) => Number(d.getAttribute(k)) || parseFloat(d.style?.[k]) || null;
          return { src, xml: '', width: size('width'), height: size('height'), alt: d.getAttribute('alt') || '' };
        },
      }],
    },

    opaque_inline: {
      inline: true,
      group: 'inline',
      atom: true,
      marks: 'link ins del sdt',
      attrs: { xml: { default: '' }, text: { default: '' }, kind: { default: '' }, width: { default: null }, height: { default: null } },
      toDOM(n) {
        const a = { class: 'odx-opaque-inline', 'data-kind': n.attrs.kind, contenteditable: 'false' };
        // A graphic we can't draw yet keeps its size, so lines and pages break where Word's do.
        if (n.attrs.width != null) a.style = `display:inline-block;width:${n.attrs.width}px;height:${n.attrs.height}px`;
        // Field hints, used to show live page numbers in headers and footers.
        if (n.attrs.kind === 'fldChar') a['data-fld'] = (/fldCharType="(\w+)"/.exec(n.attrs.xml) || [])[1];
        // A legacy check box or drop-down shows its state, which lives in the field's ffData.
        const ff = n.attrs.kind === 'fldChar' && legacyInfo(n.attrs.xml);
        if (ff && ff.kind !== 'text') {
          a.class += ' odx-ff';
          return ['span', a, ff.kind === 'checkbox' ? (ff.checked ? '\u2612' : '\u2610') : ff.text];
        }
        // A symbol (w:sym) in its own font: Wingdings' 0xE0 is an arrow, not the "à" the run's font would draw.
        else if (n.attrs.kind === 'sym') { const f = /w:font="([^"]+)"/.exec(n.attrs.xml); if (f) a.style = `font-family:"${f[1]}"`; }
        else if (n.attrs.kind === 'instrText') a['data-instr'] = (/<w:instrText[^>]*>([^<]*)</.exec(n.attrs.xml) || [])[1] || '';
        else if (n.attrs.kind === 'fldSimple') a['data-instr'] = (/w:instr="([^"]*)"/.exec(n.attrs.xml) || [])[1] || '';
        // An equation, laid out as MathML (ooxml/omml.js); its text stays the fallback where that can't be read.
        else if ((n.attrs.kind === 'oMath' || n.attrs.kind === 'oMathPara') && typeof document !== 'undefined') {
          const math = ommlToMathML(n.attrs.xml);
          if (math) {
            const el = document.createElement('span');
            for (const [k, v] of Object.entries(a)) el.setAttribute(k, v);
            el.classList.add('odx-math');
            el.innerHTML = math;
            return el;
          }
        }
        return ['span', a, n.attrs.text];
      },
    },
  },

  marks: {
    // Order here is the w:rPr child order used when writing.
    cstyle: { attrs: { val: {} }, toDOM: (m) => ['span', { 'data-cstyle': m.attrs.val }, 0] },
    font: { attrs: { val: {}, xml: { default: null } }, parseDOM: [{ style: 'font-family', getAttrs: pastedFont }], toDOM: (m) => ['span', { style: `font-family:${fontStack(m.attrs.val)}${/w:cs="([^"]+)"/.test(m.attrs.xml || '') ? `;--odx-fcs:${fontStack(/w:cs="([^"]+)"/.exec(m.attrs.xml)[1])}` : ''}${/^cambria math$/i.test(m.attrs.val) ? ';--odx-lhv:1.02' : (lineFactor(m.attrs.val) > 1.25 ? `;--odx-lhv:calc(${lineFactor(m.attrs.val)} * var(--odx-runmult, 1));--odx-flh:${lineFactor(m.attrs.val)}` : '')}` }, 0] }, // a tall font (symbols, Segoe UI) makes its line taller, as in Word, // Word doesn't let Cambria Math's tall metrics raise a text line, only its descent (0.222em, below Arial's 0.212): a 1.02 box adds about what Word does (fitted on sample S044, Arial 10pt: +0.26px a line)
    bold: {
      attrs: { on: { default: true } },
      parseDOM: [{ tag: 'strong' }, { tag: 'b' }, { style: 'font-weight', getAttrs: (v) => (/^(bold|[6-9]\d\d)$/.test(v) ? {} : false) }],
      toDOM: (m) => (m.attrs.on ? ['strong', 0] : ['span', { style: 'font-weight:normal' }, 0]),
    },
    italic: {
      attrs: { on: { default: true } },
      parseDOM: [{ tag: 'em' }, { tag: 'i' }, { style: 'font-style=italic' }],
      toDOM: (m) => (m.attrs.on ? ['em', 0] : ['span', { style: 'font-style:normal' }, 0]),
    },
    strike: {
      attrs: { on: { default: true } },
      parseDOM: [{ tag: 's' }, { tag: 'del', priority: 40 }, { style: 'text-decoration=line-through' }],
      toDOM: (m) => (m.attrs.on ? ['s', 0] : ['span', { style: 'text-decoration:none' }, 0]),
    },
    color: { attrs: { val: {} }, parseDOM: [{ style: 'color', getAttrs: pastedColor }], toDOM: (m) => ['span', { style: `color:#${hexColor(m.attrs.val) || '000000'}` }, 0] },
    size: { attrs: { val: {} }, parseDOM: [{ style: 'font-size', getAttrs: pastedSize }], toDOM: (m) => ['span', { style: sizeCss(m.attrs.val / 2) }, 0] },
    highlight: { attrs: { val: {} }, toDOM: (m) => ['span', { style: `background-color:${highlightColor(m.attrs.val)}` }, 0] },
    underline: {
      attrs: { val: { default: 'single' } },
      parseDOM: [{ tag: 'u' }, { style: 'text-decoration=underline' }],
      toDOM: (m) => ['u', { style: m.attrs.val === 'none' ? 'text-decoration:none' : m.attrs.val === 'double' ? 'text-decoration-style:double' : '' }, 0],
    },
    vertAlign: {
      attrs: { val: {} },
      parseDOM: [{ tag: 'sup', attrs: { val: 'superscript' } }, { tag: 'sub', attrs: { val: 'subscript' } }],
      // "baseline" (the normal position, often written out explicitly by other generators) keeps the text as it is.
      toDOM: (m) => (m.attrs.val === 'superscript' ? ['sup', 0] : m.attrs.val === 'subscript' ? ['sub', 0] : ['span', { style: 'vertical-align:baseline' }, 0]),
    },
    rextra: { // unmodelled w:rPr children, raw
      attrs: { items: { default: [] } },
      toDOM: (m) => {
        const change = m.attrs.items.find((it) => it[0] === 'rPrChange');
        const a = { style: rextraCss(m.attrs.items) };
        if (a.style.includes('--odx-grow-')) a.class = 'odx-grow'; // (editor.css: a taller line for emphasis marks and run borders)
        // A complex-script run (right-to-left, or marked w:cs) is drawn with the bCs/iCs/szCs/cs-font members, as in Word.
        if (m.attrs.items.some(([n, x]) => (n === 'rtl' || n === 'cs') && !/w:val="(0|false|off)"/.test(x))) a.class = [a.class, 'odx-cs'].filter(Boolean).join(' ');
        if (change) Object.assign(a, { class: [a.class, 'odx-fmt'].filter(Boolean).join(' '), title: revTitle('Formatted', revXmlAttrs(change[1])) });
        // Text in another language is checked by that language's dictionary, as in Word (w:lang); text Word is told not
        // to check (w:noProof: code, names) isn't checked.
        const lang = m.attrs.items.find((it) => it[0] === 'lang');
        const tag = lang && /w:val="([^"]+)"/.exec(lang[1])?.[1];
        if (tag) a.lang = tag;
        const noProof = m.attrs.items.find((it) => it[0] === 'noProof');
        if (noProof && !/w:val="(0|false|off)"/.test(noProof[1])) a.spellcheck = 'false';
        return ['span', a, 0];
      },
    },
    // An inline content control (form field), written back as <w:sdt> around its runs.
    sdt: {
      attrs: { key: { default: 0 }, pr: { default: '' }, endPr: { default: '' }, tag: { default: null }, alias: { default: null }, kind: { default: 'richText' }, phText: { default: null } },
      inclusive: true,
      excludes: '',
      toDOM: (m) => {
        const a = { class: `odx-field odx-field-${m.attrs.kind}${m.attrs.phText != null ? ' odx-field-ph' : ''}`, 'data-tag': m.attrs.tag || '', title: m.attrs.alias || m.attrs.tag || '' };
        // A check box is one to screen readers, with its state (NVDA read its glyph, "☐", outside forms mode).
        if (m.attrs.kind === 'checkbox') Object.assign(a, { role: 'checkbox', 'aria-checked': String(/<w14:checked\b[^>]*w14:val="(1|true)"/.test(m.attrs.pr)), 'aria-label': m.attrs.alias || m.attrs.tag || 'Check box' });
        return ['span', a, 0];
      },
    },
    link: {
      attrs: { href: { default: null }, rId: { default: null }, anchor: { default: null }, extra: { default: null } },
      inclusive: false,
      parseDOM: [{ tag: 'a[href]', getAttrs: (d) => { const h = d.getAttribute('href') || ''; return safeHref(h) ? { href: h.trim() } : h.startsWith('#') && h.length > 1 ? { anchor: h.slice(1) } : false; } }], // (pasted: web, mail and phone links, and links within the document)
      toDOM: (m) => ['a', { href: safeHref(m.attrs.href) || (m.attrs.anchor ? '#' + m.attrs.anchor : '#'), title: m.attrs.href || '', class: 'odx-link' }, 0],
    },
    ins: {
      attrs: { id: { default: null }, author: { default: null }, date: { default: null }, dateUtc: { default: null }, move: { default: false } },
      inclusive: false,
      parseDOM: [{ tag: 'ins.odx-ins', getAttrs: revAttrs }],
      toDOM: (m) => ['ins', { class: m.attrs.move ? 'odx-ins odx-move' : 'odx-ins', 'data-author': m.attrs.author, title: revTitle(m.attrs.move ? 'Moved here' : 'Inserted', m.attrs) }, 0],
    },
    // Tracked deletion: the text stays, struck through, until the change is accepted.
    del: {
      attrs: { id: { default: null }, author: { default: null }, date: { default: null }, dateUtc: { default: null }, move: { default: false } },
      inclusive: false,
      parseDOM: [{ tag: 'del.odx-del', priority: 60, getAttrs: revAttrs }],
      toDOM: (m) => ['del', { class: m.attrs.move ? 'odx-del odx-move' : 'odx-del', 'data-author': m.attrs.author, title: revTitle(m.attrs.move ? 'Moved from here' : 'Deleted', m.attrs) }, 0],
    },
  },
});

const EMPHASIS = { dot: 'filled dot', comma: 'filled sesame', circle: 'open circle', underDot: 'filled dot' };

/** Display-only CSS for run properties we carry as raw XML (character spacing, caps, shading...). */
function rextraCss(items) {
  const css = [];
  const effects = new Set();
  const val = (xml) => (/w:val="([^"]*)"/.exec(xml) || [])[1];
  const on = (xml) => !/w:val="(0|false|off)"/.test(xml);
  for (const [name, xml] of items) {
    if (name === 'rFonts' && /w:cs="([^"]+)"/.test(xml)) css.push(`--odx-fcs:${fontStack(/w:cs="([^"]+)"/.exec(xml)[1])}`);
    if (name === 'spacing') css.push(`letter-spacing:${Number(val(xml)) / 20}pt`);
    else if (name === 'rFonts' && /w:(?:ascii|hAnsi)Theme="(major|minor)/.test(xml)) {
      // A theme font on a run: the theme's body or heading font (variables set with the document's styles).
      const t = /w:(?:ascii|hAnsi)Theme="(major|minor)/.exec(xml)[1];
      css.push(`font-family:var(--odx-font-${t})`, `--odx-lh:var(--odx-lh-${t})`);
    }
    // Complex-script formatting (editor.css: a run marked w:rtl or w:cs takes it instead of the Latin members).
    else if (name === 'bCs') css.push(`--odx-bcs:${on(xml) ? 'bold' : 'normal'}`);
    else if (name === 'iCs') css.push(`--odx-ics:${on(xml) ? 'italic' : 'normal'}`);
    else if (name === 'szCs' && Number(val(xml))) css.push(`--odx-szcs:${Number(val(xml)) / 2}pt`);
    else if (name === 'caps' && on(xml)) css.push('text-transform:uppercase');
    else if (name === 'smallCaps' && on(xml)) css.push('font-variant:small-caps');
    else if (name === 'vanish' && on(xml)) css.push('display:none');
    else if (name === 'shd') {
      const fill = (/w:fill="([0-9A-Fa-f]{6})"/.exec(xml) || [])[1];
      if (fill) css.push(`background-color:#${fill}`);
    } else if (name === 'dstrike' && on(xml)) css.push('text-decoration:line-through double');
    else if (name === 'w' && Number(val(xml)) && Number(val(xml)) !== 100) css.push(scaleCss(val(xml)));
    else if (name === 'kern') css.push(`font-kerning:${Number(val(xml)) ? 'normal' : 'none'}`);
    // Raised or lowered text (half-points) moves its line's top and bottom with it, as in Word: a baseline shift, not an offset.
    else if (name === 'position') css.push(`vertical-align:${Number(val(xml)) / 2}pt`);
    // A colour that also names a theme colour (kept raw so saving keeps the theme link): its value still applies.
    else if (name === 'color' && /^[0-9A-Fa-f]{6}$/.test(val(xml) || '')) css.push(`color:#${val(xml)}`);
    // Text effects as Word 2013+ draws them (compared with its PDFs): outline is hollow letters, shadow a grey copy down
    // and right; emboss and imprint are outlined letters with that copy down-right or up-left. None moves the text.
    else if (['outline', 'shadow', 'emboss', 'imprint'].includes(name) && on(xml)) effects.add(name);
    // Emphasis marks: over the letters (dot, comma, circle) or under them (underDot).
    // Word makes the line that much taller (0.31em, its PDFs at 11pt and 18pt): above it, or below for underDot.
    else if (name === 'em' && EMPHASIS[val(xml)]) css.push(`text-emphasis:${EMPHASIS[val(xml)]}`, `text-emphasis-position:${val(xml) === 'underDot' ? 'under' : 'over'} right`, `--odx-grow-${val(xml) === 'underDot' ? 'down' : 'up'}:0.31em`);
    // A border round the run, with its space beside the text (w:space, points) as in Word.
    else if (name === 'bdr') {
      const b = /w:val="([^"]+)"/.exec(xml)?.[1];
      if (!b || b === 'none' || b === 'nil') continue;
      const sz = Math.max(0.25, Number(/w:sz="(\d+)"/.exec(xml)?.[1] || 4) / 8), space = Number(/w:space="(\d+)"/.exec(xml)?.[1] || 0);
      const color = /w:color="([0-9A-Fa-f]{6})"/.exec(xml)?.[1];
      // (Its line is taller by what the border and its space reach past the line's own room, above and below, as in Word:
      // a 0.5pt box adds nothing (sample S125), a 1pt one 1px, 1.5pt with 4pt of space 8px (test/fixtures/effects.docx).)
      const grow = Math.max(0, sz + space - 0.5);
      css.push(`border:${sz}pt ${b === 'double' ? 'double' : b === 'dotted' ? 'dotted' : /dash/i.test(b) ? 'dashed' : 'solid'} ${color ? '#' + color : 'currentColor'}`, `padding:0 ${space}pt`);
      if (grow) css.push(`--odx-grow-up:${grow}pt`, `--odx-grow-down:${grow}pt`);
    }
  }
  if (effects.size) {
    // (Measured on Word's PDF of test/fixtures/effects.docx: the shadow 0.06em down and right in a mid grey, the
    // outline about 0.03em wide.)
    const shade = '0.06em 0.06em 0 #a6a6a6';
    if (effects.has('outline') && !effects.has('emboss') && !effects.has('imprint')) css.push('-webkit-text-stroke:0.03em currentColor', '-webkit-text-fill-color:#fff');
    if (effects.has('emboss') || effects.has('imprint')) css.push('-webkit-text-stroke:0.03em currentColor');
    const shadow = effects.has('imprint') ? '-0.06em -0.06em 0 #a6a6a6' : effects.has('shadow') || effects.has('emboss') ? shade : null;
    if (shadow) css.push(`text-shadow:${shadow}`);
  }
  return css.join(';');
}

/** w:author and w:date of a revision element's XML. */
export function revXmlAttrs(xml) {
  const a = (n) => (new RegExp(`^<[^>]*\\sw:${n}="([^"]*)"`).exec(xml) || [])[1] || null;
  return { id: a('id'), author: a('author'), date: a('date') };
}
// Pasted text's colour, size and font (from a web page, another editor, or this one): as marks, where they're definite.
// Black is no colour (some editors set it on every run); a size in pt or px only; a font by its first family name,
// not a generic one.
function pastedColor(v) {
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+)\s*)?\)$/.exec(v || '');
  const hex = rgb ? (rgb[4] != null && Number(rgb[4]) === 0 ? null : rgb.slice(1, 4).map((x) => Math.min(255, Number(x)).toString(16).padStart(2, '0')).join('')) : /^#([0-9a-f]{6})$/i.exec(v || '')?.[1] || (/^#([0-9a-f]{3})$/i.exec(v || '')?.[1]?.replace(/./g, '$&$&'));
  return hex && !/^0{6}$/.test(hex) ? { val: hex.toUpperCase() } : false;
}
function pastedSize(v) {
  const m = /^([\d.]+)(pt|px)$/.exec(v || '');
  const pt = m && (m[2] === 'pt' ? Number(m[1]) : Number(m[1]) * 0.75);
  return pt >= 1 && pt <= 400 ? { val: Math.round(pt * 2) } : false;
}
function pastedFont(v) {
  const first = (v || '').split(',')[0].trim().replace(/^["']|["']$/g, '');
  return first && !/^(serif|sans-serif|monospace|cursive|fantasy|system-ui|inherit|initial|unset|-apple-system|blinkmacsystemfont|var\()/i.test(first) ? { val: first } : false;
}

function revTitle(what, { author, date, dateUtc }) {
  const d = wordDate(date, dateUtc);
  const when = d ? ' on ' + d.toLocaleString() : '';
  return `${what} by ${author || 'unknown'}${when}`;
}

function alignFromCss(v) {
  return { center: 'center', right: 'right', end: 'right', justify: 'both' }[v] || null;
}

