// Readers for run (w:rPr) and paragraph (w:pPr) properties into plain objects,
// shared by direct formatting and style definitions. Units stay OOXML-native
// (half-points for size, twips for spacing/indent) until CSS is generated.
import { kid, kids, attr, onOff } from './xml.js';

export function readRPr(rPr) {
  const p = {};
  if (!rPr) return p;
  const flag = (name, key = name) => {
    const v = onOff(kid(rPr, name));
    if (v != null) p[key] = v;
  };
  flag('b', 'bold');
  flag('i', 'italic');
  flag('strike');
  flag('caps');
  flag('smallCaps');
  flag('vanish', 'hidden');
  const u = kid(rPr, 'u');
  if (u) p.underline = attr(u, 'val') || 'single';
  const sz = kid(rPr, 'sz');
  if (sz) p.size = Number(attr(sz, 'val'));
  const color = kid(rPr, 'color');
  if (color && hexColor(attr(color, 'val'))) p.color = attr(color, 'val');
  const hl = kid(rPr, 'highlight');
  if (hl && attr(hl, 'val') !== 'none') p.highlight = attr(hl, 'val');
  const fonts = kid(rPr, 'rFonts');
  if (fonts) {
    if (attr(fonts, 'ascii')) p.font = attr(fonts, 'ascii');
    else if (attr(fonts, 'hAnsi')) p.font = attr(fonts, 'hAnsi');
    else if (attr(fonts, 'asciiTheme')) p.themeFont = attr(fonts, 'asciiTheme');
  }
  // The complex-script members (Hebrew, Arabic...: runs marked w:rtl or w:cs), which Word uses for such text instead.
  flag('bCs', 'boldCs');
  flag('iCs', 'italicCs');
  const szCs = kid(rPr, 'szCs');
  if (szCs) p.sizeCs = Number(attr(szCs, 'val'));
  if (fonts && attr(fonts, 'cs')) p.fontCs = attr(fonts, 'cs');
  else if (fonts && attr(fonts, 'cstheme')) p.themeFontCs = attr(fonts, 'cstheme');
  const va = kid(rPr, 'vertAlign');
  if (va && attr(va, 'val') !== 'baseline') p.vertAlign = attr(va, 'val');
  const rStyle = kid(rPr, 'rStyle');
  if (rStyle) p.rStyle = attr(rStyle, 'val');
  const kern = kid(rPr, 'kern');
  if (kern) p.kern = Number(attr(kern, 'val') || 0);
  const scale = kid(rPr, 'w');
  if (scale) p.scale = Number(attr(scale, 'val') || 100);
  return p;
}

/** A w:pBdr / w:tblBorders / w:tcBorders element -> { top: {val, sz, color, space}, ... } */
export function readBorders(el) {
  const out = {};
  if (!el) return out;
  for (const c of kids(el)) {
    const side = { start: 'left', end: 'right' }[c.localName] || c.localName;
    out[side] = { val: attr(c, 'val'), sz: Number(attr(c, 'sz')) || 4, color: hexColor(attr(c, 'color')), space: Number(attr(c, 'space')) || 0 };
  }
  return out;
}

/** CSS for one border side; Word sizes are eighths of a point. */
export function borderCss(b) {
  if (!b || !b.val || b.val === 'nil' || b.val === 'none') return 'none';
  const style = { double: 'double', dotted: 'dotted', dashed: 'dashed', dashSmallGap: 'dashed', dotDash: 'dashed', triple: 'double', thinThickSmallGap: 'double', thickThinSmallGap: 'double' }[b.val] || 'solid';
  // A double border's size is each of its lines, with a gap as wide between: 3x in all (Word's PDF, doc 01: 2.88px
  // lines, 2.88px apart, for sz 18). (At least 2.25pt, 3px, for the browser to draw two lines.)
  // A thick-thin pair's size is its thick line; the thin line and the gap add 0.75pt (sample S078's 2.25pt
  // header rule: Word's body starts 0.8pt lower than with the size alone).
  const w = b.val === 'double' ? Math.max(2.25, (3 * b.sz) / 8) : /^th(ick|in)Th(in|ick)SmallGap$/.test(b.val) ? Math.max(2.25, b.sz / 8 + 0.75) : Math.max(0.5, b.sz / 8);
  return `${w}pt ${style} ${b.color && b.color !== 'auto' ? '#' + b.color : '#000'}`;
}

export function shdFill(shd) {
  const fill = hexColor(shd && attr(shd, 'fill'));
  // A percentage pattern (pct10: a tenth of the pattern colour over the fill, both "auto" meaning black over white) or
  // a solid one is drawn as the colour Word's dots blend to: pct10 of auto over auto is a light grey, not white.
  const pat = shd && /^(?:pct(\d+)|(solid))$/.exec(attr(shd, 'val') || '');
  if (pat) {
    const share = pat[2] ? 1 : Number(pat[1]) / 100;
    const color = hexColor(attr(shd, 'color')) || '000000';
    const base = fill || 'FFFFFF';
    const mix = [0, 2, 4].map((i) => Math.round(parseInt(base.slice(i, i + 2), 16) * (1 - share) + parseInt(color.slice(i, i + 2), 16) * share));
    return '#' + mix.map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  return fill ? '#' + fill : null;
}

// Word's paragraph alignments the layout knows (w:jc); anything else is left out.
const ALIGN = new Set(['left', 'right', 'center', 'both', 'distribute', 'start', 'end', 'justify', 'thaiDistribute', 'lowKashida', 'mediumKashida', 'highKashida']);

export function readPPr(pPr) {
  const p = {};
  if (!pPr) return p;
  const st = kid(pPr, 'pStyle');
  if (st) p.styleId = attr(st, 'val');
  const jc = kid(pPr, 'jc');
  if (jc && ALIGN.has(attr(jc, 'val'))) p.align = attr(jc, 'val');
  // Right-to-left (w:bidi): its left and right (alignment, indents) are its start and end, as in Word.
  const bidi = onOff(kid(pPr, 'bidi'));
  if (bidi != null) p.bidi = bidi;
  const sp = kid(pPr, 'spacing');
  if (sp) {
    for (const k of ['before', 'after', 'line']) if (attr(sp, k) != null) p[k] = Number(attr(sp, k));
    if (attr(sp, 'lineRule')) p.lineRule = attr(sp, 'lineRule');
    // HTML auto spacing is 14pt each side; Word drops it above the first paragraph of the document or a cell.
    if (/^(1|true|on)$/.test(attr(sp, 'beforeAutospacing') || '')) { p.before = 280; p.beforeAuto = true; }
    if (/^(1|true|on)$/.test(attr(sp, 'afterAutospacing') || '')) { p.after = 280; p.afterAuto = true; }
  }
  const ind = kid(pPr, 'ind');
  if (ind) {
    const left = attr(ind, 'left') ?? attr(ind, 'start');
    const right = attr(ind, 'right') ?? attr(ind, 'end');
    if (left != null) p.indLeft = Number(left);
    if (right != null) p.indRight = Number(right);
    if (attr(ind, 'firstLine') != null) p.indFirst = Number(attr(ind, 'firstLine'));
    if (attr(ind, 'hanging') != null) p.indFirst = -Number(attr(ind, 'hanging'));
  }
  const numPr = kid(pPr, 'numPr');
  if (numPr) {
    const numId = kid(numPr, 'numId');
    const ilvl = kid(numPr, 'ilvl');
    if (numId) p.numId = attr(numId, 'val');
    p.ilvl = ilvl ? Number(attr(ilvl, 'val')) : 0;
  }
  const outline = kid(pPr, 'outlineLvl');
  if (outline) p.outlineLvl = Number(attr(outline, 'val'));
  if (onOff(kid(pPr, 'pageBreakBefore'))) p.pageBreakBefore = true;
  const bdr = kid(pPr, 'pBdr');
  if (bdr) p.borders = readBorders(bdr);
  const shd = shdFill(kid(pPr, 'shd'));
  if (shd) p.shading = shd;
  const kn = onOff(kid(pPr, 'keepNext'));
  if (kn != null) p.keepNext = kn;
  const kl = onOff(kid(pPr, 'keepLines'));
  if (kl != null) p.keepLines = kl;
  const wc = onOff(kid(pPr, 'widowControl'));
  if (wc != null) p.widowControl = wc;
  const nh = onOff(kid(pPr, 'suppressAutoHyphens'));
  if (nh != null) p.noHyph = nh;
  const tabs = kid(pPr, 'tabs');
  if (tabs) p.tabs = kids(tabs).filter((t) => t.localName === 'tab').map((t) => ({ val: attr(t, 'val'), pos: Number(attr(t, 'pos')), leader: attr(t, 'leader') || null }));
  const cs = onOff(kid(pPr, 'contextualSpacing'));
  if (cs != null) p.contextualSpacing = cs;
  return p;
}

// --- CSS generation -------------------------------------------------------

export const twipsToPt = (t) => t / 20;
/**
 * Character scale (w:w, percent): Word widens each character's advance. CSS can't stretch glyphs and still wrap, so
 * the extra goes between the characters (an average advance of 0.55em), which keeps the lines' widths and breaks.
 */
// Edge shapes text at its font size cut to 1/100px (11pt, 14.667px, sets as 14.66px): a line comes out 0.45 per mille
// narrower than the font's advances, which Word breaks lines by (a model of 1700 near ties missed 3),
// so a word Word wraps (by 0.3px in 08-headers-footers) still fits. The missing share comes back as letter spacing: the
// size's cut-off fraction times an average advance (0.43em: at 0.48 lines came out 0.06 per mille wide).
export function sizeCss(pt) {
  const px = (pt * 4) / 3;
  const ls = (px - Math.floor(px * 100 + 1e-6) / 100) * 0.43;
  return `font-size:${pt}pt;letter-spacing:${ls > 1e-4 ? ls.toFixed(4) : 0}px`; // (0 too: else a size's own would pass to text of another size)
}
export const scaleCss = (pct) => `letter-spacing:${+(((Number(pct) - 100) / 100) * 0.55).toFixed(4)}em`;

const HIGHLIGHT = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff',
  red: '#ff0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000',
  darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080',
  lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
};
// (A document's values reach CSS only as these allow: a colour as six hex digits, a highlight by name, a font name
// inside a quoted string it can't end. Anything else would let a crafted file write its own CSS into the page.)
export const highlightColor = (v) => HIGHLIGHT[v] || (/^[a-z]{3,20}$/i.test(v || '') ? v : 'transparent');
/** A colour value (w:color, w:fill...) when it's one: six hex digits; else null. */
export const hexColor = (v) => (/^[0-9a-f]{6}$/i.test(v || '') ? v : null);
/** A font name as a quoted CSS string: what could end the string or the declaration taken out. */
const quoted = (name) => `"${String(name).replace(/["\\\r\n\f{}<>]/g, '')}"`;

// Metric-compatible open fonts first-class. DejaVu Sans (Bitstream Vera's
// metrics) is what Word itself swaps for Verdana when the font is missing.
const SUBST = { Calibri: 'Carlito', 'Calibri Light': 'Carlito', Cambria: 'Caladea', Arial: 'Liberation Sans', 'Times New Roman': 'Liberation Serif', 'Courier New': 'Liberation Mono', Consolas: 'odx-consolas', Aptos: 'odx-aptos', 'Century Schoolbook': 'Century', 'Aptos Narrow': 'odx-aptos-narrow', 'Aptos Display': 'odx-aptos-display', 'DejaVu Sans': 'Verdana', 'Arial Unicode MS': 'Segoe UI Symbol', 'Open Sans': 'odx-open-sans',
  // Seen in Word's PDFs, not derived: a display font Word lacks, which it sets in Cambria (the I DO CREW forms, 8 files).
  "Gabriel Weiss' Friends Font": 'Cambria' };
// A font list (some generators write "Liberation Sans;Arial") names no font Word knows. One led by a Liberation font is that font's Microsoft twin (Word
// knows the pair: "Liberation Sans;Arial" prints in Arial), but "Open Sans;Arial" Word sets in Segoe UI, its UI font.
const listed = (name) => /^open sans;/i.test(name || '');
/** A list led by a Liberation font ("Liberation Sans;Arial"): the font Word sets it in, its twin named after it. */
const twin = (name) => (/^liberation [^;]*;/i.test(name || '') ? name.split(';')[1].trim() || null : null);
// The open document's font table stand-ins (fonts.js fontAlts): font name -> the font Word uses for it when missing.
const ALTS = new Map();
// Merged in, not replaced: two editors on one page each keep their documents' stand-ins (a font name's stand-in comes
// from its own font table entry, so one name rarely has two).
export function setFontAlts(map) { for (const [k, v] of map || []) ALTS.set(k, v); }

export function fontStack(name) {
  if (listed(name)) return '"Segoe UI", sans-serif';
  if (twin(name)) return fontStack(twin(name));
  if (!LINE[(name || '').toLowerCase()] && face(name)) return [quoted(name), ...fontStack(face(name)).split(', ')].join(', ');
  const alt = !LINE[(name || '').toLowerCase()] && !SUBST[name] && ALTS.get(name);
  if (alt) return [quoted(name), ...fontStack(alt).split(', ')].join(', ');
  const generic = /mono|courier|consol/i.test(name) ? 'monospace' : /times|cambria|georgia|garamond|serif/i.test(name) && !/sans/i.test(name) ? 'serif' : 'sans-serif';
  return [quoted(name), SUBST[name] && quoted(SUBST[name]), generic].filter(Boolean).join(', ');
}

// Word's "single" line height is the font's own ascent + descent (+ line gap),
// so it differs per font. Units per em from each font's OS/2 and hhea tables.
const LINE = {
  calibri: 1.2207, 'calibri light': 1.2207, carlito: 1.2207, cambria: 1.1724, caladea: 1.1724,
  arial: 1.1499, 'liberation sans': 1.1499, helvetica: 1.1499, 'times new roman': 1.1499,
  'liberation serif': 1.1499, consolas: 1.1709, 'courier new': 1.1328, 'liberation mono': 1.1328,
  'segoe ui': 1.3301, 'segoe ui light': 1.3301, 'segoe ui semibold': 1.3301, verdana: 1.2153, tahoma: 1.2070,
  georgia: 1.1362, aptos: 1.2207, 'aptos display': 1.2207, 'century schoolbook': 1.2021, century: 1.2021, 'aptos narrow': 1.2207, 'open sans': 1.3618, 'noto sans': 1.362, 'noto serif': 1.362,
  ubuntu: 1.149, 'arial unicode ms': 1.3398, roboto: 1.1719, lato: 1.2, candara: 1.2207, corbel: 1.2207, constantia: 1.2207,
  'trebuchet ms': 1.1621, 'comic sans ms': 1.3936, impact: 1.2197, 'arial black': 1.4097,
  'ms gothic': 1.3, // (not its 1.0em ascent + descent: Word's own layout gives a checkbox line 1.30em, 15.6pt over sample S072's 12pt Times and 13.0pt over the contingency plan's 10pt Calibri)
};
const measured = new Map();
/** ascent + descent per em of a font the table doesn't know, measured in the browser if it is installed. */
function measure(key) {
  if (measured.has(key)) return measured.get(key);
  let f = null;
  try {
    const ctx = document.createElement('canvas').getContext('2d');
    const probe = 'mmmmmmmmmmlli10WQ';
    const installed = ['monospace', 'serif'].some((g) => {
      ctx.font = `100px ${g}`;
      const base = ctx.measureText(probe).width;
      ctx.font = `100px "${key}", ${g}`;
      return ctx.measureText(probe).width !== base;
    });
    if (installed) {
      const m = ctx.measureText('Hg');
      if (m.fontBoundingBoxAscent) f = Math.round((m.fontBoundingBoxAscent + m.fontBoundingBoxDescent) * 100) / 10000;
    }
  } catch { /* no DOM */ }
  measured.set(key, f);
  return f;
}
// A face named as a font ("Times New Roman Bold", sample S078's headings): its family's metrics.
const face = (name) => { const m = /^(.+?) (bold|italic|bold italic)$/i.exec(name || ''); return m && LINE[m[1].toLowerCase()] ? m[1] : null; };
export const lineFactor = (name) => {
  if (twin(name)) return lineFactor(twin(name));
  if (!LINE[(name || '').toLowerCase()] && face(name)) return lineFactor(face(name));
  const key = listed(name) ? 'segoe ui' : (name || '').toLowerCase();
  return LINE[key] || (key && typeof document !== 'undefined' && measure(key)) || LINE[(SUBST[name] || '').toLowerCase()] || (ALTS.get(name) && ALTS.get(name) !== name ? lineFactor(ALTS.get(name)) : 0) || 1.17;
};
/** CSS for a font: the family stack plus the custom property paragraph line heights scale. */
export const fontCss = (name) => `font-family:${fontStack(name)};--odx-lh:${lineFactor(name)}`;

export function rPrToCss(p, theme) {
  const css = [];
  if (p.bold != null) css.push(`font-weight:${p.bold ? 'bold' : 'normal'}`);
  if (p.italic != null) css.push(`font-style:${p.italic ? 'italic' : 'normal'}`);
  const deco = [];
  if (p.underline && p.underline !== 'none') deco.push('underline');
  if (p.strike) deco.push('line-through');
  if (deco.length) css.push(`text-decoration:${deco.join(' ')}`);
  if (p.size) css.push(sizeCss(p.size / 2));
  if (p.color) css.push(`color:#${p.color}`);
  if (p.highlight) css.push(`background-color:${highlightColor(p.highlight)}`);
  const font = p.font || (p.themeFont && theme && theme[p.themeFont.startsWith('major') ? 'major' : 'minor']);
  if (font) css.push(fontCss(font));
  if (p.caps) css.push('text-transform:uppercase');
  if (p.smallCaps) css.push('font-variant:small-caps');
  // w:kern is a threshold (half-points): Word kerns text of that size or larger (a heading style's kern 28 leaves its
  // 12pt runs unkerned). Where the size is known with it.
  if (p.kern != null) css.push(`font-kerning:${p.kern && (p.size == null || p.size >= p.kern) ? 'normal' : 'none'}`);
  if (p.scale && p.scale !== 100) css.push(scaleCss(p.scale));
  if (p.hidden) css.push('display:none');
  css.push(...csVars(p, theme));
  return css.join(';');
}

/** The complex-script run properties as custom properties, which a complex-script run (.odx-cs, editor.css) applies. */
export function csVars(p, theme) {
  const css = [];
  if (p.boldCs != null) css.push(`--odx-bcs:${p.boldCs ? 'bold' : 'normal'}`);
  if (p.italicCs != null) css.push(`--odx-ics:${p.italicCs ? 'italic' : 'normal'}`);
  if (p.sizeCs) css.push(`--odx-szcs:${p.sizeCs / 2}pt`);
  const font = p.fontCs || (p.themeFontCs && theme && theme[p.themeFontCs.startsWith('major') ? 'majorCs' : 'minorCs']);
  if (font) css.push(`--odx-fcs:${fontStack(font)}`);
  return css;
}

export function pPrToCss(p) {
  const css = [];
  // (Start and end, not left and right: a right-to-left paragraph's 'left' is its start, at the right.)
  if (p.bidi != null) css.push(`direction:${p.bidi ? 'rtl' : 'ltr'}`);
  // (Distributed, w:jc="distribute" / thaiDistribute: its last line spread to both margins as well, as Word does.)
  if (p.align === 'distribute' || p.align === 'thaiDistribute') css.push('text-align-last:justify');
  if (p.align) css.push(`text-align:${{ both: 'justify', distribute: 'justify', thaiDistribute: 'justify', lowKashida: 'justify', mediumKashida: 'justify', highKashida: 'justify', left: 'start', right: 'end', center: 'center', start: 'start', end: 'end', justify: 'justify' }[p.align] || 'start'}`);
  if (p.before != null) css.push(`margin-top:${twipsToPt(p.before)}pt;--odx-sb:${twipsToPt(p.before)}pt`); // (--odx-sb: where a picture's paragraph-relative offset starts, reader.js)
  if (p.after != null) css.push(`margin-bottom:${twipsToPt(p.after)}pt`);
  if (p.line != null) {
    // auto: a multiple of the font's single line; exact: fixed; atLeast: the larger of the two.
    if (!p.lineRule || p.lineRule === 'auto') css.push(`--odx-lhv:calc(var(--odx-lh, 1.17) * ${+(p.line / 240).toFixed(4)});--odx-mult:${+(p.line / 240).toFixed(4)};--odx-runmult:${+(p.line / 240).toFixed(4)}`);
    else if (p.lineRule === 'exact') css.push(`--odx-lhv:${twipsToPt(p.line)}pt;--odx-mult:1;--odx-runmult:0;--odx-lhlen:1;min-height:0`); // runs in other fonts can't make it taller, nor the empty-paragraph floor (sample S085's 11.2pt headings in 10pt Times)
    else css.push(`--odx-lhv:max(${twipsToPt(p.line)}pt, calc(var(--odx-lh, 1.17) * 1em));--odx-mult:1;--odx-runmult:1;--odx-lhlen:1;--odx-atleast:${twipsToPt(p.line)}pt`);
  }
  // A side border and its space sit outside the indent, in Word: the text stays where the indent puts it (--odx-bdl/-bdr).
  if (p.indLeft != null) css.push(`margin-inline-start:calc(${twipsToPt(p.indLeft)}pt - var(--odx-bdl, 0pt))`);
  if (p.indRight != null) css.push(`margin-inline-end:calc(${twipsToPt(p.indRight)}pt - var(--odx-bdr, 0pt))`);
  if (p.indFirst != null) css.push(`text-indent:${twipsToPt(p.indFirst)}pt`);
  if (p.borders) {
    for (const side of ['top', 'bottom', 'left', 'right']) {
      const b = p.borders[side];
      if (!b) continue;
      const bc = borderCss(b);
      css.push(`border-${side}:${bc}`);
      // Edge draws a border a whole number of pixels wide (under 1px as 1, else rounded down): a 1pt (1.333px) rule takes
      // 1px. Word's takes its own width, so the difference goes to the padding (doc 05's Heading 1 rule: the text under it
      // 0.25pt high in Word's own layout).
      const solid = / solid /.test(bc);
      const px = solid && (side === 'top' || side === 'bottom') ? parseFloat(bc) * 4 / 3 : 0;
      const adj = px ? (Math.max(0.25, b.sz / 8) * 4) / 3 - (px < 1 ? 1 : Math.floor(px + 1e-6)) : 0;
      if (adj) css.push(`padding-${side}:${+Math.max(0, ((b.space || 0) * 4) / 3 + adj).toFixed(3)}px`);
      else if (b.space) css.push(`padding-${side}:${b.space}pt`);
      if ((side === 'left' || side === 'right') && bc !== 'none') css.push(`--odx-bd${side[0]}:${parseFloat(bc) + (b.space || 0)}pt`);
    }
  }
  if (p.shading) css.push(`background-color:${p.shading}`);
  return css.join(';');
}

export { kids };

/** Merge tab stop lists: later lists add stops and `clear` stops remove inherited ones. */
export function mergeTabs(...lists) {
  let out = [];
  for (const l of lists) {
    if (!l) continue;
    for (const t of l) {
      out = out.filter((o) => Math.abs(o.pos - t.pos) > 1);
      if (t.val !== 'clear') out.push(t);
    }
  }
  return out.sort((a, b) => a.pos - b.pos);
}
