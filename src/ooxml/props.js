// Readers for run (w:rPr) and paragraph (w:pPr) properties into plain objects,
// shared by direct formatting and style definitions. Units stay OOXML-native
// (half-points for size, twips for spacing/indent) until CSS is generated.
import { kid, kids, attr, onOff } from './xml.js';
import { OMF } from '../fonts/omf.js';
import { drawsOwn } from './fontprobe.js';

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
  const csp = kid(rPr, 'spacing'); // character spacing, twips (a style's: Title's -10)
  if (csp) p.charSpacing = Number(attr(csp, 'val')) || 0;
  // Word 2010's outer shadow (a style's: the room it takes on the line, texteffects.js shadowRoom). Numbers, bounded there.
  const sh = kid(rPr, 'shadow', W14);
  if (sh) p.shadow14 = { blur: Number(attr(sh, 'blurRad', W14)) || 0, dist: Number(attr(sh, 'dist', W14)) || 0, dir: Number(attr(sh, 'dir', W14)) || 0, flat: attr(sh, 'sx', W14) === '0' && attr(sh, 'sy', W14) === '0' };
  return p;
}
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

/** A w:pBdr / w:tblBorders / w:tcBorders element -> { top: {val, sz, color, space}, ... } */
export function readBorders(el) {
  const out = {};
  if (!el) return out;
  for (const c of kids(el)) {
    const side = { start: 'left', end: 'right' }[c.localName] || c.localName;
    // (Size and space bounded as Word bounds them: a size of eighths of a point up to 96, an art border's points up to
    // 31, so at most 255 here; space up to 31pt. A file's huge size made a paragraph millions of pixels tall.)
    out[side] = { val: attr(c, 'val'), sz: Math.min(255, Number(attr(c, 'sz')) || 4), color: hexColor(attr(c, 'color')), space: Math.min(31, Number(attr(c, 'space')) || 0) };
  }
  return out;
}

// Borders of two or three lines: their whole width (pt) from their size (pt), as Word lays them out (hdrgap-word: an
// empty header paragraph with each bordered below it at 0.75pt to 6pt, the body that much lower in Word's PDF, within
// its 0.12pt steps). A double's size is each line, with a gap as wide between (doc 01: 2.88px lines, 2.88px apart, for
// sz 18); a triple's, its three lines and two gaps. A thick-thin pair's size is its thick line: with a small gap, the
// thin line and the gap add 0.75pt each; a medium one, half the size each; a large one, 0.75pt and 1.5pt. Thin-thick-thin
// the same on both sides, but a large gap's thin lines and gaps adding 3pt and the size again.
const COMPOUND = {
  double: (w) => 3 * w,
  triple: (w) => 5 * w,
  thinThickSmallGap: (w) => w + 1.5,
  thickThinSmallGap: (w) => w + 1.5,
  thinThickMediumGap: (w) => 2 * w,
  thickThinMediumGap: (w) => 2 * w,
  thinThickLargeGap: (w) => w + 2.25,
  thickThinLargeGap: (w) => w + 2.25,
  thinThickThinSmallGap: (w) => w + 3,
  thinThickThinMediumGap: (w) => 3 * w,
  thinThickThinLargeGap: (w) => 2 * w + 3,
};

// Their lines and gaps (pt, line first) as Word prints them, in the page's order whichever side they're on: top to
// bottom across the page, left to right down it (thinbdr-word: boxes of each at 0.25pt to 1.5pt, Word's PDF; a
// thin-thick border's size line above or left of its thin one on all four sides; bdr3-word: 3pt to 6pt the same, and
// a large gap's lines 1.5pt and 0.75pt at every size, the size its gap). Their sums are COMPOUND's.
const STRIPES = {
  double: (w) => [w, w, w],
  triple: (w) => [w, w, w, w, w],
  thinThickSmallGap: (w) => [w, 0.75, 0.75],
  thickThinSmallGap: (w) => [0.75, 0.75, w],
  thinThickMediumGap: (w) => [w, w / 2, w / 2],
  thickThinMediumGap: (w) => [w / 2, w / 2, w],
  thinThickLargeGap: (w) => [1.5, w, 0.75],
  thickThinLargeGap: (w) => [0.75, w, 1.5],
  thinThickThinSmallGap: (w) => [0.75, 0.75, w, 0.75, 0.75],
  thinThickThinMediumGap: (w) => [w / 2, w / 2, w, w / 2, w / 2],
  thinThickThinLargeGap: (w) => [0.75, w, 1.5, w, 0.75],
};

const BROKEN = { dotted: 'dotted', dashed: 'dashed', dashSmallGap: 'dashed', dotDash: 'dashed' };

const drawnSide = (b) => b && b.val && b.val !== 'nil' && b.val !== 'none';

/**
 * Borders as readBorders gives them, whatever came in: a co-editor's step carries any JSON in a paragraph's attributes,
 * so sizes and spaces made numbers within Word's bounds, colours six hex digits or none, kinds strings. (A space given
 * as a string made the CSS's arithmetic throw, and the paragraph couldn't be drawn.)
 */
export function cleanBorders(borders) {
  const out = {};
  if (!borders || typeof borders !== 'object') return out;
  for (const side of ['top', 'bottom', 'left', 'right', 'between', 'bar']) {
    const b = Object.hasOwn(borders, side) ? borders[side] : null;
    if (!b || typeof b !== 'object') continue;
    const num = (v, max, d) => { const n = Number(v); return Number.isFinite(n) && n >= 0 ? Math.min(max, n) : d; };
    out[side] = { val: typeof b.val === 'string' ? b.val : null, sz: num(b.sz, 255, 4) || 4, color: hexColor(b.color), space: num(b.space, 31, 0) };
  }
  return out;
}

/**
 * Paragraph borders with a compound side (CSS double draws two equal lines, none under 2.25pt, a triple as two): all
 * their sides drawn by one picture (a border-image cut in nine, its edges stretched along the sides), each side's lines as
 * Word prints them, each corner's lines joined one inside the other, as Word's are. Null when there's none, or a side
 * the picture can't draw (dotted, dashed: those keep CSS's own). Returns the image's CSS and each side's width (px, whole:
 * the browser draws a border's width in whole pixels).
 */
function stripedBorders(borders) {
  const sides = ['top', 'right', 'bottom', 'left'];
  const has = (s) => drawnSide(borders[s]);
  if (!sides.some((s) => has(s) && Object.hasOwn(COMPOUND, borders[s].val))) return null;
  if (sides.some((s) => has(s) && Object.hasOwn(BROKEN, borders[s].val))) return null;
  const width = {}, lines = {}, color = {};
  for (const s of sides) {
    const b = borders[s];
    if (!has(s)) continue;
    const segs = Object.hasOwn(STRIPES, b.val) ? STRIPES[b.val](b.sz / 8) : [Math.max(0.25, b.sz / 8)];
    const total = segs.reduce((a, v) => a + v, 0);
    width[s] = total * 4 / 3;
    // Each line as [start, length], fractions of the side's width from its top or left edge.
    lines[s] = [];
    for (let i = 0, at = 0; i < segs.length; at += segs[i], i++) if (i % 2 === 0) lines[s].push([at / total, segs[i] / total]);
    color[s] = /^[0-9A-Fa-f]{6}$/.test(b.color || '') ? '#' + b.color : '#000'; // (only a colour goes into the picture)
  }
  // The picture: 30 units square, cut at 10 and 20 (border-image-slice 10), its middle not drawn.
  const f = (v) => +v.toFixed(3);
  const rect = (x, y, w, h, c) => `<rect x="${f(x)}" y="${f(y)}" width="${f(w)}" height="${f(h)}" fill="${c}"/>`;
  let svg = '';
  for (const [s, x0, y0, across] of [['top', 10, 0, true], ['bottom', 10, 20, true], ['left', 0, 10, false], ['right', 20, 10, false]]) {
    for (const [a, l] of lines[s] || []) svg += across ? rect(x0, y0 + 10 * a, 10, 10 * l, color[s]) : rect(x0 + 10 * a, y0, 10 * l, 10, color[s]);
  }
  // Corners: the lines of its two sides from the outside in (`flip`: a bottom or right side's run from the inside out
  // in page order); paired one to one, each pair an L whose arms run from its partner's outer edge into the side;
  // otherwise each side's lines straight across.
  const out = (s, flip) => (flip ? lines[s].map(([a, l]) => [1 - a - l, l]).reverse() : lines[s]);
  for (const [h, v, cx, cy] of [['top', 'left', 0, 0], ['top', 'right', 20, 0], ['bottom', 'left', 0, 20], ['bottom', 'right', 20, 20]]) {
    if (!lines[h] || !lines[v]) continue;
    const H = out(h, h === 'bottom'), V = out(v, v === 'right');
    // From the corner's outer point: u along the page's width, t down it, in the corner's own fractions.
    const put = (u0, t0, du, dt, c) => {
      const x = v === 'right' ? 1 - u0 - du : u0, y = h === 'bottom' ? 1 - t0 - dt : t0;
      svg += rect(cx + 10 * x, cy + 10 * y, 10 * du, 10 * dt, c);
    };
    for (let i = 0; i < Math.max(H.length, V.length); i++) {
      const paired = H.length === V.length;
      if (H[i]) put(paired ? V[i][0] : 0, H[i][0], paired ? 1 - V[i][0] : 1, H[i][1], color[h]);
      if (V[i]) put(V[i][0], paired ? H[i][0] : 0, V[i][1], paired ? 1 - H[i][0] : 1, color[v]);
    }
  }
  const data = `<svg xmlns="http://www.w3.org/2000/svg" width="30" height="30" viewBox="0 0 30 30" preserveAspectRatio="none">${svg}</svg>`;
  return { width, color, css: `border-image:url('data:image/svg+xml,${encodeURIComponent(data).replace(/'/g, '%27')}') 10 stretch` };
}

/**
 * A compound border side (double, triple, thick-thin...) as Word prints it: its lines ([start, length], fractions of its
 * width, in the page's order), its width (px) and colour (#rrggbb). Null for any other side.
 */
export function compoundSide(b) {
  if (!b || typeof b !== 'object' || !drawnSide(b) || !Object.hasOwn(STRIPES, b.val)) return null;
  const n = Number(b.sz);
  const segs = STRIPES[b.val]((Number.isFinite(n) && n > 0 ? Math.min(255, n) : 4) / 8);
  const total = segs.reduce((a, v) => a + v, 0);
  const lines = [];
  for (let i = 0, at = 0; i < segs.length; at += segs[i], i++) if (i % 2 === 0) lines.push([at / total, segs[i] / total]);
  return { lines, px: (total * 4) / 3, color: /^[0-9A-Fa-f]{6}$/.test(b.color || '') ? '#' + b.color : '#000000' };
}

/**
 * Where two compound borders meet, as Word prints the square they share (bdr3-word, tables of each: a cross, a T, a
 * corner): its lines H across (fractions down the square) and V down it (fractions across), and which arms the junction
 * has (l, r: the line across goes on to the left, right; u, d: the line down goes on up, down). Each line comes in from
 * its arm and stops at the far edge of the line it turns into: a line on the upper side of the middle turns up, on the
 * lower side down (the middle one both ways), its partner the line as far from that corner; with no such arm it goes
 * straight through, or, with nothing beyond, turns the other way (a corner's nested Ls). Rects [x, y, w, h, 'h' | 'v'].
 */
export function junctionRects(H, V, arms) {
  const out = [];
  // Along one line from one of its arms: [from, to] across the square.
  const reach = (i, n, m, P, fromHigh, sides, through) => {
    const mid = (n - 1) / 2;
    let turns = (i < mid ? [sides[0]] : i > mid ? [sides[1]] : sides).filter((t) => arms[t]);
    if (!turns.length) {
      if (arms[through]) return [0, 1];
      turns = sides.filter((t) => arms[t]);
      if (!turns.length) return [0, 1];
    }
    let lo = 1, hi = 0;
    for (const t of turns) {
      const p = fromHigh ? (t === sides[0] ? m - 1 - i : m - 1 - (n - 1 - i)) : (t === sides[0] ? i : n - 1 - i);
      if (p < 0 || p >= m) return [0, 1];
      lo = Math.min(lo, P[p][0]);
      hi = Math.max(hi, P[p][0] + P[p][1]);
    }
    return fromHigh ? [lo, 1] : [0, hi];
  };
  const along = (L, P, hiArm, loArm, sides, kind) => L.forEach(([at, len], i) => {
    const parts = [];
    if (arms[hiArm]) parts.push(reach(i, L.length, P.length, P, true, sides, loArm));
    if (arms[loArm]) parts.push(reach(i, L.length, P.length, P, false, sides, hiArm));
    if (parts.length === 2 && parts[1][1] >= parts[0][0]) parts.splice(0, 2, [0, 1]);
    for (const [a, b] of parts) if (b > a) out.push(kind === 'h' ? [a, at, b - a, len, 'h'] : [at, a, len, b - a, 'v']);
  });
  along(H, V, 'r', 'l', ['u', 'd'], 'h');
  along(V, H, 'd', 'u', ['l', 'r'], 'v');
  return out;
}

/**
 * A table cell's compound border lines, drawn as Word prints them by pictures behind the cell (border-collapse leaves no
 * border-image): each side's lines along its edge, and a square where two meet (junctionRects). `sides` are the sides this
 * cell draws (compoundSide; a line between rows is the lower row's, between cells the right-hand one's, as Word puts
 * them: from the grid line in); `near` the borders meeting its corners from outside it: tl and bl the left neighbour's
 * top and bottom, tu and ru the left and right lines of the cell above. Returns the background CSS and the same lines as
 * `--odx-lines` (each rect x, y, width, height as a share of the cell's width or height plus px, then its colour) for the
 * PDF export; null when there's nothing to draw.
 */
export function cellLines(sides, near) {
  const { top: T, bottom: B, left: L, right: R } = sides;
  if (!T && !B && !L && !R) return null;
  const layers = [];
  // A layer: anchored left or right, top or bottom, whole px from there (the browser puts a cell's pictures on whole
  // pixels: a 4.33px square beside a strip 4.33px in left a pixel's gap); each way either a fixed size (px, drawn in the
  // next whole px up, its lines reaching the inner edge carried on to it) or the cell's less some px. Its rects in
  // fractions of that.
  const layer = (ax, ox, ay, oy, w, h, rects) => { if (rects.length) layers.push({ ax, ox, ay, oy, w, h, rects }); };
  const up = (v) => Math.ceil(v - 1e-6);
  const tlV = L || (T && near.tu) || null;
  const tl = T && tlV ? junctionRects(T.lines, tlV.lines, { l: !!near.tl, r: true, u: !!near.tu, d: !!L }) : null;
  const tr = T && R ? junctionRects(T.lines, R.lines, { l: true, r: false, u: !!near.ru, d: true }) : null;
  const bl = B && L ? junctionRects(B.lines, L.lines, { l: !!near.bl, r: true, u: true, d: false }) : null;
  const br = B && R ? junctionRects(B.lines, R.lines, { l: true, r: false, u: true, d: false }) : null;
  const paint = (rects, hc, vc) => rects.map(([x, y, w, h, k]) => [x, y, w, h, k === 'h' ? hc : vc]);
  const across = (S) => S.lines.map(([a, l]) => [0, a, 1, l, S.color]);
  const down = (S) => S.lines.map(([a, l]) => [a, 0, l, 1, S.color]);
  const fix = (px) => ({ fix: px }), cut = (px) => ({ cut: px });
  if (T) {
    const a = tl ? up(tlV.px) : 0, b = tr ? up(R.px) : 0;
    if (tl) layer('l', 0, 't', 0, fix(tlV.px), fix(T.px), paint(tl, T.color, tlV.color));
    layer('l', a, 't', 0, cut(a + b), fix(T.px), across(T));
    if (tr) layer('r', 0, 't', 0, fix(R.px), fix(T.px), paint(tr, T.color, R.color));
  }
  if (B) {
    const a = bl ? up(L.px) : 0, b = br ? up(R.px) : 0;
    if (bl) layer('l', 0, 'b', 0, fix(L.px), fix(B.px), paint(bl, B.color, L.color));
    layer('l', a, 'b', 0, cut(a + b), fix(B.px), across(B));
    if (br) layer('r', 0, 'b', 0, fix(R.px), fix(B.px), paint(br, B.color, R.color));
  }
  const t = T ? up(T.px) : 0, b = B ? up(B.px) : 0;
  if (L) layer('l', 0, 't', t, fix(L.px), cut(t + b), down(L));
  if (R) layer('r', 0, 't', t, fix(R.px), cut(t + b), down(R));
  if (!layers.length) return null;
  const f = (v) => +v.toFixed(3);
  const images = [], pos = [], size = [], lines = [];
  for (const { ax, ox, ay, oy, w, h, rects } of layers) {
    // Each way: the layer's size [share of the cell, px], where it starts [share, px], and a rect's [start, length] in
    // the picture's own units (px when fixed, hundredths when stretched).
    const axis = (d, far, o) => {
      const sz = d.fix != null ? [0, up(d.fix)] : [1, -d.cut];
      const at = far ? [1 - sz[0], -o - sz[1]] : [0, o];
      const place = (x, l) => {
        if (d.fix == null) return [100 * x, 100 * l];
        let p = x * d.fix + (far ? sz[1] - d.fix : 0), n = l * d.fix;
        if (!far && x + l >= 1 - 1e-6) n = sz[1] - p;
        if (far && x <= 1e-6) { n += p; p = 0; }
        return [p, n];
      };
      return { sz, at, place, units: d.fix != null ? sz[1] : 100, css: d.fix != null ? `${sz[1]}px` : d.cut ? `calc(100% - ${f(d.cut)}px)` : '100%' };
    };
    const X = axis(w, ax === 'r', ox), Y = axis(h, ay === 'b', oy);
    let svg = '';
    for (const [x, y, rw, rh, c] of rects) {
      const [px, pw] = X.place(x, rw), [py, ph] = Y.place(y, rh);
      svg += `<rect x="${f(px)}" y="${f(py)}" width="${f(pw)}" height="${f(ph)}" fill="${c}"/>`;
      // For the PDF: a share of the cell and px each way, from the cell's top left.
      const one = (A, p, n) => (A.units === 100 ? [A.at[0] + (p / 100) * A.sz[0], A.at[1] + (p / 100) * A.sz[1], (n / 100) * A.sz[0], (n / 100) * A.sz[1]] : [A.at[0], A.at[1] + p, 0, n]);
      const [x0, x1, w0, w1] = one(X, px, pw), [y0, y1, h0, h1] = one(Y, py, ph);
      lines.push([x0, x1, y0, y1, w0, w1, h0, h1].map(f).join(' ') + ' ' + c);
    }
    const data = `<svg xmlns="http://www.w3.org/2000/svg" width="${X.units}" height="${Y.units}" viewBox="0 0 ${X.units} ${Y.units}" preserveAspectRatio="none">${svg}</svg>`;
    images.push(`url('data:image/svg+xml,${encodeURIComponent(data).replace(/'/g, '%27')}')`);
    pos.push(`${ax === 'l' ? 'left' : 'right'} ${ox}px ${ay === 't' ? 'top' : 'bottom'} ${oy}px`);
    size.push(`${X.css} ${Y.css}`);
  }
  return {
    css: `background-image:${images.join(',')};background-position:${pos.join(',')};background-size:${size.join(',')};background-repeat:no-repeat`,
    lines: lines.join(','),
  };
}

/** CSS for one border side; Word sizes are eighths of a point. */
export function borderCss(b) {
  if (!b || !b.val || b.val === 'nil' || b.val === 'none') return 'none';
  const lines = Object.hasOwn(COMPOUND, b.val) ? COMPOUND[b.val] : null;
  const style = lines ? 'double' : Object.hasOwn(BROKEN, b.val) ? BROKEN[b.val] : 'solid';
  // A border of two or three lines is as wide as all of them and their gaps (COMPOUND), at least 2.25pt (3px) for the
  // browser to draw two lines.
  const w = lines ? Math.max(2.25, lines(b.sz / 8)) : Math.max(0.5, b.sz / 8);
  // (Only a number and six hex digits go into the CSS, whatever a co-editor's step carried.)
  return `${Number.isFinite(w) ? w : 0.5}pt ${style} ${hexColor(b.color) ? '#' + b.color : '#000'}`;
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
    // (Turned off outright, w:afterAutospacing="0", over a style that has it: the paragraph's own spacing, as in Word: a
    // web-style list's items keep their 7.5pt after between them.)
    if (/^(0|false|off)$/.test(attr(sp, 'beforeAutospacing') || '')) p.beforeAuto = false;
    if (/^(0|false|off)$/.test(attr(sp, 'afterAutospacing') || '')) p.afterAuto = false;
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
  // (Plus the character spacing in force, --odx-csp, which a run of its own size would otherwise lose.)
  return `font-size:${pt}pt;letter-spacing:calc(var(--odx-csp, 0px) + ${ls > 1e-4 ? ls.toFixed(4) : 0}px)`; // (0 too: else a size's own would pass to text of another size)
}
// (--odx-w: the scale itself, which widths.js makes exact for each run's own text once it's laid out.)
export const scaleCss = (pct) => `--odx-w:${Number(pct) / 100};letter-spacing:calc(var(--odx-csp, 0px) + ${+(((Number(pct) - 100) / 100) * 0.55).toFixed(4)}em)`;

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
// (These tables are looked up by font names from the document: no prototype, so "constructor" or "__proto__" find nothing.)
const SUBST = Object.assign(Object.create(null), { Calibri: 'Carlito', 'Calibri Light': 'Carlito', Cambria: 'Caladea', Arial: 'Liberation Sans', 'Times New Roman': 'Liberation Serif', 'Courier New': 'Liberation Mono', Consolas: 'odx-consolas', Aptos: 'odx-aptos', 'Century Schoolbook': 'Century', 'Aptos Narrow': 'odx-aptos-narrow', 'Aptos Display': 'odx-aptos-display', 'DejaVu Sans': 'Verdana', 'Arial Unicode MS': 'Segoe UI Symbol', 'Open Sans': 'odx-open-sans',
  // Seen in Word's PDFs, not derived: a display font Word lacks, which it sets in Cambria (the I DO CREW forms, 8 files).
  "Gabriel Weiss' Friends Font": 'Cambria' });
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
// The host's own substitutes (mountDocxEditor's fonts.substitutes, fontload.js): font name -> the family it wants drawn
// where that font is missing, before the editor's stand-ins.
const HOST = new Map();
export function setFontSubstitutes(map) { for (const [k, v] of map || []) HOST.set(k, v); }

// The open-metric-fonts stand-ins (src/fonts/omf.js: Grandview, Bahnschrift, Sitka, Aptos's weights...), where the host
// loads them (fontload.js): font name -> the family it registered for it, drawn where the font is missing, after the
// font itself and the editor's own stand-in (Aptos keeps odx-aptos, fitted to Word's prints: Morrow Sans draws the
// letters it lacks). Null when they're off (fonts.omf false) or outside a browser.
let omfFamily = null;
export function setOmf(fn) { omfFamily = fn; }
const OMF_KEY = new Map(Object.keys(OMF).map((k) => [k.toLowerCase(), k]));
const OMF_LINE = new Map(Object.entries(OMF).map(([k, v]) => [k.toLowerCase(), v[1]]));

export function fontStack(name) {
  const host = HOST.get(name);
  const stack = host && host !== name ? [quoted(name), quoted(host), ...baseStack(name).split(', ').slice(1)].join(', ') : baseStack(name);
  const key = omfFamily && OMF_KEY.get(String(name || '').toLowerCase());
  const omf = key && omfFamily(key);
  if (!omf) return stack;
  // (Before the generic family at the end.)
  const parts = stack.split(', ');
  parts.splice(parts.length - 1, 0, quoted(omf));
  return parts.join(', ');
}

// Names Windows itself maps to another font (its FontSubstitutes), which Word prints in that font: "Times" is Times New
// Roman in Word's PDFs, where the browser's "Times" is another face with shorter lines (1.11 em, not 1.15).
const WINDOWS = Object.assign(Object.create(null), { times: 'Times New Roman', courier: 'Courier New' }); // (Courier: Courier New in Word's PDFs, 1.133 em a line; the browser's own "Courier" measured 1.15)
const windowsFont = (name) => WINDOWS[String(name || '').toLowerCase()] || null;
// Fonts Word draws in another where they're missing (the fonts in its PDFs of documents naming each): drawn in that
// font here too, after the font itself, unless the document's own font table names a stand-in.
const WORD_ALTS = Object.assign(Object.create(null), { 'tms rmn': 'Times New Roman', 'times-roman': 'Times New Roman', timesnewromanpsmt: 'Times New Roman', 'cg times (w1)': 'Times New Roman', 'new york': 'Times New Roman',
  'helvetica neue': 'Arial', 'calibri (body)': 'Calibri', palatino: 'Palatino Linotype', 'futura md bt': 'Century Gothic', 'futura lt bt': 'Century Gothic',
  'lucida grande': 'Segoe UI', 'charter bt': 'Cambria', inter: 'Cambria',
  // (Roboto Condensed, not installed: Calibri in Word's print of a sample, its bold heading 146.8px wide as Calibri Bold's
  // 146.6px, though the file's font table gives Arial as its other name: 187.2px.)
  'roboto condensed': 'Calibri' });
// (Futura in Century Gothic: lines of 1.2261 em in Word's prints, not the 1.19 an installed Century Gothic measures.)
const WORD_ALT_LINE = Object.assign(Object.create(null), { 'futura md bt': 1.2261, 'futura lt bt': 1.2261 });
const altOf = (name) => WORD_ALTS[String(name || '').toLowerCase()] || ALTS.get(name) || null; // (what Word was seen to draw, over the file's own word for it)

function baseStack(name) {
  if (windowsFont(name)) return fontStack(windowsFont(name));
  if (listed(name)) return '"Segoe UI", sans-serif';
  if (twin(name)) return fontStack(twin(name));
  if (!LINE[(name || '').toLowerCase()] && face(name)) return [quoted(name), ...fontStack(face(name)).split(', ')].join(', ');
  const alt = !LINE[(name || '').toLowerCase()] && !SUBST[name] && altOf(name);
  if (alt) return [quoted(name), ...fontStack(alt).split(', ')].join(', ');
  const generic = /mono|courier|consol/i.test(name) ? 'monospace' : /times|cambria|georgia|garamond|serif/i.test(name) && !/sans/i.test(name) ? 'serif' : 'sans-serif';
  return [quoted(name), SUBST[name] && quoted(SUBST[name]), generic].filter(Boolean).join(', ');
}

// Word's "single" line height is the font's own ascent + descent (+ line gap),
// so it differs per font. Units per em from each font's OS/2 and hhea tables.
const LINE = Object.assign(Object.create(null), {
  calibri: 1.2207, 'calibri light': 1.2207, carlito: 1.2207, cambria: 1.1724, caladea: 1.1724,
  arial: 1.1499, 'liberation sans': 1.1499, helvetica: 1.1499, 'times new roman': 1.1499,
  'liberation serif': 1.1499, consolas: 1.1709, 'courier new': 1.1328, 'liberation mono': 1.1328,
  'segoe ui': 1.3301, 'segoe ui light': 1.3301, 'segoe ui semibold': 1.3301, verdana: 1.2153, tahoma: 1.2070,
  georgia: 1.1362, aptos: 1.2207, 'aptos display': 1.2207, 'century schoolbook': 1.2021, century: 1.2021, 'aptos narrow': 1.2207, 'open sans': 1.3618, 'noto sans': 1.362, 'noto serif': 1.362,
  ubuntu: 1.149, 'arial unicode ms': 1.3398, roboto: 1.1719, lato: 1.2, candara: 1.2207, corbel: 1.2207, constantia: 1.2207,
  // (Arial Narrow: its Windows ascent + descent, Word's 11.473pt at 10pt and 13.769pt at 12pt; measured in the browser
  // from its hhea it came out 1.13, lines of 1.08 and 1.15 spacing 0.2pt short each, 10px a page in the samples.)
  'arial narrow': 1.1474, 'trebuchet ms': 1.1621, 'comic sans ms': 1.3936, impact: 1.2197, 'arial black': 1.4097,
  'ms gothic': 1.3, // (not its 1.0em ascent + descent: Word's own layout gives a checkbox line 1.30em, 15.6pt over sample S072's 12pt Times and 13.0pt over the contingency plan's 10pt Calibri)
});
const measured = new Map();
/** ascent + descent per em of a font the table doesn't know, measured in the browser if it is installed. */
function measure(key) {
  if (measured.has(key)) return measured.get(key);
  let f = null;
  try {
    if (drawsOwn(key, { probe: 'mmmmmmmmmmlli10WQ', size: 100 })) {
      const ctx = document.createElement('canvas').getContext('2d');
      ctx.font = `100px "${key}", serif`;
      const m = ctx.measureText('Hg');
      if (m.fontBoundingBoxAscent) f = Math.round((m.fontBoundingBoxAscent + m.fontBoundingBoxDescent) * 100) / 10000;
    }
  } catch { /* no DOM */ }
  measured.set(key, f);
  return f;
}
// A face named as a font ("Times New Roman Bold", sample S078's headings): its family's metrics.
const face = (name) => { const m = /^(.+?) (bold|italic|bold italic)$/i.exec(name || ''); return m && (LINE[m[1].toLowerCase()] || OMF_LINE.has(m[1].toLowerCase())) ? m[1] : null; };
export const lineFactor = (name) => {
  // (A font missing here, drawn in the host's substitute: that font's line, unless the editor knows the missing one's.)
  if (HOST.has(name) && !LINE[(name || '').toLowerCase()] && !(typeof document !== 'undefined' && measure((name || '').toLowerCase()))) return lineFactor(HOST.get(name));
  if (windowsFont(name) && !HOST.has(name)) return lineFactor(windowsFont(name));
  if (twin(name)) return lineFactor(twin(name));
  if (!LINE[(name || '').toLowerCase()] && face(name)) return lineFactor(face(name));
  const key = listed(name) ? 'segoe ui' : (name || '').toLowerCase();
  // (A font open-metric-fonts stands in for, missing from the computer: its line as Word sets it, from the stand-in's
  // metrics, which are the original's. An installed one is measured as before: run 419 put Book Antiqua and Century
  // Gothic forms 6 lines off with the computed value ahead of the measure.)
  return LINE[key] || (key && typeof document !== 'undefined' && measure(key)) || OMF_LINE.get(key) || LINE[(SUBST[name] || '').toLowerCase()] || (!ALTS.get(name) && WORD_ALT_LINE[key]) || (altOf(name) && altOf(name) !== name ? lineFactor(altOf(name)) : 0) || 1.17;
};
/** CSS for a font: the family stack plus the custom property paragraph line heights scale. */
export const fontCss = (name) => `font-family:${fontStack(name)};--odx-lh:${lineFactor(name)}${faceCss(name)}`;
/**
 * A face named as a font ("Arial Bold", "Times New Roman Bold Italic") is drawn in that face, bold or italic not asked
 * for besides (Word's PDFs of such documents embed Arial-BoldMT alone): the weight and slant to add, or ''.
 */
export const faceCss = (name) => {
  if (LINE[(name || '').toLowerCase()] || !face(name)) return '';
  const w = /^.+? (bold|italic|bold italic)$/i.exec(name)[1].toLowerCase();
  return `${w.includes('bold') ? ';font-weight:bold' : ''}${w.includes('italic') ? ';font-style:italic' : ''}`;
};

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
  // Character spacing from a style (Title's 0.5pt tighter, compared with Word's PDF): kept in --odx-csp, which a run's
  // own size (sizeCss) and widths.js's correction add to.
  if (p.charSpacing != null) css.push(`--odx-csp:${p.charSpacing / 20}pt`, ...(p.size ? [] : ['letter-spacing:var(--odx-csp)']));
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
  // (Centred or set to the right: the spaces that end the paragraph take no part in placing its last line, editor.css.)
  if (p.align) css.push(`--odx-trail:${p.align === 'center' || p.align === 'right' || p.align === 'end' ? 1 : 0}`);
  if (p.align) css.push(`text-align:${{ both: 'justify', distribute: 'justify', thaiDistribute: 'justify', lowKashida: 'justify', mediumKashida: 'justify', highKashida: 'justify', left: 'start', right: 'end', center: 'center', start: 'start', end: 'end', justify: 'justify' }[p.align] || 'start'}`);
  if (p.before != null) css.push(`margin-top:${twipsToPt(p.before)}pt;--odx-sb:${twipsToPt(p.before)}pt`); // (--odx-sb: where a picture's paragraph-relative offset starts, reader.js)
  if (p.after != null) css.push(`margin-bottom:${twipsToPt(p.after)}pt`);
  if (p.line != null) {
    // auto: a multiple of the font's single line; exact: fixed; atLeast: the larger of the two.
    if (!p.lineRule || p.lineRule === 'auto') {
      css.push(`--odx-lhv:calc(var(--odx-lh, 1.17) * ${+(p.line / 240).toFixed(4)});--odx-mult:${+(p.line / 240).toFixed(4)};--odx-runmult:${+(p.line / 240).toFixed(4)}`);
      // Under a single line, an empty paragraph's floor (editor.css) by the same multiple: Word's empty line is that much
      // shorter (generated documents: ten empty paragraphs at 0.25 lines 3.8px apart in the body; one at 0.06 lines under a
      // footer's text taking 1px, the footer that much lower and the text above it a line longer).
      if (p.line > 0 && p.line < 240) css.push(`--odx-floor:${+(1.15 * p.line / 240).toFixed(4)}em`);
    }
    else if (p.lineRule === 'exact') css.push(`--odx-lhv:${twipsToPt(p.line)}pt;--odx-mult:1;--odx-runmult:0;--odx-lhlen:1;min-height:0`); // runs in other fonts can't make it taller, nor the empty-paragraph floor (sample S085's 11.2pt headings in 10pt Times)
    else css.push(`--odx-lhv:max(${twipsToPt(p.line)}pt, calc(var(--odx-lh, 1.17) * 1em));--odx-mult:1;--odx-runmult:1;--odx-lhlen:1;--odx-atleast:${twipsToPt(p.line)}pt`);
  }
  // A side border and its space sit outside the indent, in Word: the text stays where the indent puts it (--odx-bdl/-bdr).
  if (p.indLeft != null) css.push(`margin-inline-start:calc(${twipsToPt(p.indLeft)}pt - var(--odx-bdl, 0pt))`);
  if (p.indRight != null) css.push(`margin-inline-end:calc(${twipsToPt(p.indRight)}pt - var(--odx-bdr, 0pt))`);
  if (p.indFirst != null) css.push(`text-indent:${twipsToPt(p.indFirst)}pt`);
  if (p.borders) {
    const borders = cleanBorders(p.borders);
    const striped = stripedBorders(borders);
    // (Or none: a style's picture isn't to draw the paragraph's own borders.)
    if (striped) css.push(striped.css);
    else if (['top', 'bottom', 'left', 'right'].some((s) => borders[s])) css.push('border-image:none');
    for (const side of ['top', 'bottom', 'left', 'right']) {
      const b = borders[side];
      if (!b) continue;
      if (striped && drawnSide(b)) {
        // Drawn by the picture, at Word's width less what the browser can't draw (it draws whole pixels: the nearest,
        // or fewer when the space beside it can't give up the difference), the rest added to that space, so the text
        // is where Word's width puts it; a side's width and space both outside the indent, as below.
        const want = striped.width[side];
        const space = ((b.space || 0) * 4) / 3;
        const near = Math.max(1, Math.round(want));
        const px = (side === 'top' || side === 'bottom') && near - want > space ? Math.max(1, Math.floor(want)) : near;
        css.push(`border-${side}:${px * 0.75}pt solid ${striped.color[side]}`);
        if (side === 'top' || side === 'bottom') css.push(`padding-${side}:${+Math.max(0, space + want - px).toFixed(3)}px`);
        else {
          if (b.space) css.push(`padding-${side}:${b.space}pt`);
          css.push(`--odx-bd${side[0]}:${+(px * 0.75 + (b.space || 0)).toFixed(3)}pt`);
        }
        continue;
      }
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
