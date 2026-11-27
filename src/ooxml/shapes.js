// DrawingML shapes and groups (wps:wsp, wpg:wgp, pic:pic inside them) drawn as one SVG:
// preset and custom geometry, solid fills and outlines, pictures, and text boxes.
// Display only; the drawing's XML is saved back untouched.
import { esc, elementKids as els, childNamed as child, serializeXml } from './xml.js';
import { svgTextEffects } from './texteffects.js';
import { lineFactor, fontStack, themeFace } from './props.js';
import { geometryPaths, readGeometry, adjustOf } from './geometry.js';
import { PRESETS } from './presets.js';
import { EMU_PER_PX, TWIPS_PER_PX, emuToPx, twipsToPx } from './units.js';
import { fontMetrics } from '../fonts/wordmetrics.js';

// A font's ascent in Word's line (per em): its Windows ascent and whatever its hhea total is larger by, where its numbers
// are known (wordmetrics.js); else its stand-in's, the first font of the stack it is drawn in whose numbers are known (Word
// sets a missing font's text on its stand-in's ascent: hdrht-word c1, AvantGarde Bk BT 12pt in a header text box at
// Century Gothic's 16.1px, where 0.778 of its line had put it 1.5px high, as a sample's header on every page); else 0.778
// of its line, Calibri's.
const wordAscent = (font) => {
  const own = fontMetrics((font || '').toLowerCase())?.wordAscent;
  if (own != null) return own;
  for (const f of fontStack(font).split(',').slice(1)) {
    const m = fontMetrics(f.replace(/["']/g, '').trim().toLowerCase());
    if (m?.wordAscent != null) return m.wordAscent;
  }
  return lineFactor(font) * 0.778;
};

/**
 * Office's default theme colours, for a document without a theme part: the current Office theme's, as Word 365 draws
 * such a document's shapes and charts (checked: its accent 1 is 156082).
 */
export const OFFICE = { accent1: '156082', accent2: 'E97132', accent3: '196B24', accent4: '0F9ED5', accent5: 'A02B93', accent6: '4EA72E', dk1: '000000', lt1: 'FFFFFF', dk2: '0E2841', lt2: 'E8E8E8' };
const px = (v) => emuToPx(v || 0);
const desc = (el, name) => (el ? [...el.getElementsByTagName('*')].find((n) => n.localName === name) || null : null);

/** A colour element (srgbClr, schemeClr, ...) -> { hex, alpha } */
function color(el, colors) {
  if (!el) return null;
  let hex = null;
  if (el.localName === 'srgbClr') hex = el.getAttribute('val');
  else if (el.localName === 'schemeClr') {
    const key = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2', phClr: 'accent1' }[el.getAttribute('val')] || el.getAttribute('val');
    hex = colors?.[key] || OFFICE[key]; // a document without a theme part uses Office's colours
  } else if (el.localName === 'sysClr') hex = el.getAttribute('lastClr') || '000000';
  else if (el.localName === 'prstClr') hex = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '00FF00', blue: '0000FF' }[el.getAttribute('val')] || '000000';
  if (!hex) return null;
  let alpha = 1;
  for (const m of els(el)) {
    const v = Number(m.getAttribute('val')) / 100000;
    if (m.localName === 'alpha') alpha = v;
    else if (['shade', 'tint', 'lumMod', 'lumOff', 'hueOff', 'hueMod', 'satOff', 'satMod'].includes(m.localName)) hex = modify(hex, m.localName, v);
  }
  return { hex, alpha };
}

const toLin = (c) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
const toSrgb = (c) => (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055);
const hexOf = (rgb) => rgb.map((c) => Math.round(Math.min(1, Math.max(0, c)) * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
/**
 * A colour changed as DrawingML says: shade and tint in linear light (towards black, towards white: Word's accent 1
 * #156082 at shade 15% is #042433), lumMod and lumOff on its HSL luminance ("Accent 1, Darker 25%" is lumMod 75%),
 * hueOff/hueMod and satOff/satMod on its hue and saturation (SmartArt's colourful variations step the hue).
 */
export function modify(hex, kind, v) {
  const rgb = [16, 8, 0].map((s) => ((parseInt(hex, 16) >> s) & 255) / 255);
  if (kind === 'shade' || kind === 'tint') return hexOf(rgb.map((c) => toSrgb(kind === 'shade' ? toLin(c) * v : toLin(c) * v + 1 - v)));
  const [r, g, b] = rgb;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let l = (max + min) / 2;
  const d = max - min;
  let sat = d ? d / (1 - Math.abs(2 * l - 1)) : 0;
  let h = 0;
  if (d) h = max === r ? ((g - b) / d) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  if (kind === 'lumMod') l *= v;
  else if (kind === 'lumOff') l += v;
  else if (kind === 'satMod') sat *= v;
  else if (kind === 'satOff') sat += v;
  else if (kind === 'hueMod') h *= v;
  else if (kind === 'hueOff') h += (v * 100000) / 60000 / 60; // (60000ths of a degree, in sixths of the circle)
  h = ((h % 6) + 6) % 6;
  l = Math.min(1, Math.max(0, l));
  sat = Math.min(1, Math.max(0, sat));
  const c = (1 - Math.abs(2 * l - 1)) * sat;
  const x = c * (1 - Math.abs((h % 2 + 2) % 2 - 1));
  const m0 = l - c / 2;
  const hi = ((Math.floor(h) % 6) + 6) % 6;
  const [r1, g1, b1] = [[c, x, 0], [x, c, 0], [0, c, x], [0, x, c], [x, 0, c], [c, 0, x]][hi];
  return hexOf([r1 + m0, g1 + m0, b1 + m0]);
}

function mix(hex, keep, other) {
  const a = parseInt(hex, 16);
  const b = parseInt(other, 16);
  return [16, 8, 0].map((s) => Math.round(((a >> s) & 255) * keep + ((b >> s) & 255) * (1 - keep)).toString(16).padStart(2, '0')).join('');
}

/** fill="..." stroke="..." attributes from spPr (and the shape's style references as a fallback). */
/** A shape's fill and outline: { fill, stroke ({ hex, alpha } or null), width (px), ln (its a:ln) }. */
function paints(spPr, style, colors) {
  const solid = (el) => el && color(els(el)[0], colors);
  let fill = null;
  let stroke = null;
  // (A theme line style's width by its index: Office's are 0.5pt, 1pt and 1.5pt.)
  const lnIdx = Number(style && child(style, 'lnRef')?.getAttribute('idx')) || 0;
  let width = lnIdx ? px([6350, 12700, 19050][lnIdx - 1] || 12700) : 0.75 * (4 / 3);
  const f = els(spPr).find((n) => /Fill$/.test(n.localName));
  // (A gradient's first colour stands for it where one colour is wanted; its stops are drawn, gradientFill.)
  if (f) fill = f.localName === 'solidFill' ? solid(f) : f.localName === 'gradFill' ? gradientOf(f, colors) : null;
  else if (style && child(style, 'fillRef') && child(style, 'fillRef').getAttribute('idx') !== '0') fill = color(els(child(style, 'fillRef'))[0], colors);
  const ln = child(spPr, 'ln');
  if (ln) {
    if (ln.getAttribute('w')) width = px(ln.getAttribute('w'));
    stroke = child(ln, 'noFill') ? null : solid(child(ln, 'solidFill')) || (style && child(style, 'lnRef') ? color(els(child(style, 'lnRef'))[0], colors) : null);
  } else if (lnIdx) stroke = color(els(child(style, 'lnRef'))[0], colors);
  // Its corners: the outline's own join; else its theme line style's (Office's are mitred: a styled rectangle's corners
  // square in Word's print); else round, DrawingML's own (shapeclip-word: a text box with no style, its 8pt outline's
  // corners round in Word's print).
  const own = ['round', 'bevel', 'miter'].find((k) => child(ln, k));
  const join = own || (lnIdx ? 'miter' : 'round');
  return { fill, stroke, width, ln, join };
}

/** A gradient fill (a:gradFill) -> its first colour, with { grad: { stops, lin, path, rotWithShape } }, or null. */
function gradientOf(f, colors) {
  const stops = els(child(f, 'gsLst')).filter((g) => g.localName === 'gs')
    .map((g) => ({ pos: Math.min(1, Math.max(0, Number(g.getAttribute('pos')) / 100000 || 0)), c: color(els(g)[0], colors) }))
    .filter((s) => s.c).sort((a, b) => a.pos - b.pos);
  if (!stops.length) return null;
  const lin = child(f, 'lin');
  const path = child(f, 'path');
  const inset = (el) => Object.fromEntries(['l', 't', 'r', 'b'].map((k) => [k, Number(el?.getAttribute(k)) / 100000 || 0]));
  return {
    ...stops[0].c,
    grad: {
      stops,
      // (No a:lin or a:path: Word draws it top to bottom, as a:lin at 90 degrees.)
      lin: path ? null : { ang: Number(lin?.getAttribute('ang') || (lin ? 0 : 5400000)) / 60000, scaled: on(lin, 'scaled') },
      path: path ? { kind: path.getAttribute('path') || 'rect', focus: inset(child(path, 'fillToRect')) } : null,
      tile: inset(child(f, 'tileRect')),
      rotWithShape: f.getAttribute('rotWithShape') !== '0',
    },
  };
}

/**
 * A gradient's stops with stops between them, so a browser's straight sRGB blend comes out as Word's: Word blends in
 * linear light, easing in and out of each colour (gradients-word, red C02020 to white top to bottom: green 59 a tenth of
 * the way, 184 halfway, 251 nine tenths; a straight blend has 56, 145, 233). `ease` is that curve, near smoothstep.
 */
export function gradientStops(stops, steps = 8) {
  const out = [];
  const ease = (t) => t * t * (3 - 2 * t);
  for (let i = 0; i < stops.length; i++) {
    const a = stops[i], b = stops[i + 1];
    out.push(a);
    if (!b || b.pos - a.pos < 0.02) continue;
    const ca = [16, 8, 0].map((s) => toLin(((parseInt(a.c.hex, 16) >> s) & 255) / 255));
    const cb = [16, 8, 0].map((s) => toLin(((parseInt(b.c.hex, 16) >> s) & 255) / 255));
    for (let k = 1; k < steps; k++) {
      const t = k / steps, e = ease(t);
      out.push({ pos: a.pos + (b.pos - a.pos) * t, c: { hex: hexOf(ca.map((v, j) => toSrgb(v + (cb[j] - v) * e))), alpha: a.c.alpha + (b.c.alpha - a.c.alpha) * e } });
    }
  }
  return out;
}

/**
 * A gradient fill drawn on a path `d` of a w x h px shape turned `rot` degrees, as Word draws it (gradients-word, made by
 * Word: two-colour gradients of each kind): { defs, fill } (fill: the path's fill attribute), or { defs, under } (drawn
 * under the path, clipped to it) for a rectangle's or shape's path gradient. `mode` darkens or lightens it (modeFill).
 *  - Linear (a:lin): along its angle, clockwise from left to right; "scaled" works it out in a square stretched to the box
 *    (a 45-degree gradient on a 220 x 60 shape runs corner to corner, its colours level along the other diagonal).
 *  - Path (a:path): its first colour on the focus rectangle (a:fillToRect, insets of the shape's box: 50% each way is its
 *    centre), its last on the edge of the tile (a:tileRect, insets of the box, negative ones beyond it: Word's "from the
 *    corner" puts the focus on a corner and the tile a box's width past it, so the corner's colour reaches the far
 *    corner). A rectangle's ("rect", and "shape", drawn the same) as concentric rectangles, their corners on the lines from
 *    the focus to the tile's corners (the X Word's "from the centre" draws); a circle's as rings to the tile's corners.
 *  - Turned with the shape, unless rotWithShape="0" (SmartArt's): then it keeps its angle on the page.
 */
function gradientFill(fill, d, w, h, rot, ctx, mode) {
  const g = fill.grad;
  const id = `odxg${(ctx.gradN = (ctx.gradN || 0) + 1)}`;
  const r = (v) => +v.toFixed(4);
  const stops = gradientStops(g.stops.map((s) => ({ pos: s.pos, c: modeFill(s.c, mode) })))
    .map((s) => `<stop offset="${r(s.pos)}" stop-color="#${s.c.hex}"${s.c.alpha < 1 ? ` stop-opacity="${r(s.c.alpha)}"` : ''}/>`).join('');
  if (g.lin) {
    const a = ((g.lin.ang - (g.rotWithShape ? 0 : rot)) * Math.PI) / 180;
    const [cos, sin] = [Math.cos(a), Math.sin(a)];
    // Through the box's centre, as long as the box is across in that direction (in a unit square when scaled).
    const [bw, bh] = g.lin.scaled ? [1, 1] : [w, h];
    const half = (Math.abs(bw * cos) + Math.abs(bh * sin)) / 2;
    const [x1, y1, x2, y2] = [bw / 2 - cos * half, bh / 2 - sin * half, bw / 2 + cos * half, bh / 2 + sin * half];
    const tf = g.lin.scaled ? ` gradientTransform="matrix(${r(w)} 0 0 ${r(h)} 0 0)"` : '';
    return { defs: `<linearGradient id="${id}" gradientUnits="userSpaceOnUse" x1="${r(x1)}" y1="${r(y1)}" x2="${r(x2)}" y2="${r(y2)}"${tf}>${stops}</linearGradient>`, fill: `fill="url(#${id})"` };
  }
  // The focus rectangle and the tile, in the box's px.
  const { focus: fo } = g.path, ti = g.tile;
  let [fl, fr, ft, fb] = [fo.l * w, w - fo.r * w, fo.t * h, h - fo.b * h];
  if (fl > fr) fl = fr = (fl + fr) / 2;
  if (ft > fb) ft = fb = (ft + fb) / 2;
  const [tl, tr, tt, tb] = [ti.l * w, w - ti.r * w, ti.t * h, h - ti.b * h];
  if (g.path.kind === 'circle') {
    const cx = (fl + fr) / 2, cy = (ft + fb) / 2;
    // (Rings stretched to the tile's shape: worked out in a unit square, scaled to it.)
    const tw = tr - tl || 1, th = tb - tt || 1;
    const ux = (cx - tl) / tw, uy = (cy - tt) / th;
    const rad = Math.max(...[[0, 0], [1, 0], [0, 1], [1, 1]].map(([x, y]) => Math.hypot(x - ux, y - uy)));
    return { defs: `<radialGradient id="${id}" gradientUnits="userSpaceOnUse" cx="${r(ux)}" cy="${r(uy)}" r="${r(rad)}" gradientTransform="matrix(${r(tw)} 0 0 ${r(th)} ${r(tl)} ${r(tt)})">${stops}</radialGradient>`, fill: `fill="url(#${id})"` };
  }
  // Four sides (top, right, bottom, left), each from the focus rectangle's edge to the tile's, and the focus rectangle in
  // the first colour. (Each side drawn a hair wider, with its own gradient as its outline, so no seam shows between them.)
  const sides = [
    [[tl, tt], [tr, tt], [fr, ft], [fl, ft], [0, ft, 0, tt]],
    [[fr, ft], [tr, tt], [tr, tb], [fr, fb], [fr, 0, tr, 0]],
    [[fl, fb], [fr, fb], [tr, tb], [tl, tb], [0, fb, 0, tb]],
    [[tl, tt], [fl, ft], [fl, fb], [tl, tb], [fl, 0, tl, 0]],
  ];
  let defs = `<clipPath id="${id}c"><path d="${d}"/></clipPath>`;
  let under = '';
  sides.forEach((s, i) => {
    const [x1, y1, x2, y2] = s[4];
    if (x1 === x2 && y1 === y2) return;
    defs += `<linearGradient id="${id}s${i}" gradientUnits="userSpaceOnUse" x1="${r(x1)}" y1="${r(y1)}" x2="${r(x2)}" y2="${r(y2)}">${stops}</linearGradient>`;
    under += `<path d="M${s.slice(0, 4).map(([x, y]) => `${r(x)} ${r(y)}`).join(' L')} Z" fill="url(#${id}s${i})" stroke="url(#${id}s${i})" stroke-width="0.5"/>`;
  });
  if (fr > fl && fb > ft) under += `<rect x="${r(fl)}" y="${r(ft)}" width="${r(fr - fl)}" height="${r(fb - ft)}" ${paint('fill', modeFill(g.stops[0].c, mode))}/>`;
  return { defs, under: `<g clip-path="url(#${id}c)">${under}</g>` };
}

/** SVG paint attributes: `kind` (fill or stroke) in colour `c` (none without one). */
const paint = (kind, c) => (c ? `${kind}="#${c.hex}"${c.alpha < 1 ? ` ${kind}-opacity="${+c.alpha.toFixed(3)}"` : ''}` : `${kind}="none"`);
/** A path's fill as its fill mode makes it (a can's lighter top, a cube's darker side). */
const FILL_MODE = { darken: [0.6, '000000'], darkenLess: [0.8, '000000'], lighten: [0.6, 'FFFFFF'], lightenLess: [0.8, 'FFFFFF'] };
function modeFill(c, mode) {
  if (!c || mode === 'none') return null;
  const m = FILL_MODE[mode];
  return m ? { ...c, hex: mix(c.hex, m[0], m[1]) } : c;
}

/** A shape's paths ([{ d, fill, stroke, ends }], geometry.js) in a w x h px box: its preset's or its own (a:custGeom). */
function geometry(spPr, w, h) {
  const cust = child(spPr, 'custGeom');
  const prst = child(spPr, 'prstGeom');
  const def = cust ? readGeometry(cust) : PRESETS[prst?.getAttribute('prst')] || PRESETS.rect;
  // (In EMU, as the definitions' numbers are, drawn in px.)
  return geometryPaths(def, w * EMU_PER_PX, h * EMU_PER_PX, adjustOf(cust || prst), 1 / EMU_PER_PX);
}

/** A line end's arrowhead (a:headEnd / a:tailEnd) at [x, y] pointing along [dx, dy], for a line `width` px wide. */
function arrowhead(end, at, width, c) {
  const type = end?.getAttribute('type');
  if (!type || type === 'none' || !at) return '';
  const k = { sm: 2, med: 3, lg: 5 };
  // (Sized as for a line at least 2pt wide: Word's medium arrowhead on a 1pt line is 6pt long, measured.)
  const base = Math.max(width, 8 / 3);
  const hw = (base * (k[end.getAttribute('w')] || 3)) / 2; // half its width
  const len = base * (k[end.getAttribute('len')] || 3);
  const [x, y, dx, dy] = at;
  const nx = -dy, ny = dx; // across the line
  const p = (a, b) => `${(x + dx * a + nx * b).toFixed(2)} ${(y + dy * a + ny * b).toFixed(2)}`;
  const col = `#${c.hex}`;
  if (type === 'oval') return `<ellipse cx="${(x - dx * len / 2).toFixed(2)}" cy="${(y - dy * len / 2).toFixed(2)}" rx="${(len / 2).toFixed(2)}" ry="${hw.toFixed(2)}" transform="rotate(${(Math.atan2(dy, dx) * 180 / Math.PI).toFixed(2)} ${(x - dx * len / 2).toFixed(2)} ${(y - dy * len / 2).toFixed(2)})" fill="${col}"/>`;
  if (type === 'diamond') return `<path d="M${p(0, 0)}L${p(-len / 2, hw)}L${p(-len, 0)}L${p(-len / 2, -hw)}Z" fill="${col}"/>`;
  if (type === 'arrow') return `<path d="M${p(-len, hw)}L${p(0, 0)}L${p(-len, -hw)}" fill="none" stroke="${col}" stroke-width="${width.toFixed(2)}" stroke-linejoin="miter"/>`;
  if (type === 'stealth') return `<path d="M${p(0, 0)}L${p(-len, hw)}L${p(-len * 0.6, 0)}L${p(-len, -hw)}Z" fill="${col}"/>`;
  return `<path d="M${p(0, 0)}L${p(-len, hw)}L${p(-len, -hw)}Z" fill="${col}"/>`; // triangle
}

// A w: attribute (w:val, w:left, ...).
const wa = (el, name) => (el ? el.getAttribute(`w:${name}`) : null);

let canvas = null;
let missed = null; // canvas font -> text measured in a fallback because the page hadn't loaded that face yet

/**
 * The faces text was measured in a fallback for since the last call (canvas font -> some of its text), or null: a
 * picture drawn then wraps its lines at the fallback's widths, and is drawn again once they've loaded (svgfonts.js
 * redrawLater). (smartart.docx in Aptos: its boxes broke their lines in the fallback until drawn again.)
 */
export function unmeasured() {
  const out = missed;
  missed = null;
  return out;
}

/** Whether text is measured (in a browser), not estimated: WordArt's box is fitted to it only then (shapeSvg). */
function measures() {
  if (canvas == null) canvas = typeof document !== 'undefined' && document.createElement('canvas').getContext?.('2d') || false;
  return !!canvas;
}

/** Small caps' text in a tspan: its lowercase letters as capitals in a smaller tspan of their own (`sc`, px). */
function smallCapsText(text, sc) {
  return (text.match(/\p{Ll}+|\P{Ll}+/gu) || []).map((part) => (/\p{Ll}/u.test(part) ? `<tspan font-size="${sc.toFixed(2)}">${esc(part.toUpperCase())}</tspan>` : esc(part))).join('');
}

/** Width of `text` in px: measured in a browser, else estimated at half an em per character. */
function textWidth(text, sizePx, bold, font, italic = false) {
  if (canvas == null) canvas = typeof document !== 'undefined' && document.createElement('canvas').getContext?.('2d') || false;
  if (!canvas) return text.length * sizePx * 0.5;
  canvas.fontKerning = 'none';
  // (In its italic face where it is italic: Times New Roman's italic letters are narrower than its upright ones, and a
  // text box of italic text, measured upright, broke each line a word or two early.)
  // (At 100px and scaled: the font's design advances, which Word sets text by, not the browser's hinted ones at a small
  // size: bold Arial at 10pt came out 0.5% narrow, and a box's line kept a word Word wraps, widths.js designWidth's
  // reason in the body; ..\charscale-word w07, sample 1113 page 60.)
  const at = 100;
  canvas.font = `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${at}px ${fontStack(font)}`;
  // (A stand-in not loaded yet is asked for, and noted: this drawing measures in what the page has.)
  try {
    if (document.fonts && text.trim() && !document.fonts.check(canvas.font, text)) {
      document.fonts.load(canvas.font, text).catch(() => {});
      missed ||= new Map();
      if ((missed.get(canvas.font) || '').length < 200) missed.set(canvas.font, (missed.get(canvas.font) || '') + text);
    }
  } catch { /* not a font */ }
  return (canvas.measureText(text).width * sizePx) / at;
}

/**
 * Text box paragraphs as SVG text: indents, tab stops, spacing and exact line heights from the
 * paragraph, size, colour and weight per run. Lines wrap at the box's inner width (unless the box
 * says wrap="none"), word by word as in Word; a tab that would pass the right edge starts a new line.
 */
function textBox(txbx, bodyPr, w, h, ctx, line = 0, fit = false) {
  const font0 = ctx.font;
  // The insets are from the inside of the box's outline, which is centred on its edge: half the line in
  // (shapeclip-word: under an 8pt outline, the text 4pt in from where the insets alone put it, in Word's print, across
  // and down; tbox-word, under 0.75pt and 8pt ones, its lines where Word prints them; the whole line had put it 4pt further).
  const ins = (k, d) => (bodyPr && bodyPr.getAttribute(k) != null ? px(bodyPr.getAttribute(k)) : d) + line / 2;
  const l = ins('lIns', 9.6);
  const t = ins('tIns', 4.8);
  const r = ins('rIns', 9.6);
  // (`fit`: wrapped at the box's width all the same: WordArt's box as wide as the column, shapeSvg.)
  const wrap = fit || bodyPr?.getAttribute('wrap') !== 'none';
  // Not wrapped and resized to fit (WordArt's box): Word makes the box as wide as its text, from its left edge, so
  // centred or right-aligned lines line up against the widest line, not the box's saved width (Word's PDF of its
  // gallery's styles: each at the box's left, in a box saved 2in wide).
  const shrink = !wrap && !!child(bodyPr, 'spAutoFit');
  const out = [];
  let y = t;
  let inkRight = 0; // (where the widest line's last letter ends, for a box as wide as its text: shapeSvg)
  // One paragraph between x = L and R (`tbl`: the table style round it, whose paragraph settings beat Normal's).
  const layPara = (p, L, R, tbl) => {
    const l = L; // (tab stops count from the paragraph's box)
    const pPr = child(p, 'pPr');
    // The paragraph's style (and the document defaults) under its own formatting.
    const sid = wa(child(pPr, 'pStyle'), 'val');
    const st0 = ctx.para ? ctx.para(sid) : { pPr: {}, rPr: {}, font: font0 };
    // (In a table, its style's paragraph settings go over Normal's and the document defaults, under any other style's own.)
    const normal = !sid || sid === ctx.normal;
    const st = tbl ? { ...st0, pPr: normal ? { ...st0.pPr, ...tbl.pPr } : { ...st0.pPr, ...tbl.pPr, ...(st0.own || {}) }, rPr: normal ? { ...st0.rPr, ...tbl.rPr } : st0.rPr } : st0;
    const markSz = Number(wa(desc(child(pPr, 'rPr'), 'sz'), 'val')) / 2 || (st.rPr.size ? st.rPr.size / 2 : 11);
    const spacing = child(pPr, 'spacing');
    const sp = (k) => (wa(spacing, k) != null ? Number(wa(spacing, k)) : st.pPr[k]);
    const lineRule = wa(spacing, 'lineRule') ?? (wa(spacing, 'line') != null ? null : st.pPr.lineRule);
    const ind = child(pPr, 'ind');
    const align = wa(child(pPr, 'jc'), 'val') ?? st.pPr.align;
    const tabs = els(child(pPr, 'tabs')).map((tab) => twipsToPx(wa(tab, 'pos'))).filter((v) => v > 0).sort((x, z) => x - z);
    // A list paragraph (w:numPr, its own or its style's): its level's indents under the paragraph's own, and its label
    // (ctx.listLabel, counted in the box's order) at the first line's start, then a tab or a space. (Without them a
    // sample's lettered lists in its form's boxes lost their letters, and a hanging indent with no left indent of its own
    // put each item's first letter outside the box, the rest of its first word inside.)
    const numPr = child(pPr, 'numPr');
    const numId = numPr ? wa(child(numPr, 'numId'), 'val') : st.pPr.numId;
    const ilvl = Number(numPr ? wa(child(numPr, 'ilvl'), 'val') ?? 0 : st.pPr.ilvl ?? 0) || 0;
    const lvl = numId != null && numId !== '0' && ctx.listLevel ? ctx.listLevel(numId, ilvl) : null;
    const label = lvl && !lvl.hidden && ctx.listLabel ? ctx.listLabel(numId, ilvl) : null;
    const lp = lvl?.pPr || {};
    const indL = wa(ind, 'left') ?? wa(ind, 'start') ?? lp.indLeft ?? st.pPr.indLeft ?? 0;
    const indR = wa(ind, 'right') ?? wa(ind, 'end') ?? lp.indRight ?? st.pPr.indRight ?? 0;
    const indF = wa(ind, 'firstLine') != null || wa(ind, 'hanging') != null ? (Number(wa(ind, 'firstLine')) || 0) - (Number(wa(ind, 'hanging')) || 0) : lp.indFirst ?? st.pPr.indFirst ?? 0;
    const left = l + twipsToPx(indL);
    const right = R - twipsToPx(indR);
    const first = left + twipsToPx(indF);
    // Lines of placed pieces: { x, ink (where its last visible glyph ends), text, open (its tspan's start tag) }.
    const lines = [{ pieces: [], x: first }];
    const newLine = () => lines.push({ pieces: [], x: left });
    let maxSize = 0;
    let lineF = lineFactor(st.font);
    if (label) {
      // The label in its level's font, size, colour and weight, else the paragraph mark's.
      const mark = child(pPr, 'rPr');
      const size = lvl.size ? lvl.size / 2 : markSz;
      const sizePx = (size * 4) / 3;
      const font = lvl.font || wa(child(mark, 'rFonts'), 'ascii') || st.font || font0;
      const bold = lvl.bold ?? !!st.rPr.bold;
      const italic = lvl.italic ?? !!st.rPr.italic;
      const c = lvl.color ?? wa(child(mark, 'color'), 'val') ?? st.rPr.color;
      const open = `<tspan font-family="${esc(fontStack(font))}" font-size="${sizePx.toFixed(2)}" fill="#${c && c !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(c) ? c : ctx.textColor || '000000'}"${bold ? ' font-weight="bold"' : ''}${italic ? ' font-style="italic"' : ''}>`;
      const line = lines[0];
      const w = textWidth(label, sizePx, bold, font, italic);
      line.pieces.push({ x: line.x, text: label, open, size: sizePx, ink: line.x + w, end: line.x + w });
      line.one = sizePx * lineFactor(font);
      line.asc = sizePx * wordAscent(font);
      maxSize = sizePx;
      line.x += w;
      if (lvl.suff === 'space') line.x += textWidth(' ', sizePx, bold, font, italic);
      else if (lvl.suff !== 'nothing') {
        // To the next tab stop past it: the paragraph's own, else the hanging indent's, else the next half inch.
        const stop = tabs.find((v) => l + v > line.x + 0.5);
        line.x = stop != null ? l + stop : left > line.x + 0.5 ? left : l + (Math.floor((line.x - l) / 48) + 1) * 48;
        line.tabbed = true;
      }
    }
    // A field in the box's text (w:fldChar): its code as it is read, and, once its result starts, the page-number
    // field it is (PAGE, NUMPAGES, SECTIONPAGES, with a number format switch or none), which marks the result's
    // text in the drawing so a header or footer can show each page's own number (pagefields.js).
    let fieldCode = null, fieldMark = '';
    for (const run of [...p.getElementsByTagName('*')].filter((n) => n.localName === 'r')) {
      const rPr = child(run, 'rPr');
      // Its character style (w:rStyle) between the paragraph style and its own formatting.
      const chr = ctx.char ? ctx.char(wa(child(rPr, 'rStyle'), 'val')) || {} : {};
      const cs = { ...st.rPr, ...chr };
      const size = Number(wa(child(rPr, 'sz'), 'val')) / 2 || (cs.size ? cs.size / 2 : markSz);
      const sizePx = (size * 4) / 3;
      const c = wa(child(rPr, 'color'), 'val') ?? cs.color;
      const on = (name, key) => { const e = child(rPr, name); return e ? !/^(0|false)$/.test(wa(e, 'val') || '') : !!cs[key]; };
      const bold = on('b', 'bold');
      const italic = on('i', 'italic');
      const letter = (Number(wa(child(rPr, 'spacing'), 'val')) || 0) / TWIPS_PER_PX;
      const rf = child(rPr, 'rFonts');
      const theme = wa(rf, 'asciiTheme');
      const ofTheme = (k) => k && themeFace(ctx.theme, k);
      // (A character style's font, named or the theme's (w:asciiTheme), over the paragraph style's and the document
      // defaults': a cover banner's "cover:title" style in Calibri (minorHAnsi) over Times New Roman defaults, Word's
      // print in Calibri, had come out in Times.)
      // (The run's theme font over a named one beside it, w:asciiTheme over w:ascii, as everywhere: props.js readRPr.)
      const font = ofTheme(theme) || wa(rf, 'ascii') || wa(rf, 'hAnsi') || chr.font || ofTheme(chr.themeFont) || st.font || font0;
      // Word 2010's text effects (WordArt's): fill, outline, shadow, glow (texteffects.js).
      const fx = svgTextEffects(rPr, (ctx.fxN = (ctx.fxN || 0) + 1), serializeXml);
      // (An outline or shadow alone keeps the run's own colour: Word draws the letters in w:color.)
      const own = `fill="#${c && c !== 'auto' ? c : ctx.textColor || '000000'}"`;
      const paint = fx?.paint ? (/\bfill=/.test(fx.paint) ? fx.paint : `${own} ${fx.paint}`) : own;
      for (const n of els(run)) {
        if (n.localName === 'fldChar') {
          const t = wa(n, 'fldCharType');
          if (t === 'begin') { fieldCode = ''; fieldMark = ''; }
          else if (t === 'separate' && fieldCode != null) {
            const m = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)(?:\s+\\\*\s*([A-Za-z]{1,12}))?/.exec(fieldCode);
            fieldMark = m ? ` data-odx-fld="${m[1]}${m[2] ? ` ${m[2]}` : ''}"` : '';
            fieldCode = null;
          } else if (t === 'end') { fieldCode = null; fieldMark = ''; }
        } else if (n.localName === 'instrText' && fieldCode != null && fieldCode.length < 200) fieldCode += n.textContent;
      }
      // Small caps (w:smallCaps, not under w:caps): lowercase letters as capitals at 80% of the size, to the half point,
      // as Word sets them (widths.js; a cover banner's subtitle, Word's print in small capitals).
      const sc = on('smallCaps', 'smallCaps') && !on('caps', 'caps') ? (Math.round(size * 0.8 * 2) / 2) * 4 / 3 : 0;
      const tw = (s) => (sc ? (s.match(/\p{Ll}+|\P{Ll}+/gu) || []).reduce((t, part) => t + (/\p{Ll}/u.test(part) ? textWidth(part.toUpperCase(), sc, bold, font, italic) : textWidth(part, sizePx, bold, font, italic)), 0) : textWidth(s, sizePx, bold, font, italic));
      const open = `<tspan font-family="${esc(fontStack(font))}" font-size="${sizePx.toFixed(2)}" ${paint}${bold ? ' font-weight="bold"' : ''}${italic ? ' font-style="italic"' : ''}${letter ? ` letter-spacing="${letter.toFixed(2)}"` : ''}${sc ? ` data-odx-sc="${sc.toFixed(2)}"` : ''}${fieldMark}>`;
      if (fx?.defs) out.push(`<defs>${fx.defs}</defs>`);
      for (const n of els(run)) {
        const line = lines[lines.length - 1];
        if (n.localName === 'tab') {
          const stop = tabs.find((v) => v > line.x - l + 0.5);
          const to = l + (stop ?? (Math.floor((line.x - l) / 48) + 1) * 48);
          if (wrap && to >= right && line.pieces.length) newLine();
          else { line.x = to; line.tabbed = true; }
        } else if (n.localName === 'br' || n.localName === 'cr') {
          newLine();
        } else if (n.localName === 't' && n.textContent) {
          if (sizePx * lineFactor(font) > maxSize * lineF) lineF = lineFactor(font);
          maxSize = Math.max(maxSize, sizePx);
          // Words keep their trailing spaces; a word that doesn't fit starts the next line (one wider than a whole line in
          // a box fitted to the column, WordArt's, broken letter by letter, as Word breaks it).
          const words = [];
          // (All caps, w:caps, from the run or its styles: drawn in capitals, as Word does; a header banner's bound title
          // had kept its lower case.)
          const text = on('caps', 'caps') ? n.textContent.toUpperCase() : n.textContent;
          for (const word of text.match(/\S*\s*/g).filter(Boolean)) {
            if (!fit || tw(word.trimEnd()) + word.trimEnd().length * letter <= right - left + 0.5) { words.push(word); continue; }
            let part = '', partW = 0;
            for (const ch of word) {
              // (The piece measured whole while short, so kerning counts; then by adding each letter, so a long word
              // costs its length, not its square.)
              let w = part.length < 64 ? tw(part + ch) : partW + tw(ch);
              if (part && /\S/.test(ch) && w + (part.length + 1) * letter > right - left + 0.5) { words.push(part); part = ''; w = tw(ch); }
              part += ch;
              partW = w;
            }
            if (part) words.push(part);
          }
          for (const word of words) {
            let cur = lines[lines.length - 1];
            const width = tw(word) + word.length * letter;
            const inked = tw(word.trimEnd()) + word.trimEnd().length * letter;
            if (wrap && cur.x + inked > right + 0.5 && cur.pieces.length) { cur.wrapped = true; newLine(); cur = lines[lines.length - 1]; }
            const last = cur.pieces[cur.pieces.length - 1];
            cur.one = Math.max(cur.one || 0, sizePx * lineFactor(font)); // (each line as tall as its own text: Word's)
            cur.asc = Math.max(cur.asc || 0, sizePx * wordAscent(font));
            if (last && last.open === open && last.end === cur.x) last.text += word;
            else cur.pieces.push({ x: cur.x, text: word, open, sc, filter: fx?.filter, refl: fx?.reflection, size: sizePx, n: ctx.fxN });
            cur.pieces[cur.pieces.length - 1].ink = cur.x + inked;
            cur.x += width;
            cur.pieces[cur.pieces.length - 1].end = cur.x;
          }
        }
      }
    }
    const sizePx = maxSize || (markSz * 4) / 3;
    const line = Number(sp('line')) || 0;
    const rule = lineRule;
    // Each font's own single line (Calibri 1.22em, Times New Roman 1.15em), as in the body, each line its own text's.
    const heightOf = (o) => (line && rule === 'exact' ? line / TWIPS_PER_PX : line && rule === 'atLeast' ? Math.max(line / TWIPS_PER_PX, o) : o * (line ? line / 240 : 1));
    // The space between paragraphs is the larger of the one's after and the next's before, as Word sets a text box's
    // (none between two of the same style that say so, w:contextualSpacing).
    // (Automatic spacing, w:beforeAutospacing on the paragraph or its style, the Normal (Web) style's: 14pt whatever number
    // comes with it, and none before a box's first paragraph. Word's prints of 12 generated documents: automatic, a
    // plain 14pt and 5pt and none, on the box's first paragraph and after a plain one.)
    const autoOf = (k) => { const v = wa(spacing, `${k}Autospacing`); return v != null ? /^(1|true|on)$/.test(v) : !!st.pPr[`${k}Auto`]; };
    const before = autoOf('before') ? (prevStyle === null ? 0 : 280 / TWIPS_PER_PX) : (Number(sp('before')) || 0) / TWIPS_PER_PX;
    const ctx2 = wa(child(pPr, 'contextualSpacing'), 'val');
    const contextual = child(pPr, 'contextualSpacing') ? !/^(0|false|off)$/.test(ctx2 || '') : !!st.pPr.contextualSpacing;
    if (contextual && prevStyle === (sid || '')) { y -= prevAfter; }
    else y += Math.max(0, before - prevAfter);
    for (const q of lines) if (q.pieces.length) inkRight = Math.max(inkRight, q.pieces[q.pieces.length - 1].ink);
    const widest = Math.max(0, ...lines.filter((q) => q.pieces.length).map((q) => q.pieces[q.pieces.length - 1].ink - (q === lines[0] ? first : left)));
    for (const ln of lines) {
      const one = ln.one || sizePx * lineF;
      const lineH = heightOf(one);
      // Baseline: the font's ascent (0.778 of its line: 0.95em of Calibri's 1.22) below the line top, centred in an exact
      // line's extra room; a multiple line's extra goes below its text, as in the body (Word's PDF of a text box at
      // 1.08 lines: its first baseline 0.7px higher than centred).
      // (Not exact: the font's own ascent in Word's line below its top, wordAscent: Word's prints of tboxbase-word's boxes,
      // a line of Arial, Times New Roman or Calibri at 7 to 14pt, its baseline 0.938, 0.934 and 0.952em down; 0.778 of
      // the line put Arial's and Times New Roman's 0.3 to 1.1px high, 1113's form footers among them.)
      // (Exact: at 80% of the line's height, whatever the font and size: Word's prints of tboxbase-word b04, Arial 10pt
      // and Times New Roman 11pt on lines of 150 to 320tw, each baseline 0.8 of its line down to 0.1px; centred, a line
      // shorter than the font's was 0.5 to 1.3px low, a taller one up to 2.3px high: sample 1113 page 60.)
      const base = y + (rule === 'exact' ? lineH * 0.8 : (rule === 'atLeast' ? lineH - one : 0) + (ln.asc || one * 0.778));
      y += lineH;
      if (!ln.pieces.length) continue;
      // Centred and right-aligned lines move as a whole (trailing spaces don't count).
      const start = ln === lines[0] ? first : left;
      const ink = ln.pieces[ln.pieces.length - 1].ink;
      const edge = shrink ? Math.min(right, start + widest) : right;
      const shift = align === 'center' ? (edge - ink) / 2 - (ln.pieces[0].x - start) / 2 : align === 'right' || align === 'end' ? edge - ink : 0;
      // A justified paragraph's lines that the next word ended (not its last, nor one a line break ends, nor one with a
      // tab): the room left to the right edge shared among the line's spaces, as Word sets them (a sample's text box
      // of justified text, its lines all ending on the box's inner edge in Word's print).
      const spacesIn = (pc, last) => ((last ? pc.text.trimEnd() : pc.text).match(/ /g) || []).length;
      const gaps = (align === 'both' || align === 'distribute') && ln.wrapped && !ln.tabbed && wrap ? ln.pieces.reduce((n, pc, i) => n + spacesIn(pc, i === ln.pieces.length - 1), 0) : 0;
      const each = gaps && right - ink > 0.05 ? (right - ink) / gaps : 0;
      let before = 0; // the line's spaces before this piece
      for (const [i, pc] of ln.pieces.entries()) {
        const widen = each ? ` style="word-spacing:${each.toFixed(3)}px"` : '';
        const text = `<text x="${(pc.x + shift + before * each).toFixed(1)}" y="${base.toFixed(1)}"${pc.filter ? ` filter="${pc.filter}"` : ''}${widen} xml:space="preserve">${pc.open}${pc.sc ? smallCapsText(pc.text, pc.sc) : esc(pc.text)}</tspan></text>`;
        before += spacesIn(pc, i === ln.pieces.length - 1);
        out.push(text);
        // A reflection (w14:reflection): the text mirrored about its baseline (where Word starts it), that far
        // below it, fading from its start opacity to its end one over its end position's share of the text's height,
        // softened as Word's (its print: a fifth of the blur, texteffects.js). Its descenders, mirrored, rise above
        // its top (Word's print of wordart-leftovers/refl.docx), so the mask starts a quarter of the size higher.
        if (pc.refl) {
          const { stA, endA, endPos, dist, blur } = pc.refl;
          const h = pc.size * 1.2, bottom = base + pc.size * 0.02, top = bottom + dist;
          const up = pc.size * 0.25, len = h * endPos + up;
          const id = `odx-rf${pc.n}-${out.length}`;
          const soft = blur > 0 ? `<filter id="${id}b" x="-10%" y="-50%" width="120%" height="200%"><feGaussianBlur stdDeviation="${+(blur / 5).toFixed(2)}"/></filter>` : '';
          out.push(`<defs><linearGradient id="${id}g" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#fff" stop-opacity="${stA}"/><stop offset="${+(up / len).toFixed(3)}" stop-color="#fff" stop-opacity="${stA}"/><stop offset="1" stop-color="#fff" stop-opacity="${endA}"/></linearGradient>`
            + `<mask id="${id}" maskUnits="userSpaceOnUse" x="-10000" y="${(top - up).toFixed(1)}" width="20000" height="${len.toFixed(1)}"><rect x="-10000" y="${(top - up).toFixed(1)}" width="20000" height="${len.toFixed(1)}" fill="url(#${id}g)"/></mask>${soft}</defs>`
            + `<g mask="url(#${id})"><g${soft ? ` filter="url(#${id}b)"` : ''}><g transform="translate(0 ${(bottom + top).toFixed(1)}) scale(1 -1)">${text}</g></g></g>`);
        }
      }
    }
    const after = autoOf('after') ? 280 / TWIPS_PER_PX : (Number(sp('after')) || 0) / TWIPS_PER_PX;
    y += after;
    prevAfter = after;
    prevStyle = sid || '';
  };
  let prevAfter = 0, prevStyle = null;
  // A table: its grid's columns from L, each row as tall as its tallest cell, the cells' paragraphs (and tables) laid
  // out in them with the cell margins, and the borders the table or its style gives drawn as lines.
  const layTable = (tbl, L, R) => {
    const tblPr = child(tbl, 'tblPr');
    const ts = ctx.tableStyle ? ctx.tableStyle(wa(child(tblPr, 'tblStyle'), 'val')) : null;
    const grid = els(child(tbl, 'tblGrid')).filter((n) => n.localName === 'gridCol').map((g) => twipsToPx(wa(g, 'w') || 0));
    if (!grid.length) return;
    const xs = [L];
    for (const g of grid) xs.push(xs[xs.length - 1] + g);
    const marEl = child(tblPr, 'tblCellMar');
    const mar = (side, d) => { const e = child(marEl, side) || child(marEl, side === 'left' ? 'start' : side === 'right' ? 'end' : side); return e ? twipsToPx(wa(e, 'w') || 0) : ts?.cellMar?.[side] ?? d; };
    const mL = mar('left', 7.2), mR = mar('right', 7.2), mT = mar('top', 0), mB = mar('bottom', 0);
    const bd = child(tblPr, 'tblBorders');
    const lined = (side) => { const e = child(bd, side); if (e) return !/^(nil|none)$/.test(wa(e, 'val') || ''); const t = ts?.borders?.[side]; return !!t && !/^(nil|none)$/.test(t.val || ''); };
    const sides = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].filter(lined);
    const rows = els(tbl).filter((n) => n.localName === 'tr');
    // Each row's cells with their grid columns, and how far down each cell merged down (w:vMerge) reaches.
    const cellsOf = rows.map((tr) => { let c = 0; return els(tr).filter((n) => n.localName === 'tc').map((tc) => { const tcPr = child(tc, 'tcPr'); const span = Number(wa(child(tcPr, 'gridSpan'), 'val')) || 1; const vm = child(tcPr, 'vMerge'); const cell = { tc, c, span, merge: vm ? (wa(vm, 'val') === 'restart' ? 'restart' : 'cont') : null }; c += span; return cell; }); });
    const last = (j, c) => { let k = j; while (k + 1 < rows.length && cellsOf[k + 1].some((x) => x.c === c && x.merge === 'cont')) k++; return k; };
    const top0 = y;
    const rowTops = [];
    const owed = new Map(); // row index -> the lowest a cell merged down to it from above reaches
    // A line's width (px) from a border element (w:sz eighths of a point), 0 for none; a cell's own side, else the table's.
    const lineW = (e) => (e && !/^(nil|none)$/.test(wa(e, 'val') || '') ? (Math.max(2, Number(wa(e, 'sz')) || 4) / 8) * (4 / 3) : 0);
    const cellLine = (tc, side, tableSide) => { const own = child(child(child(tc, 'tcPr'), 'tcBorders'), side); return own ? lineW(own) : lineW(child(bd, tableSide)); };
    const cellLines = []; // [x0, y0, x1, y1, width] of the cells' own lines, drawn after
    rows.forEach((tr, j) => {
      const top = y;
      rowTops.push(top);
      // The row's minimum or exact height (w:trHeight) and the widest line over its cells, which Word puts in the row
      // over that minimum (as in the body: styles.js, editor.css), its cells' text under it. (Ignored, a sample's codes
      // table in a text box came out 5 to 17px short, row by row: 1113 page 60.)
      const trH = child(child(tr, 'trPr'), 'trHeight');
      const minH = trH ? twipsToPx(Number(wa(trH, 'val')) || 0) : 0;
      const exact = wa(trH, 'hRule') === 'exact';
      const tl = Math.max(0, ...cellsOf[j].map(({ tc }) => cellLine(tc, 'top', j === 0 ? 'top' : 'insideH')));
      let bottom = top + tl + (exact ? 0 : minH);
      for (const { tc, c, span, merge } of cellsOf[j]) {
        const x0 = xs[Math.min(c, grid.length)], x1 = xs[Math.min(c + span, grid.length)];
        for (const [side, tside] of [['top', j === 0 ? 'top' : 'insideH'], ['left', c === 0 ? 'left' : 'insideV'], ['right', c + span >= grid.length ? 'right' : 'insideV']]) {
          if (!child(child(child(tc, 'tcPr'), 'tcBorders'), side)) continue; // (the table's own lines are drawn below)
          const w = cellLine(tc, side, tside);
          if (w) cellLines.push({ side, x0, x1, j, w });
        }
        if (merge === 'cont') continue; // (its text is the cell above's)
        y = top + tl + mT;
        prevAfter = 0; prevStyle = null;
        layBlocks(tc, x0 + mL, x1 - mR, ts || {});
        const k = merge === 'restart' ? last(j, c) : j;
        if (k > j) owed.set(k, Math.max(owed.get(k) ?? 0, y + mB));
        else bottom = Math.max(bottom, y + mB);
      }
      y = exact && minH ? top + minH : Math.max(bottom, owed.get(j) ?? 0);
    });
    rowTops.push(y);
    for (const { side, x0, x1, j, w } of cellLines) {
      const t = rowTops[j], b = rowTops[j + 1];
      const d = side === 'top' ? `M${x0.toFixed(1)} ${(t + w / 2).toFixed(2)}L${x1.toFixed(1)} ${(t + w / 2).toFixed(2)}` : side === 'left' ? `M${x0.toFixed(1)} ${t.toFixed(1)}L${x0.toFixed(1)} ${b.toFixed(1)}` : `M${x1.toFixed(1)} ${t.toFixed(1)}L${x1.toFixed(1)} ${b.toFixed(1)}`;
      out.push(`<path d="${d}" stroke="#000" stroke-width="${w.toFixed(2)}" fill="none"/>`);
    }
    rowTops.pop();
    prevAfter = 0; prevStyle = null;
    const x0 = xs[0], x1 = xs[xs.length - 1];
    const ln = (a, b, c2, d) => out.push(`<path d="M${a.toFixed(1)} ${b.toFixed(1)}L${c2.toFixed(1)} ${d.toFixed(1)}" stroke="#000" stroke-width="0.67" fill="none"/>`);
    if (sides.includes('top')) ln(x0, top0, x1, top0);
    if (sides.includes('bottom')) ln(x0, y, x1, y);
    if (sides.includes('left')) ln(x0, top0, x0, y);
    if (sides.includes('right')) ln(x1, top0, x1, y);
    if (sides.includes('insideH')) for (const t of rowTops.slice(1)) ln(x0, t, x1, t);
    if (sides.includes('insideV')) for (const x of xs.slice(1, -1)) ln(x, top0, x, y);
  };
  // A run of blocks: paragraphs, tables, and the paragraphs of block content controls (a cover page's title and author
  // boxes hold theirs in w:sdt).
  function layBlocks(el, L, R, tbl) {
    for (const n of els(el)) {
      if (n.localName === 'p') layPara(n, L, R, tbl);
      else if (n.localName === 'tbl') layTable(n, L, R);
      else if (n.localName === 'sdt') layBlocks(child(n, 'sdtContent'), L, R, tbl);
    }
  }
  layBlocks(desc(txbx, 'txbxContent'), l, w - r, null);
  ctx.textRight = inkRight + r;
  ctx.textEnd = y + ins('bIns', 4.8); // (where the text ends, for a box that grows to fit it: shapeSvg)
  // The text's place top to bottom (wps:bodyPr anchor): Word's own shapes centre theirs ("ctr"; groupturn-word's
  // "Text" in the middle of each shape). Only text shorter than its room moves: more stays at the top, where a box that
  // grows to fit it has it.
  const anchor = bodyPr?.getAttribute('anchor');
  const room = h - ins('bIns', 4.8) - y;
  const dy = room > 0 ? (anchor === 'ctr' ? room / 2 : anchor === 'b' ? room : 0) : 0;
  return dy ? `<g transform="translate(0 ${+dy.toFixed(2)})">${out.join('')}</g>` : out.join('');
}

/**
 * A DrawingML text body (a:txBody: SmartArt's shapes) as SVG text in the box { x, y, w, h } (px): its insets, vertical
 * anchor, paragraphs' alignment, margins, bullets, line and paragraph spacing, and runs' size, weight, colour and font.
 * Lines wrap word by word at the box's inner width. Sizes are a:rPr sz (hundredths of a point); a normAutofit's
 * fontScale and lnSpcReduction shrink them as Word does.
 */
function dmlText(txBody, box, ctx) {
  const bodyPr = child(txBody, 'bodyPr');
  const ins = (k, d) => px(bodyPr?.getAttribute(k) ?? d);
  const [l, t, r, b] = [ins('lIns', 91440), ins('tIns', 45720), ins('rIns', 91440), ins('bIns', 45720)];
  const auto = child(bodyPr, 'normAutofit');
  const scale = auto ? (Number(auto.getAttribute('fontScale')) || 100000) / 100000 : 1;
  const lnRed = auto ? (Number(auto.getAttribute('lnSpcReduction')) || 0) / 100000 : 0;
  const wrap = bodyPr?.getAttribute('wrap') !== 'none';
  const left0 = box.x + l, right0 = box.x + box.w - r;
  const lines = []; // { y (top, from 0), h, base, pieces: [{ x, text, open, ink }], align, left, right }
  let y = 0;
  let lastAft = 0;
  const pts = (v) => (Number(v) || 0) / 100 * (4 / 3); // hundredths of a point -> px
  for (const p of els(txBody).filter((n) => n.localName === 'p')) {
    const pPr = child(p, 'pPr');
    const runs = els(p).filter((n) => n.localName === 'r' || n.localName === 'br' || n.localName === 'fld');
    const endPr = child(p, 'endParaRPr');
    const firstRPr = child(runs.find((n) => n.localName !== 'br'), 'rPr') || endPr;
    const size0 = (Number(firstRPr?.getAttribute('sz')) || 1800) / 100 * scale;
    const align = { ctr: 'center', r: 'right', just: 'left', dist: 'left' }[pPr?.getAttribute('algn')] || 'left';
    const marL = px(pPr?.getAttribute('marL') || 0);
    const indent = px(pPr?.getAttribute('indent') || 0);
    const spc = (name, sizePt) => {
      const e = child(pPr, name);
      const pct = child(e, 'spcPct');
      const ptsEl = child(e, 'spcPts');
      if (ptsEl) return pts(ptsEl.getAttribute('val'));
      return pct ? (Number(pct.getAttribute('val')) / 100000) * sizePt * 1.2 * (4 / 3) : 0;
    };
    const lnPct = child(child(pPr, 'lnSpc'), 'spcPct');
    const lnPts = child(child(pPr, 'lnSpc'), 'spcPts');
    const mult = (lnPct ? Number(lnPct.getAttribute('val')) / 100000 : 1) - lnRed;
    y += spc('spcBef', size0);
    const bu = child(pPr, 'buChar');
    let cur = { pieces: [], x: left0 + marL + indent, max: 0, lf: 1.2 };
    const para = [cur];
    if (bu && runs.length) {
      const buFont = child(pPr, 'buFont')?.getAttribute('typeface') || ctx.font;
      const sizePx = (size0 * 4) / 3;
      cur.pieces.push({ x: cur.x, text: bu.getAttribute('char') || '\u2022', open: `<tspan font-family="${esc(fontStack(buFont))}" font-size="${sizePx.toFixed(2)}" fill="#${ctx.textColor || '000000'}">`, ink: cur.x + sizePx * 0.5 });
      cur.x = left0 + marL;
      cur.max = sizePx;
    }
    const newLine = () => { cur = { pieces: [], x: left0 + marL, max: 0, lf: 1.2 }; para.push(cur); };
    for (const run of runs) {
      if (run.localName === 'br') { newLine(); continue; }
      const rPr = child(run, 'rPr');
      const sizePt = (Number(rPr?.getAttribute('sz')) || Number(firstRPr?.getAttribute('sz')) || 1800) / 100 * scale;
      const sizePx = (sizePt * 4) / 3;
      const fill = child(rPr, 'solidFill');
      const c = fill ? color(els(fill)[0], ctx.colors)?.hex : null;
      const bold = rPr?.getAttribute('b') === '1';
      const italic = rPr?.getAttribute('i') === '1';
      const latin = child(rPr, 'latin')?.getAttribute('typeface');
      const font = !latin || latin.startsWith('+mn') ? ctx.font : latin.startsWith('+mj') ? ctx.theme?.major || ctx.font : latin;
      const open = `<tspan font-family="${esc(fontStack(font))}" font-size="${sizePx.toFixed(2)}" fill="#${c || ctx.textColor || '000000'}"${bold ? ' font-weight="bold"' : ''}${italic ? ' font-style="italic"' : ''}>`;
      const text = textOfRun(run);
      cur.lf = Math.max(cur.lf, lineFactor(font));
      for (const word of text.match(/\S*\s*/g).filter(Boolean)) {
        const width = textWidth(word, sizePx, bold, font);
        const inked = textWidth(word.trimEnd(), sizePx, bold, font);
        if (wrap && cur.x + inked > right0 + 0.5 && cur.pieces.length) newLine();
        const last = cur.pieces[cur.pieces.length - 1];
        if (last && last.open === open && last.end === cur.x) last.text += word;
        else cur.pieces.push({ x: cur.x, text: word, open });
        const pc = cur.pieces[cur.pieces.length - 1];
        pc.ink = cur.x + inked;
        cur.x += width;
        pc.end = cur.x;
        cur.max = Math.max(cur.max, sizePx);
      }
    }
    for (const ln of para) {
      const sizePx = ln.max || (size0 * 4) / 3;
      const h = lnPts ? pts(lnPts.getAttribute('val')) : sizePx * ln.lf * mult;
      lines.push({ y, h, base: y + h - sizePx * ln.lf * 0.222 * Math.min(1, mult), pieces: ln.pieces, align, left: left0 + marL, right: right0 });
      y += h;
    }
    lastAft = spc('spcAft', size0);
    y += lastAft;
  }
  // (The last paragraph's space after doesn't count where the text is centred or set at the bottom, as in Word.)
  const used = y - lastAft;
  const anchor = bodyPr?.getAttribute('anchor') || 't';
  const room = box.h - t - b;
  const top = box.y + t + (anchor === 'ctr' ? (room - used) / 2 : anchor === 'b' ? room - used : 0);
  const out = [];
  for (const ln of lines) {
    if (!ln.pieces.length) continue;
    const ink = ln.pieces[ln.pieces.length - 1].ink;
    const shift = ln.align === 'center' ? (ln.right - ink) / 2 - (ln.pieces[0].x - ln.left) / 2 : ln.align === 'right' ? ln.right - ink : 0;
    for (const pc of ln.pieces) out.push(`<text x="${(pc.x + shift).toFixed(1)}" y="${(top + ln.base).toFixed(1)}" xml:space="preserve">${pc.open}${esc(pc.text)}</tspan></text>`);
  }
  return out.join('');
}

/** A text run's (a:r, a:fld) text. */
const textOfRun = (run) => els(run).filter((n) => n.localName === 't').map((n) => n.textContent).join('');

const GROUPS = ['wgp', 'grpSp', 'spTree'];
const NO_MAP = { sx: 1, sy: 1, tx: 0, ty: 0 };
const on = (xfrm, k) => /^(1|true)$/.test(xfrm?.getAttribute(k) || '');

/**
 * Where a shape or group (its a:xfrm) is drawn: its box in the drawing's px, mapped there by `m` (a group's child
 * coordinates, chOff/chExt, to the drawing: a scale and a move, never a turn), and how it's turned (degrees clockwise)
 * and flipped about the box's centre. A group's own map for its children comes too: its box is mapped first and only
 * then turned, as Word draws a group (a group stretched to 1.8 times its height keeps its turned rectangles square-
 * cornered, groupturn-word), so a child turned inside a stretched group isn't skewed.
 */
function frameOf(xfrm, m = NO_MAP, top = false) {
  const off = child(xfrm, 'off');
  const ext = child(xfrm, 'ext');
  const n = (el, k) => px(el?.getAttribute(k));
  // (The drawing's own group's frame is the drawing itself, wherever its a:off says.)
  const turn = Number(xfrm?.getAttribute('rot') || 0) / 60000;
  // A shape turned nearer upright than level (45 to 135 degrees, or 225 to 315) is stretched along its other side: its
  // width by the group's height scale (a right arrow at 60 degrees in a group stretched to 1.8 times its height comes out
  // 1.8 times as long, groupturn-word). Its centre is mapped as it is.
  const q = ((turn % 180) + 180) % 180;
  const swap = !top && q >= 45 && q < 135;
  const w = (swap ? m.sy : m.sx) * n(ext, 'cx');
  const h = (swap ? m.sx : m.sy) * n(ext, 'cy');
  const cx = m.tx + m.sx * ((top ? 0 : n(off, 'x')) + n(ext, 'cx') / 2);
  const cy = m.ty + m.sy * ((top ? 0 : n(off, 'y')) + n(ext, 'cy') / 2);
  const x = cx - w / 2, y = cy - h / 2;
  // (The drawing's own shape or group is turned by the page, not here: its picture is turned whole, shapeTurn.)
  const f = { x, y, w, h, rot: top ? 0 : turn, fh: on(xfrm, 'flipH'), fv: on(xfrm, 'flipV') };
  const chOff = child(xfrm, 'chOff');
  const chExt = child(xfrm, 'chExt');
  if (chOff || chExt) {
    const sx = n(chExt, 'cx') ? w / n(chExt, 'cx') : m.sx;
    const sy = n(chExt, 'cy') ? h / n(chExt, 'cy') : m.sy;
    f.map = { sx, sy, tx: x - sx * n(chOff, 'x'), ty: y - sy * n(chOff, 'y') };
  }
  return f;
}

/** An SVG transform turning and flipping frame `f` about its centre, in the drawing's coordinates ('' for none). */
function turnAbout(f) {
  const cx = f.x + f.w / 2, cy = f.y + f.h / 2;
  const r = (v) => +v.toFixed(2);
  return [f.rot ? `rotate(${r(f.rot)} ${r(cx)} ${r(cy)})` : '', f.fh || f.fv ? `matrix(${f.fh ? -1 : 1} 0 0 ${f.fv ? -1 : 1} ${f.fh ? r(2 * cx) : 0} ${f.fv ? r(2 * cy) : 0})` : ''].filter(Boolean).join(' ');
}

/**
 * The width (px) of a shape's outline that its text is inset from (textBox: half of it, the half inside the box), 0 for
 * none. (The editing overlay places its text by this too: docx.js textBoxes.)
 */
export function outlineWidth(spPr) {
  const ln = child(spPr, 'ln');
  return ln && !child(ln, 'noFill') && child(ln, 'solidFill') ? px(ln.getAttribute('w') || EMU_PER_PX) : 0;
}

function shapeSvg(el, ctx, top = false, m = NO_MAP) {
  const name = el.localName;
  if (GROUPS.includes(name)) {
    const xfrm = child(child(el, 'grpSpPr'), 'xfrm');
    const f = xfrm ? frameOf(xfrm, m, name === 'wgp' || top) : null;
    const inner = els(el).filter((n) => ['wsp', 'sp', 'pic', 'grpSp'].includes(n.localName)).map((n) => shapeSvg(n, ctx, false, f?.map || m)).join('');
    // A group turned or flipped inside another turns its shapes with it, about its own centre.
    const t = f ? turnAbout(f) : '';
    return t && inner ? `<g transform="${t}">${inner}</g>` : inner;
  }
  const spPr = child(el, 'spPr');
  const xfrm = child(spPr, 'xfrm');
  if (!xfrm) return '';
  const { x, y, h, rot, fh, fv } = frameOf(xfrm, m, top);
  let { w } = frameOf(xfrm, m, top);
  let t = `translate(${x.toFixed(2)} ${y.toFixed(2)})${rot ? ` rotate(${rot} ${w / 2} ${h / 2})` : ''}${fh || fv ? ` matrix(${fh ? -1 : 1} 0 0 ${fv ? -1 : 1} ${fh ? w : 0} ${fv ? h : 0})` : ''}`;
  // A flipped shape's text isn't mirrored: Word keeps it readable, turned with the shape, and upside down when the
  // shape is flipped top to bottom (groupturn-word: "Text" in a rectangle flipped both ways and in one flipped
  // vertically, turned 15 degrees, reads upside down; in one flipped left to right, as it was).
  const textTurn = (rot + (fv ? 180 : 0)) % 360;
  let tt = `translate(${x.toFixed(2)} ${y.toFixed(2)})${textTurn ? ` rotate(${textTurn} ${w / 2} ${h / 2})` : ''}`;
  if (name === 'pic') {
    const blip = desc(el, 'blip');
    const href = blip && ctx.image(blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed'));
    return href ? `<image transform="${t}" width="${w.toFixed(2)}" height="${h.toFixed(2)}" preserveAspectRatio="none" href="${href}"/>` : '';
  }
  if (name !== 'wsp' && name !== 'sp') return '';
  const style = child(el, 'style');
  const pt = paints(spPr, style, ctx.colors);
  const txbx = child(el, 'txbx');
  const line = outlineWidth(spPr);
  // (Text without a colour of its own takes the shape's font colour, wps:style's fontRef: white on Word's default shapes.)
  const fontRef = style && child(style, 'fontRef');
  const textColor = fontRef && els(fontRef)[0] ? color(els(fontRef)[0], ctx.colors)?.hex : null;
  const tctx = textColor ? { ...ctx, textColor } : ctx;
  const bodyPr = child(el, 'bodyPr');
  let text = txbx ? textBox(txbx, bodyPr, w, h, tctx, line) : '';
  // WordArt's box (not wrapped and resized to fit its text: wrap="none", spAutoFit) as Word lays it out: as wide as its
  // widest line, but never wider than the column (ctx.colW), where its lines wrap instead, word by word, a word wider
  // than that letter by letter (Word's PDF of wordart-leftovers/wide.docx: Style 1 at 36 and 54pt saved 2in and 9in wide,
  // at the column's edge, 2in in and against the page: each box 6.5in wide, a 6.5in column's, its lines wrapped there
  // and centred in it, the box running past the page's edge when moved in). One narrower than its box is left as saved:
  // its lines are lined up against the widest (textBox `shrink`).
  if (txbx && top && ctx.colW > 0 && bodyPr?.getAttribute('wrap') === 'none' && child(bodyPr, 'spAutoFit') && measures()) {
    const natural = tctx.textRight + 0.5;
    if (natural > w + 0.5 || w > ctx.colW + 0.5) {
      w = Math.round(Math.min(ctx.colW, natural) * 100) / 100;
      text = textBox(txbx, bodyPr, w, h, tctx, line, true);
      ctx.widened = w;
      t = `translate(${x.toFixed(2)} ${y.toFixed(2)})${rot ? ` rotate(${rot} ${w / 2} ${h / 2})` : ''}${fh || fv ? ` matrix(${fh ? -1 : 1} 0 0 ${fv ? -1 : 1} ${fh ? w : 0} ${fv ? h : 0})` : ''}`;
      tt = `translate(${x.toFixed(2)} ${y.toFixed(2)})${textTurn ? ` rotate(${textTurn} ${w / 2} ${h / 2})` : ''}`;
    }
  }
  // "Resize shape to fit text" (a:spAutoFit): Word makes the box as tall as its text when it lays the page out, whatever
  // height the file gives it (a text box of four paragraphs in 4.2in: Word's body text wraps beside it to its end).
  let hh = h;
  if (txbx && child(bodyPr, 'spAutoFit') && tctx.textEnd > h + 0.5) { hh = tctx.textEnd; ctx.grown = Math.max(ctx.grown || 0, hh - h); }
  const paths = geometry(spPr, w, hh);
  const strokeAttrs = pt.stroke ? ` stroke-width="${+pt.width.toFixed(2)}"${pt.join === 'miter' ? ' stroke-miterlimit="8"' : ` stroke-linejoin="${pt.join}"`}` : '';
  let body = paths.map((q) => {
    const stroke = q.stroke ? paint('stroke', pt.stroke) + strokeAttrs : 'stroke="none"';
    if (pt.fill?.grad && q.fill !== 'none') {
      const gf = gradientFill(pt.fill, q.d, w, hh, Number(xfrm.getAttribute('rot') || 0) / 60000, ctx, q.fill);
      return `<defs>${gf.defs}</defs>${gf.under || ''}<path d="${q.d}" ${gf.fill || 'fill="none"'} ${stroke} vector-effect="non-scaling-stroke"/>`;
    }
    return `<path d="${q.d}" ${paint('fill', modeFill(pt.fill, q.fill))} ${stroke} vector-effect="non-scaling-stroke"/>`;
  }).join('');
  // Arrowheads, where the outline has them (on a line, or any open path).
  const open = paths.find((q) => q.ends && !/Z$/.test(q.d));
  if (pt.stroke && open && pt.ln) body += arrowhead(child(pt.ln, 'headEnd'), open.ends[0], pt.width, pt.stroke) + arrowhead(child(pt.ln, 'tailEnd'), open.ends[1], pt.width, pt.stroke);
  // DrawingML text (a:txBody: SmartArt's shapes), in its own rectangle (dsp:txXfrm) where it has one, unrotated.
  const txBody = child(el, 'txBody');
  let dml = '';
  if (txBody) {
    const tx = child(child(el, 'txXfrm'), 'off') ? child(el, 'txXfrm') : null;
    const tf = tx && frameOf(tx, m);
    const box = tf ? { x: tf.x - x, y: tf.y - y, w: tf.w, h: tf.h } : { x: 0, y: 0, w, h };
    dml = `<g transform="translate(${x.toFixed(2)} ${y.toFixed(2)})">${dmlText(txBody, box, tctx)}</g>`;
  }
  return `<g transform="${t}">${body}</g>${text ? `<g transform="${tt}">${text}</g>` : ''}${dml}`;
}

/**
 * SVG for a drawing whose graphic is a shape or a group of shapes, sized to its extent.
 * @param {Element} drawing w:drawing
 * @param {{ image: (rId: string) => string|null, colors?: object, font?: string }} ctx image() returns a data: URL
 */
export function shapesSvg(drawing, ctx, width, height, pad = 0) {
  const data = desc(drawing, 'graphicData');
  const root = data && els(data).find((n) => n.localName === 'wgp' || n.localName === 'wsp');
  // (A straight line across or down the page is as tall or wide as nothing, wp:extent cy or cx 0: drawn too, in the room
  // round it for its outline and arrowheads. Word's floating ones had not been drawn at all.)
  if (!root || !(width >= 0) || !(height >= 0) || !(width || height)) return null;
  const c = { colors: {}, font: 'Calibri', ...ctx };
  const inner = shapeSvg(root, c, true);
  if (!inner) return null;
  // A single text box that grows to fit its text (shapeSvg) is that much taller; ctx.grown tells the caller.
  if (root.localName === 'wsp' && c.grown) { height = Math.round((height + c.grown) * 100) / 100; ctx.grownTo = height; }
  // (And WordArt's box as wide as its text or the column, ctx.colW: its new width, ctx.widenedTo.)
  if (root.localName === 'wsp' && c.widened) { width = c.widened; ctx.widenedTo = width; }
  // (A whole number of pixels, the part of one added left empty at the right and the foot: the browser takes a drawing's
  // size down to whole pixels and fits what it holds into that, so one 278.64px tall was drawn 0.23% smaller, its text
  // with it; a text box's twelfth line was a pixel above Word's. schema.js gives its picture the same box.)
  const W = Math.ceil(width + 2 * pad - 0.005), H = Math.ceil(height + 2 * pad - 0.005);
  // (No ligatures or kerning, as Word sets text and the body has it: a text box's "ti" came out as one glyph.)
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="${-pad} ${-pad} ${W} ${H}" overflow="visible" style="font-variant-ligatures:none;font-kerning:none">${inner}</svg>`;
}

/**
 * SVG for a SmartArt diagram (a:graphicData .../diagram), from the drawing Word saves with it (diagrams/drawingN.xml:
 * dsp:drawing's dsp:spTree, its shapes placed in the diagram's frame), sized to its extent.
 */
export function diagramSvg(drawingPart, ctx, width, height) {
  const tree = desc(drawingPart, 'spTree');
  if (!tree || !width || !height) return null;
  const inner = shapeSvg(tree, { colors: {}, font: 'Calibri', ...ctx });
  if (!inner) return null;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" overflow="visible" style="font-variant-ligatures:none;font-kerning:none">${inner}</svg>`;
}

/**
 * How far a drawing's outlines and arrowheads are drawn past its box (px), which Word draws there (an image can't: the
 * room is added around it, shapesSvg's pad, and taken back off its margins, so it takes the same room as before).
 */
export function shapesPad(drawing) {
  let pad = 0;
  // (Its shapes' outlines: Word draws one centred on the box's edge, half outside.)
  for (const ln of [...drawing.getElementsByTagName('*')].filter((n) => n.localName === 'ln' && n.parentNode?.localName === 'spPr')) {
    if (child(ln, 'noFill')) continue;
    pad = Math.max(pad, (ln.getAttribute('w') ? px(ln.getAttribute('w')) : 1.33) / 2 + 1);
  }
  for (const ln of [...drawing.getElementsByTagName('*')].filter((n) => n.localName === 'ln' && n.parentNode?.localName === 'spPr')) {
    if (child(ln, 'noFill')) continue;
    const w = ln.getAttribute('w') ? px(ln.getAttribute('w')) : 1.33;
    const ends = ['headEnd', 'tailEnd'].map((k) => child(ln, k)).filter((e) => e && e.getAttribute('type') && e.getAttribute('type') !== 'none');
    if (!ends.length) continue;
    const k = { sm: 2, med: 3, lg: 5 };
    pad = Math.max(pad, w / 2 + Math.max(...ends.map((e) => Math.max(w, 8 / 3) * Math.max(k[e.getAttribute('w')] || 3, k[e.getAttribute('len')] || 3))));
  }
  const data = desc(drawing, 'graphicData');
  // A group whose shapes, turned, reach past its box (a 30-degree rectangle at its edge; groupturn-word): room for them,
  // which Word draws past the box too, and their outlines.
  const group = data && [...data.childNodes].find((n) => n.localName === 'wgp');
  const gx = group && child(child(group, 'grpSpPr'), 'xfrm');
  if (gx) {
    const f = frameOf(gx, NO_MAP, true);
    const b = groupReach(group, f.map || NO_MAP, [1, 0, 0, 1, 0, 0]);
    const out = Math.max(-b[0], -b[1], b[2] - f.w, b[3] - f.h);
    if (out > 1) pad = Math.max(pad, out + 1.33 / 2 + 1);
  }
  // A single shape drawn past its box (a callout's pointer, or one dragged out by its adjust handle): room for that
  // too, only when it's more than a pixel out (outlines just inside or on the box keep none, as above).
  const sp = data && [...data.childNodes].find((n) => n.localName === 'wsp');
  const spPr = sp && child(sp, 'spPr');
  const ext = spPr && desc(spPr, 'ext');
  if (ext) {
    const w = px(ext.getAttribute('cx')), h = px(ext.getAttribute('cy'));
    const out = w && h ? Math.max(0, ...geometry(spPr, w, h).map((q) => overflow(q.d, w, h))) : 0;
    if (out > 1) {
      const ln = child(spPr, 'ln');
      pad = Math.max(pad, out + (ln?.getAttribute('w') ? px(ln.getAttribute('w')) : 1.33) / 2 + 1);
    }
    // A turned shape: room for its outline too, which would otherwise be cut off at its box and show, turned, as flat
    // spots on its edge (an oval at 45 degrees).
    if (shapeTurn(drawing)) {
      const ln = child(spPr, 'ln');
      pad = Math.max(pad, (ln?.getAttribute('w') ? px(ln.getAttribute('w')) : 1.33) / 2 + 1);
    }
  }
  return Math.ceil(pad);
}

/**
 * The box [minX, minY, maxX, maxY] (px, in the drawing) a group's shapes cover, turned and flipped as shapeSvg draws
 * them: each shape's box mapped by `m` (frameOf), its corners turned about its centre, then by `a` (the turns of the
 * groups it's in, an SVG matrix [a b c d e f]).
 */
export function groupReach(group, m, a) {
  const box = [Infinity, Infinity, -Infinity, -Infinity];
  const turn = (f) => {
    const t = ((f.rot || 0) * Math.PI) / 180, cx = f.x + f.w / 2, cy = f.y + f.h / 2;
    const [cos, sin] = [Math.cos(t), Math.sin(t)];
    return [cos, sin, -sin, cos, cx - cos * cx + sin * cy, cy - sin * cx - cos * cy]; // (a flip about the centre moves no corner)
  };
  const mul = (p, q) => [p[0] * q[0] + p[2] * q[1], p[1] * q[0] + p[3] * q[1], p[0] * q[2] + p[2] * q[3], p[1] * q[2] + p[3] * q[3], p[0] * q[4] + p[2] * q[5] + p[4], p[1] * q[4] + p[3] * q[5] + p[5]];
  const walk = (g, m, a) => {
    for (const n of els(g)) {
      if (n.localName === 'grpSp') {
        const x = child(child(n, 'grpSpPr'), 'xfrm');
        if (!x) { walk(n, m, a); continue; }
        const f = frameOf(x, m);
        walk(n, f.map || m, mul(a, turn(f)));
      } else if (['wsp', 'sp', 'pic'].includes(n.localName)) {
        const x = child(child(n, 'spPr'), 'xfrm');
        if (!x) continue;
        const f = frameOf(x, m);
        const k = mul(a, turn(f));
        for (const [px0, py0] of [[f.x, f.y], [f.x + f.w, f.y], [f.x, f.y + f.h], [f.x + f.w, f.y + f.h]]) {
          const X = k[0] * px0 + k[2] * py0 + k[4], Y = k[1] * px0 + k[3] * py0 + k[5];
          box[0] = Math.min(box[0], X); box[1] = Math.min(box[1], Y); box[2] = Math.max(box[2], X); box[3] = Math.max(box[3], Y);
        }
      }
    }
  };
  walk(group, m, a);
  return box[0] === Infinity ? [0, 0, 0, 0] : box;
}

/** How far an SVG path (as geometry.js writes them) goes outside a w x h box, by its points (an arc by its ends). */
function overflow(d, w, h) {
  let out = 0;
  const see = (x, y) => { out = Math.max(out, -x, -y, x - w, y - h); };
  for (const [, k, args] of d.matchAll(/([MLCQA])([^MLCQAZ]*)/g)) {
    const n = args.trim().split(/[\s,]+/).filter(Boolean).map(Number);
    if (k === 'A') { if (n.length >= 7) see(n[5], n[6]); continue; }
    for (let i = 0; i + 1 < n.length; i += 2) see(n[i], n[i + 1]);
  }
  return out;
}

/**
 * How far a drawing's shape (or group) is turned, in degrees clockwise (its a:xfrm rot): its picture is drawn unturned
 * and turned whole about its centre on the page (reader.js), as Word turns it, so nothing of it is cut off at its box.
 */
export function shapeTurn(drawing) {
  const deg = Number(ownXfrm(drawing)?.getAttribute('rot') || 0) / 60000;
  return ((deg % 360) + 360) % 360;
}

/** Whether a drawing's group has a shape or group turned inside it (any a:xfrm rot under its wpg:wgp but its own). */
export function turnedInside(drawing) {
  const data = desc(drawing, 'graphicData');
  const group = data && [...data.childNodes].find((n) => n.localName === 'wgp');
  return !!group && [...group.getElementsByTagName('*')].some((n) => n.localName === 'xfrm' && n.parentNode?.parentNode !== group && Number(n.getAttribute('rot')) % 21600000 !== 0);
}

/** A picture's flips (its pic:spPr a:xfrm flipH, flipV): drawn mirrored by the page, as Word draws it. */
export function pictureFlip(drawing) {
  const xfrm = ownXfrm(drawing);
  if (!xfrm || xfrm.parentNode?.parentNode?.localName !== 'pic') return { h: false, v: false };
  const on = (k) => /^(1|true)$/.test(xfrm.getAttribute(k) || '');
  return { h: on('flipH'), v: on('flipV') };
}

/** The a:xfrm of a drawing's own shape, group or picture (wps:spPr, wpg:grpSpPr, pic:spPr). */
function ownXfrm(drawing) {
  const data = desc(drawing, 'graphicData');
  const root = data && [...data.childNodes].find((n) => n.localName === 'wsp' || n.localName === 'wgp' || n.localName === 'pic');
  return root && child(child(root, root.localName === 'wgp' ? 'grpSpPr' : 'spPr'), 'xfrm');
}
