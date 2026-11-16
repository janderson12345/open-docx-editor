// Word 2010's text effects on a run (w14:textFill, w14:textOutline, w14:shadow, w14:glow, w14:reflection: what Home >
// Text Effects and WordArt write), as CSS. They're kept as the run's raw XML (reader.js rPrMarks), so this only draws
// them; saving writes them back as they were.
//
// Checked against Word (odx-test/mkeffects14.ps1: Word itself setting each effect through its Font object): a w14
// colour's w14:alpha is its transparency (Word's 50% transparent red is alpha 50000), not DrawingML's opacity; Word
// also writes the old w:outline beside a w14:textOutline, which then is the outline drawn.
import { modify, OFFICE } from './shapes.js';
import { startTags, tagAttr } from './xml.js';

// The open document's theme colours (docx.js, from its theme part): scheme name -> hex.
const THEME = {};
export function setThemeColors(colors) {
  for (const k of Object.keys(THEME)) delete THEME[k];
  Object.assign(THEME, colors || {});
}

const SCHEME = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };
const num = (xml, name) => Number(new RegExp(`\\bw14:${name}="(-?\\d+)"`).exec(xml)?.[1] ?? NaN);

// Tags read one pass each (xml.js startTag), not by regexes reading two attributes of a tag, which took time growing
// with the square of a crafted tag's length.
/** A fragment's first w14:srgbClr or w14:schemeClr with a w14:val: [, name, val, , inner XML], or null. */
function firstColor(xml) {
  const tags = [...['srgbClr', 'schemeClr'].flatMap((k) => [...startTags(xml, `w14:${k}`)].map((t) => [k, t]))].sort((p, q) => p[1].at - q[1].at);
  for (const [k, t] of tags) {
    const val = tagAttr(t.tag, 'w14:val');
    if (!val) continue;
    if (t.selfClosing) return [null, k, val, null, ''];
    const close = xml.indexOf(`</w14:${k}>`, t.end);
    if (close >= 0) return [null, k, val, null, xml.slice(t.end, close)];
  }
  return null;
}
/** A gradient fill's stops: [, position (as written), inner XML] each. */
function gradStops(fill) {
  const out = [];
  for (const t of startTags(fill, 'w14:gs')) {
    const pos = tagAttr(t.tag, 'w14:pos');
    const close = fill.indexOf('</w14:gs>', t.end);
    if (close < 0) break;
    if (/^\d+$/.test(pos || '')) out.push([null, pos, fill.slice(t.end, close)]);
  }
  return out;
}

/** The first colour in a w14 fragment (srgbClr or schemeClr, with its changes) as a CSS colour; null if none. */
export function w14Color(xml) {
  const m = firstColor(xml || '');
  if (!m) return null;
  let hex = m[1] === 'srgbClr' ? m[2] : THEME[SCHEME[m[2]] || m[2]] || OFFICE[SCHEME[m[2]] || m[2]];
  if (!/^[0-9A-Fa-f]{6}$/.test(hex || '')) return null;
  let alpha = 1;
  for (const c of (m[4] || '').matchAll(/<w14:(\w+)\b[^>]*?\bw14:val="(-?\d+)"/g)) {
    const v = Number(c[2]) / 100000;
    if (c[1] === 'alpha') alpha = 1 - v;
    else if (['shade', 'tint', 'lumMod', 'lumOff', 'hueOff', 'hueMod', 'satOff', 'satMod'].includes(c[1])) hex = modify(hex, c[1], v);
  }
  const rgb = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  return alpha >= 1 ? `#${hex.toUpperCase()}` : `rgba(${rgb.join(',')},${+Math.max(0, alpha).toFixed(3)})`;
}

// (EMUs to points: 12700 a point; to CSS pixels: 9525 a pixel.)
const pt = (emu) => +(emu / 12700).toFixed(2);
const px = (emu) => +(emu / 9525).toFixed(2);
// A CSS colour as an SVG paint and its opacity.
const solid = (c) => {
  const m = /^rgba\((\d+),(\d+),(\d+),([\d.]+)\)$/.exec(c);
  return m ? `"#${[m[1], m[2], m[3]].map((v) => Number(v).toString(16).padStart(2, '0')).join('')}" ` : `"${c}" `;
};
const opacity = (c) => /^rgba\(/.exec(c) ? Number(/,([\d.]+)\)$/.exec(c)[1]) : 1;

// An inner shadow's filter primitives: the letters' outside (their alpha inverted) moved by the offset, blurred, in the
// shadow's colour, and kept only where the letters are. `merge` draws the letters under it (a text box's SVG text);
// without, it is the shadow alone (body text: laid over the letters, editor.css .odx-is).
// Word's PDFs of odx-test/mkinnershadow.ps1 (36pt Arial bold, Word's own inner shadows through its Font.TextShadow, in
// body text and in a text box alike): the shadow lies along the edges the direction points away from (225 degrees, up
// and left, darkens the letters' left and top edges; 0 their right edges), so the outside is moved against the
// direction; and it is a third of the size written: the band half dark about a third of the distance in (1.3pt for
// 4pt, 1pt for 2.83pt unblurred), fading over about a sixth of the blur (a 5pt blur over 0.8pt).
const innerShadowFx = (dx, dy, blur, c, merge) => `<feComponentTransfer in="SourceAlpha" result="i0"><feFuncA type="table" tableValues="1 0"/></feComponentTransfer><feOffset in="i0" dx="${px(-dx / 3)}" dy="${px(-dy / 3)}" result="i1"/>`
  + `<feGaussianBlur in="i1" stdDeviation="${px(blur / 6)}" result="i2"/><feFlood flood-color=${solid(c)}flood-opacity="${opacity(c)}"/><feComposite in2="i2" operator="in" result="i3"/>`
  + `<feComposite in="i3" in2="SourceAlpha" operator="in" result="i4"/>${merge ? '<feMerge><feMergeNode in="SourceGraphic"/><feMergeNode in="i4"/></feMerge>' : ''}`;

// A bevel (w14:props3d's w14:bevelT: Word 2010's 3-D format on text) as filter primitives on the letters
// (SourceGraphic), its result `out`: the letters lit from above and a little to the left, as Word's default lighting
// lights them, over a bevel as wide as it says (w), as steep as its height (h) to that: their outline blurred into a
// slope that wide, its light and shade added to the letters' own colour, which stays as it was where they're flat; a
// shiny material (not matte or flat) also gets a highlight where the slope faces the light. Word's print of
// wordart-leftovers/bevel.docx (gallery Styles 5 and 10, 36 and 72pt, in the text and in text boxes): the soft bevel
// (circle, 2pt by 3pt, softEdge) a rounded rim, bright along edges facing up, dark along those facing down; the sharp
// one (angle, 5pt by 1pt, matte) flat facets, drawn as such (facetFx). The bevel is as wide at 72pt as at
// 36pt, so it's in points, not ems. Word's extrusion (extrusionH) and contour, seen face on, don't show: not drawn.
// (An approximation: Word renders the letters in 3-D; a lit relief of their outline is what can be drawn here.)
export function bevelFx(props3d, out = 'odxbev', box = false) {
  // (Its tags read one pass each, xml.js startTag. A bevel's size defaults to 6pt by 6pt, its shape to a circle.)
  const bt = props3d ? [...startTags(props3d, 'w14:bevelT')][0]?.tag : null;
  if (!bt) return '';
  // (At most 64px: a file's huge bevel would make every repaint of its filters' erosions and blurs crawl.)
  const size = (k) => { const v = tagAttr(bt, k); return Math.min(64, /^\d{1,12}$/.test(v || '') ? px(Number(v)) : px(76200)); };
  const w = size('w14:w'), h = size('w14:h');
  const prst = tagAttr(bt, 'w14:prst') || 'circle';
  if (!(w > 0) || !(h > 0)) return '';
  const material = tagAttr([...startTags(props3d, 'w14:props3d')][0]?.tag || '', 'w14:prstMaterial') || 'warmMatte';
  const shiny = !['matte', 'flat', 'dkEdge', 'warmMatte', 'legacyMatte'].includes(material);
  // (Straight-sided bevels: flat facets, facetFx.)
  if (['angle', 'coolSlant', 'slope', 'divot'].includes(prst)) return facetFx(w, h, shiny, out, box);
  // (The slope a third of the bevel's width each side of the edge.)
  const sd = +(w / 3).toFixed(2);
  const scale = +Math.min(12, h).toFixed(2);
  // (The light's height, and what it gives a flat face, subtracted so that stays the letters' own colour.)
  const el = 55, rad = (el * Math.PI) / 180, flat = Math.sin(rad);
  const k = material === 'matte' || material === 'flat' ? 0.4 : 0.45;
  const exp = 40, flatSpec = 0.9 * ((1 + flat) / Math.hypot(Math.cos(rad), 1 + flat)) ** exp;
  const spec = shiny
    ? `<feSpecularLighting in="bv0" surfaceScale="${scale}" specularConstant="0.9" specularExponent="${exp}" lighting-color="#fff" result="bv3"><feDistantLight azimuth="240" elevation="${el}"/></feSpecularLighting>`
      + `<feComposite in="bv3" in2="SourceAlpha" operator="in" result="bv4"/><feComposite in="bv2" in2="bv4" operator="arithmetic" k2="1" k3="0.7" k4="${+(-0.7 * flatSpec).toFixed(4)}" result="bv5"/>`
    : '';
  return `<feGaussianBlur in="SourceAlpha" stdDeviation="${sd}" result="bv0"/>`
    + `<feDiffuseLighting in="bv0" surfaceScale="${scale}" diffuseConstant="1" lighting-color="#fff" result="bv1"><feDistantLight azimuth="240" elevation="${el}"/></feDiffuseLighting>`
    + `<feComposite in="SourceGraphic" in2="bv1" operator="arithmetic" k2="1" k3="${k}" k4="${+(-k * flat).toFixed(4)}" result="bv2"/>`
    + spec
    + `<feComposite in="${shiny ? 'bv5' : 'bv2'}" in2="SourceAlpha" operator="in" result="${out}"/>`;
}

// An angled bevel's flat facets (bevelFx): a slope rising evenly from the outline to the bevel's width (w) and its
// height (h), the outline eroded in steps and the steps averaged, smoothed a little; lit, each facet one tone. Word's
// print of wordart-leftovers/bevel.docx (Style 10: angle 5pt by 1pt, matte, accent 3 196B24): facets facing up 7.5%
// lighter than the colour, facing left 5% lighter, those facing away from the light about as the face, which is 4%
// darker than the colour; where a stem is narrower than two bevels its facets meet in a ridge. So the light (from up
// and left) maps to that: no darker than the face, lighter as a facet turns to it. In a text box (`box`) Word lights
// the same bevel otherwise, from the left and harder (its print of word-wordart-styles.docx, Style 10 at 36pt): facets
// facing left 7% lighter than the colour, right 42% darker, up and down 16% darker.
function facetFx(w, h, shiny, out, box) {
  const n = 6, el = 55, az = box ? 180 : 230, rad = (el * Math.PI) / 180, flat = Math.sin(rad);
  let fx = '';
  for (let i = 1; i <= n; i++) {
    fx += `<feMorphology in="SourceAlpha" operator="erode" radius="${+((w * i) / n).toFixed(2)}" result="bf${i}"/>`;
    fx += i === 1 ? '' : `<feComposite in="bf${i}" in2="${i === 2 ? 'bf1' : `bs${i - 1}`}" operator="arithmetic" k2="${+(1 / n).toFixed(4)}" k3="${i === 2 ? +(1 / n).toFixed(4) : 1}" result="bs${i}"/>`;
  }
  // (The light a facet facing it straight up gets, for this slope; the tones between the face's and its as a table.)
  const s = h / w, up = (s * Math.cos(rad) * Math.sin(((az - 180) * Math.PI) / 180) + flat) / Math.hypot(1, s);
  // (In a box: the tone straight with the light, about a facet across it, at 2.2 times its change.)
  const across = flat / Math.hypot(1, s);
  const tone = (l) => +(Math.max(0, box ? 0.84 + 2.2 * (l - across)
    : 0.96 + (l > flat ? ((l - flat) / Math.max(1e-3, up - flat)) * 0.105 : (flat - l) * 0.25)) / 2).toFixed(4);
  const table = Array.from({ length: 41 }, (_, i) => tone(i / 40)).join(' ');
  const spec = shiny
    ? `<feSpecularLighting in="bf0" surfaceScale="${h}" specularConstant="0.9" specularExponent="40" lighting-color="#fff" result="bf7"><feDistantLight azimuth="${az}" elevation="${el}"/></feSpecularLighting>`
      + `<feComposite in="bf8" in2="bf7" operator="arithmetic" k2="1" k3="0.5" result="bf9"/>`
    : '';
  return fx
    + `<feGaussianBlur in="bs${n}" stdDeviation="${+((w / n) * 0.7).toFixed(2)}" result="bf0"/>`
    + `<feDiffuseLighting in="bf0" surfaceScale="${h}" diffuseConstant="1" lighting-color="#fff" result="bfl"><feDistantLight azimuth="${az}" elevation="${el}"/></feDiffuseLighting>`
    + `<feComponentTransfer in="bfl" result="bft"><feFuncR type="table" tableValues="${table}"/><feFuncG type="table" tableValues="${table}"/><feFuncB type="table" tableValues="${table}"/></feComponentTransfer>`
    // (Both opaque, the letters' colour as it is, so the product brightens as well as darkens; their edge restored after.)
    + `<feComponentTransfer in="SourceGraphic" result="bfs"><feFuncA type="table" tableValues="1 1"/></feComponentTransfer>`
    + `<feComposite in="bfs" in2="bft" operator="arithmetic" k1="2" result="bf8"/>`
    + spec
    + `<feComposite in="${shiny ? 'bf9' : 'bf8'}" in2="SourceAlpha" operator="in" result="${out}"/>`;
}

// Filters for body text, defined once each in the page (an SVG of no size at the end of the body, which prints with
// it): their markup -> id.
const INNER = new Map();
/** The id of the page's filter of primitives `fx` (named `odx-{prefix}N`), defined in the page if new. */
function pageFilter(fx, prefix) {
  if (!INNER.has(fx)) INNER.set(fx, `odx-${prefix}${INNER.size}`);
  const id = INNER.get(fx);
  if (typeof document !== 'undefined' && document.body && !document.getElementById(id)) {
    let svg = document.getElementById('odx-is-defs');
    if (!svg) {
      svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      svg.id = 'odx-is-defs';
      svg.setAttribute('aria-hidden', 'true');
      svg.setAttribute('style', 'position:absolute;width:0;height:0;overflow:hidden');
      document.body.append(svg);
    }
    // (The region round the copy generous enough for its descenders: editor.css trims its box to the baseline.)
    svg.insertAdjacentHTML('beforeend', `<filter id="${id}" x="-50%" y="-50%" width="200%" height="200%" color-interpolation-filters="sRGB">${fx}</filter>`);
  }
  return id;
}
/** The id of the page's filter for an inner shadow (offset and blur EMUs, CSS colour), defined in the page if new. */
export const innerShadowFilter = (dx, dy, blur, c) => pageFilter(innerShadowFx(dx, dy, blur, c, false), 'is');

/**
 * The CSS for a run's w14 effects (`items`: [localName, xml] of its unmodelled w:rPr children). Fill: the letters'
 * colour (none: hollow; a gradient: drawn through the letters); outline: a stroke of that width centred on the
 * letters' edges, as Word draws it; shadow: a copy offset by its distance and direction, blurred; glow: the colour
 * spread round the letters by its radius; reflection: custom properties for its mirrored copy (textspans.js puts the
 * copy on each word, editor.css draws it).
 */
export function textEffectsCss(items) {
  const css = [];
  const shadows = [];
  const w14 = (name) => items.find(([n, x]) => n === name && x.startsWith('<w14:'))?.[1];
  const fill = w14('textFill');
  if (fill) {
    if (/<w14:noFill\b/.test(fill)) css.push('-webkit-text-fill-color:transparent');
    else if (/<w14:gradFill\b/.test(fill)) {
      const stops = gradStops(fill).map((g) => `${w14Color(g[2]) || 'currentColor'} ${+(Number(g[1]) / 1000).toFixed(1)}%`);
      const ang = num(fill, 'ang');
      // (A linear gradient's angle is clockwise from left to right; CSS's from bottom to top.)
      const deg = Number.isFinite(ang) ? 90 + ang / 60000 : 90;
      const kind = /<w14:path\b/.test(fill) ? `radial-gradient(${stops.join(',')})` : `linear-gradient(${+deg.toFixed(2)}deg,${stops.join(',')})`;
      // (Also as a custom property, which inherits, for a reflection's copy: background-image doesn't.)
      if (stops.length) css.push(`background-image:${kind}`, `--odx-fill-bg:${kind}`, '-webkit-background-clip:text', 'background-clip:text', '-webkit-text-fill-color:transparent');
    } else {
      const c = w14Color(fill);
      if (c) css.push(`-webkit-text-fill-color:${c}`);
    }
  }
  const line = w14('textOutline');
  if (line && !/<w14:noFill\b/.test(line)) {
    const c = w14Color(line);
    const w = num(line, 'w');
    if (c) css.push(`-webkit-text-stroke:${Number.isFinite(w) ? pt(w) : 0.75}pt ${c}`);
  } else if (line) css.push('-webkit-text-stroke:0');
  const sh = w14('shadow');
  if (sh) {
    const c = w14Color(sh);
    const dist = num(sh, 'dist') || 0, dir = (num(sh, 'dir') || 0) / 60000, blur = num(sh, 'blurRad') || 0;
    // (A perspective shadow, scaled flat by sx/sy, Word draws as its own shape: drawn as the plain one here. One scaled to
    // nothing Word draws inside the letters, the letters looking pressed in, which text-shadow can't: an SVG filter on a
    // copy of each word laid over it, innerShadowFilter.)
    const a = (dir * Math.PI) / 180;
    if (c && num(sh, 'sx') === 0 && num(sh, 'sy') === 0) css.push(`--odx-is:url(#${innerShadowFilter(dist * Math.cos(a), dist * Math.sin(a), blur, c)})`);
    else if (c) shadows.push(`${pt(dist * Math.cos(a))}pt ${pt(dist * Math.sin(a))}pt ${pt(blur)}pt ${c}`);
  }
  const glow = w14('glow');
  if (glow) {
    const c = w14Color(glow);
    const r = num(glow, 'rad') || 0;
    // (Spread round the letters: copies at a third, two thirds and the whole radius, as Word's soft edge fades out.)
    if (c && r) for (const k of [1 / 3, 2 / 3, 1]) shadows.push(`0 0 ${pt(r * k)}pt ${c}`);
    // (Also as a custom property, for PDF export (pdf.js), which draws it from the letters' outlines; the colour as
    // rgb(), which it reads: custom properties aren't computed.)
    const rgbOf = (s) => (s.startsWith('#') ? `rgb(${[1, 3, 5].map((i) => parseInt(s.slice(i, i + 2), 16)).join(',')})` : s);
    if (c && r) css.push(`--odx-glow:${pt(r)}pt ${rgbOf(c)}`);
  }
  if (shadows.length) css.push(`text-shadow:${shadows.join(',')}`);
  // A bevel (w14:props3d): the run through the page's filter for it (bevelFx).
  const bevel = bevelFx(w14('props3d'));
  if (bevel) css.push(`filter:url(#${pageFilter(bevel, 'bv')})`);
  // The room Word gives them: a glow's line is taller above and below, a reflection's below (Word's PDFs of
  // mkeffects14c.ps1: 36pt Arial with glows of 2 to 30pt, the line 0.46 of the radius higher and 0.62 lower; at other
  // sizes in proportion, near enough: 12 to 72pt with a 10pt glow). A shadow takes none.
  let up = 0, down = 0;
  if (glow) {
    const r = pt(num(glow, 'rad') || 0);
    up = (0.46 * r) / 36;
    down = (0.62 * r) / 36;
  }
  const refl = w14('reflection');
  if (refl) {
    down = Math.max(down, reflectionRoom(num(refl, 'endPos') || 0, num(refl, 'dist') || 0));
    // The copy: from its start opacity at the top to its end one at its end position, a share of the font's ascent
    // (the part mirrored), and that opacity on below. Word's PDFs of effects-word/effects14b.docx, 36pt Arial bold:
    // Reflection 1 (endPos 45.5%) fades out 15.2pt below the baseline, 2 (58%) 19pt, then stays at its end opacity
    // to the mirrored letters' foot; 0.905em (Arial's ascent) times each. Its gap below the baseline is twice the
    // distance less the blur, near enough: 3.8pt for a 2.36pt distance and 0.5pt blur (Reflection 4), 8.25pt for
    // 4.72pt (Reflection 7), none for none. (Word's copy keeps the round letters' overshoot below the baseline,
    // which starts it a little higher; ours is cut at the baseline, so the gap is fitted to where each starts.) Its blur:
    // Word's print softens the copy's edges far less than a shadow of that blur, over about 0.3pt for the gallery's
    // 0.5pt (wordart-leftovers/refl.pdf), so a fifth of it is the spread.
    const a = (k) => +Math.min(1, Math.max(0, (num(refl, k) || 0) / 100000)).toFixed(4);
    const gap = Math.max(0, 2 * pt(num(refl, 'dist') || 0) - pt(num(refl, 'blurRad') || 0));
    const k = Math.max(0, (num(refl, 'endPos') || 0) / 100000);
    css.push(`--odx-rf-a:${a('stA')}`, `--odx-rf-b:${a('endA')}`, `--odx-rf-pos:${+(k * 100).toFixed(2)}%`, `--odx-rf-k:${+k.toFixed(4)}`, `--odx-rf-gap:${+gap.toFixed(2)}pt`);
    const blur = Math.min(100, pt(num(refl, 'blurRad') || 0)) / 5; // (bounded: a huge blur on every word crawls)
    if (blur > 0) css.push(`--odx-rf-blur:${+blur.toFixed(3)}pt`);
  }
  if (up) css.push(`--odx-grow-up:${+up.toFixed(4)}em`);
  if (down) css.push(`--odx-grow-down:${+down.toFixed(4)}em`);
  return css;
}

// Word's nine reflections (Text Effects > Reflection; its Font.Reflection.Type 1 to 9) as it writes them: end position
// and distance (EMUs), and the room below the line it gives each, in ems (its PDFs at 18, 36 and 72pt, in proportion
// to the size: the distance's share too).
const REFLECTIONS = [
  [45500, 0, 0.257], [58000, 0, 0.848], [85000, 0, 0.684],
  [50000, 29997, 0.462], [60000, 29997, 1.003], [85000, 29997, 0.845],
  [50000, 60007, 0.617], [60000, 60007, 1.161], [85000, 60007, 0.995],
];
/** The room below the line (em) Word gives a reflection: the nearest of its own nine. */
function reflectionRoom(endPos, dist) {
  let best = REFLECTIONS[0], d0 = Infinity;
  for (const r of REFLECTIONS) {
    const d = Math.abs(r[0] - endPos) / 1000 + Math.abs(r[1] - dist) / 1000;
    if (d < d0) { d0 = d; best = r; }
  }
  return best[2];
}

/** Whether a run's items have any w14 effect (which then stands instead of the old w:outline/w:shadow). */
export const hasTextEffects = (items) => items.some(([n, x]) => ['textFill', 'textOutline', 'shadow', 'glow', 'reflection', 'props3d'].includes(n) && x.startsWith('<w14:'));

/**
 * A text box run's w14 effects for its SVG text (shapes.js): `paint`, the tspan's fill and stroke attributes (or null
 * to keep its own colour), and `filter` and `defs` for the text element (a shadow, a glow; a gradient fill's
 * definition). `rPr` is the run's w:rPr element; `n` makes the definitions' ids unique in the drawing.
 */
export function svgTextEffects(rPr, n, serialize) {
  const items = [];
  for (const c of rPr ? [...rPr.childNodes].filter((k) => k.nodeType === 1) : []) {
    if (['textFill', 'textOutline', 'shadow', 'glow', 'reflection', 'props3d'].includes(c.localName)) {
      const x = serialize(c);
      if (x.startsWith('<w14:')) items.push([c.localName, x]);
    }
  }
  if (!items.length) return null;
  const get = (name) => items.find(([k]) => k === name)?.[1];
  const attrs = [];
  const defs = [];
  const fill = get('textFill');
  if (fill) {
    if (/<w14:noFill\b/.test(fill)) attrs.push('fill="none"');
    else if (/<w14:gradFill\b/.test(fill)) {
      const stops = gradStops(fill);
      const ang = num(fill, 'ang');
      const a = ((Number.isFinite(ang) ? ang / 60000 : 0) * Math.PI) / 180;
      const [x1, y1, x2, y2] = [0.5 - Math.cos(a) / 2, 0.5 - Math.sin(a) / 2, 0.5 + Math.cos(a) / 2, 0.5 + Math.sin(a) / 2].map((v) => +v.toFixed(3));
      defs.push(`<linearGradient id="odx-tf${n}" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">${stops.map((g) => {
        const c = w14Color(g[2]) || '#000';
        return `<stop offset="${+(Number(g[1]) / 100000).toFixed(3)}" stop-color=${solid(c)}stop-opacity="${opacity(c)}"/>`;
      }).join('')}</linearGradient>`);
      attrs.push(`fill="url(#odx-tf${n})"`);
    } else {
      const c = w14Color(fill);
      if (c) attrs.push(`fill=${solid(c).trim()}`, ...(opacity(c) < 1 ? [`fill-opacity="${opacity(c)}"`] : []));
    }
  }
  const line = get('textOutline');
  if (line && !/<w14:noFill\b/.test(line)) {
    const c = w14Color(line);
    const w = num(line, 'w');
    if (c) attrs.push(`stroke=${solid(c).trim()}`, `stroke-width="${Number.isFinite(w) ? px(w) : 1}"`, ...(opacity(c) < 1 ? [`stroke-opacity="${opacity(c)}"`] : []));
  }
  const fx = [];
  // A bevel first, on the letters; the shadow and glow then take the bevelled letters for theirs.
  const bevel = bevelFx(get('props3d'), 'odxbev', true);
  if (bevel) fx.push(bevel);
  const src = bevel ? 'odxbev' : 'SourceGraphic';
  const sh = get('shadow');
  if (sh) {
    const c = w14Color(sh);
    const dist = num(sh, 'dist') || 0, a = (((num(sh, 'dir') || 0) / 60000) * Math.PI) / 180, blur = num(sh, 'blurRad') || 0;
    // (Scaled to nothing, sx and sy 0, Word draws it inside the letters: an inner shadow, the letters looking pressed in.)
    const inner = num(sh, 'sx') === 0 && num(sh, 'sy') === 0;
    if (c && inner) fx.push(innerShadowFx(dist * Math.cos(a), dist * Math.sin(a), blur, c, true).replace(/"SourceGraphic"/g, `"${src}"`));
    else if (c) fx.push(`<feDropShadow dx="${px(dist * Math.cos(a))}" dy="${px(dist * Math.sin(a))}" stdDeviation="${px(blur / 2)}" flood-color=${solid(c)}flood-opacity="${opacity(c)}"/>`);
  }
  const glow = get('glow');
  if (glow) {
    const c = w14Color(glow);
    const r = px(num(glow, 'rad') || 0);
    if (c && r) {
      fx.push(`<feMorphology in="SourceAlpha" operator="dilate" radius="${+(r / 2).toFixed(2)}" result="g1"/><feGaussianBlur in="g1" stdDeviation="${+(r / 3).toFixed(2)}" result="g2"/>`
        + `<feFlood flood-color=${solid(c)}flood-opacity="${opacity(c)}"/><feComposite in2="g2" operator="in" result="glow"/><feMerge><feMergeNode in="glow"/><feMergeNode in="${src}"/></feMerge>`);
    }
  }
  if (fx.length) defs.push(`<filter id="odx-tx${n}" x="-50%" y="-50%" width="200%" height="200%"${bevel ? ' color-interpolation-filters="sRGB"' : ''}>${fx.join('')}</filter>`);
  // (A reflection: its start and end opacities, how far down it fades out (a share of the text's height), and its
  // distance below the text, px.)
  const rf = get('reflection');
  const reflection = rf ? { stA: Math.min(1, (num(rf, 'stA') || 0) / 100000), endA: Math.min(1, (num(rf, 'endA') || 0) / 100000), endPos: (num(rf, 'endPos') || 0) / 100000, dist: px(num(rf, 'dist') || 0), blur: Math.min(133, px(num(rf, 'blurRad') || 0)) } : null;
  return { paint: attrs.length ? attrs.join(' ') : null, filter: fx.length ? `url(#odx-tx${n})` : null, defs: defs.join(''), reflection };
}
