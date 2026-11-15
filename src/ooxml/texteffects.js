// Word 2010's text effects on a run (w14:textFill, w14:textOutline, w14:shadow, w14:glow, w14:reflection: what Home >
// Text Effects and WordArt write), as CSS. They're kept as the run's raw XML (reader.js rPrMarks), so this only draws
// them; saving writes them back as they were.
//
// Checked against Word (odx-test/mkeffects14.ps1: Word itself setting each effect through its Font object): a w14
// colour's w14:alpha is its transparency (Word's 50% transparent red is alpha 50000), not DrawingML's opacity; Word
// also writes the old w:outline beside a w14:textOutline, which then is the outline drawn.
import { modify, OFFICE } from './shapes.js';

// The open document's theme colours (docx.js, from its theme part): scheme name -> hex.
const THEME = {};
export function setThemeColors(colors) {
  for (const k of Object.keys(THEME)) delete THEME[k];
  Object.assign(THEME, colors || {});
}

const SCHEME = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };
const num = (xml, name) => Number(new RegExp(`\\bw14:${name}="(-?\\d+)"`).exec(xml)?.[1] ?? NaN);

/** The first colour in a w14 fragment (srgbClr or schemeClr, with its changes) as a CSS colour; null if none. */
export function w14Color(xml) {
  const m = /<w14:(srgbClr|schemeClr)\b[^>]*?\bw14:val="([^"]+)"[^>]*?(\/>|>([\s\S]*?)<\/w14:\1>)/.exec(xml || '');
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

// (EMUs to points: 12700 a point.)
const pt = (emu) => +(emu / 12700).toFixed(2);

/**
 * The CSS for a run's w14 effects (`items`: [localName, xml] of its unmodelled w:rPr children). Fill: the letters'
 * colour (none: hollow; a gradient: drawn through the letters); outline: a stroke of that width centred on the
 * letters' edges, as Word draws it; shadow: a copy offset by its distance and direction, blurred; glow: the colour
 * spread round the letters by its radius. (A reflection isn't drawn.)
 */
export function textEffectsCss(items) {
  const css = [];
  const shadows = [];
  const w14 = (name) => items.find(([n, x]) => n === name && x.startsWith('<w14:'))?.[1];
  const fill = w14('textFill');
  if (fill) {
    if (/<w14:noFill\b/.test(fill)) css.push('-webkit-text-fill-color:transparent');
    else if (/<w14:gradFill\b/.test(fill)) {
      const stops = [...fill.matchAll(/<w14:gs\b[^>]*?\bw14:pos="(\d+)"[^>]*>([\s\S]*?)<\/w14:gs>/g)].map((g) => `${w14Color(g[2]) || 'currentColor'} ${+(Number(g[1]) / 1000).toFixed(1)}%`);
      const ang = num(fill, 'ang');
      // (A linear gradient's angle is clockwise from left to right; CSS's from bottom to top.)
      const deg = Number.isFinite(ang) ? 90 + ang / 60000 : 90;
      const kind = /<w14:path\b/.test(fill) ? `radial-gradient(${stops.join(',')})` : `linear-gradient(${+deg.toFixed(2)}deg,${stops.join(',')})`;
      if (stops.length) css.push(`background-image:${kind}`, '-webkit-background-clip:text', 'background-clip:text', '-webkit-text-fill-color:transparent');
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
    // nothing Word draws inside the letters, which CSS can't: not drawn.)
    const a = (dir * Math.PI) / 180;
    if (c && !(num(sh, 'sx') === 0 && num(sh, 'sy') === 0)) shadows.push(`${pt(dist * Math.cos(a))}pt ${pt(dist * Math.sin(a))}pt ${pt(blur)}pt ${c}`);
  }
  const glow = w14('glow');
  if (glow) {
    const c = w14Color(glow);
    const r = num(glow, 'rad') || 0;
    // (Spread round the letters: copies at a third, two thirds and the whole radius, as Word's soft edge fades out.)
    if (c && r) for (const k of [1 / 3, 2 / 3, 1]) shadows.push(`0 0 ${pt(r * k)}pt ${c}`);
  }
  if (shadows.length) css.push(`text-shadow:${shadows.join(',')}`);
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
  if (refl) down = Math.max(down, reflectionRoom(num(refl, 'endPos') || 0, num(refl, 'dist') || 0));
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
export const hasTextEffects = (items) => items.some(([n, x]) => ['textFill', 'textOutline', 'shadow', 'glow', 'reflection'].includes(n) && x.startsWith('<w14:'));

/**
 * A text box run's w14 effects for its SVG text (shapes.js): `paint`, the tspan's fill and stroke attributes (or null
 * to keep its own colour), and `filter` and `defs` for the text element (a shadow, a glow; a gradient fill's
 * definition). `rPr` is the run's w:rPr element; `n` makes the definitions' ids unique in the drawing.
 */
export function svgTextEffects(rPr, n, serialize) {
  const items = [];
  for (const c of rPr ? [...rPr.childNodes].filter((k) => k.nodeType === 1) : []) {
    if (['textFill', 'textOutline', 'shadow', 'glow', 'reflection'].includes(c.localName)) {
      const x = serialize(c);
      if (x.startsWith('<w14:')) items.push([c.localName, x]);
    }
  }
  if (!items.length) return null;
  const get = (name) => items.find(([k]) => k === name)?.[1];
  const px = (emu) => +(emu / 9525).toFixed(2);
  const attrs = [];
  const defs = [];
  const solid = (c) => {
    const m = /^rgba\((\d+),(\d+),(\d+),([\d.]+)\)$/.exec(c);
    return m ? `"#${[m[1], m[2], m[3]].map((v) => Number(v).toString(16).padStart(2, '0')).join('')}" ` : `"${c}" `;
  };
  const opacity = (c) => /^rgba\(/.exec(c) ? Number(/,([\d.]+)\)$/.exec(c)[1]) : 1;
  const fill = get('textFill');
  if (fill) {
    if (/<w14:noFill\b/.test(fill)) attrs.push('fill="none"');
    else if (/<w14:gradFill\b/.test(fill)) {
      const stops = [...fill.matchAll(/<w14:gs\b[^>]*?\bw14:pos="(\d+)"[^>]*>([\s\S]*?)<\/w14:gs>/g)];
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
  const sh = get('shadow');
  if (sh) {
    const c = w14Color(sh);
    const dist = num(sh, 'dist') || 0, a = (((num(sh, 'dir') || 0) / 60000) * Math.PI) / 180, blur = num(sh, 'blurRad') || 0;
    // (Scaled to nothing, sx and sy 0, Word draws it inside the letters: an inner shadow, the letters looking pressed in.)
    const inner = num(sh, 'sx') === 0 && num(sh, 'sy') === 0;
    if (c && inner) {
      fx.push(`<feComponentTransfer in="SourceAlpha" result="i0"><feFuncA type="table" tableValues="1 0"/></feComponentTransfer><feOffset in="i0" dx="${px(dist * Math.cos(a))}" dy="${px(dist * Math.sin(a))}" result="i1"/>`
        + `<feGaussianBlur in="i1" stdDeviation="${px(blur / 2)}" result="i2"/><feFlood flood-color=${solid(c)}flood-opacity="${opacity(c)}"/><feComposite in2="i2" operator="in" result="i3"/>`
        + '<feComposite in="i3" in2="SourceAlpha" operator="in" result="i4"/><feMerge><feMergeNode in="SourceGraphic"/><feMergeNode in="i4"/></feMerge>');
    } else if (c) fx.push(`<feDropShadow dx="${px(dist * Math.cos(a))}" dy="${px(dist * Math.sin(a))}" stdDeviation="${px(blur / 2)}" flood-color=${solid(c)}flood-opacity="${opacity(c)}"/>`);
  }
  const glow = get('glow');
  if (glow) {
    const c = w14Color(glow);
    const r = px(num(glow, 'rad') || 0);
    if (c && r) {
      fx.push(`<feMorphology in="SourceAlpha" operator="dilate" radius="${+(r / 2).toFixed(2)}" result="g1"/><feGaussianBlur in="g1" stdDeviation="${+(r / 3).toFixed(2)}" result="g2"/>`
        + `<feFlood flood-color=${solid(c)}flood-opacity="${opacity(c)}"/><feComposite in2="g2" operator="in" result="glow"/><feMerge><feMergeNode in="glow"/><feMergeNode in="SourceGraphic"/></feMerge>`);
    }
  }
  if (fx.length) defs.push(`<filter id="odx-tx${n}" x="-50%" y="-50%" width="200%" height="200%">${fx.join('')}</filter>`);
  // (A reflection: its start and end opacities, how far down it fades out (a share of the text's height), and its
  // distance below the text, px.)
  const rf = get('reflection');
  const reflection = rf ? { stA: Math.min(1, (num(rf, 'stA') || 0) / 100000), endA: Math.min(1, (num(rf, 'endA') || 0) / 100000), endPos: (num(rf, 'endPos') || 0) / 100000, dist: px(num(rf, 'dist') || 0) } : null;
  return { paint: attrs.length ? attrs.join(' ') : null, filter: fx.length ? `url(#odx-tx${n})` : null, defs: defs.join(''), reflection };
}
