// DrawingML shapes and groups (wps:wsp, wpg:wgp, pic:pic inside them) drawn as one SVG:
// preset and custom geometry, solid fills and outlines, pictures, and text boxes.
// Display only; the drawing's XML is saved back untouched.
import { esc, elementKids as els, childNamed as child } from './xml.js';
import { lineFactor } from './props.js';
import { geometryPaths, readGeometry, adjustOf } from './geometry.js';
import { PRESETS } from './presets.js';
import { EMU_PER_PX, TWIPS_PER_PX, emuToPx, twipsToPx } from './units.js';

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
function modify(hex, kind, v) {
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
  if (f) fill = f.localName === 'solidFill' ? solid(f) : f.localName === 'gradFill' ? color(desc(f, 'gs') && els(desc(f, 'gs'))[0], colors) : null;
  else if (style && child(style, 'fillRef') && child(style, 'fillRef').getAttribute('idx') !== '0') fill = color(els(child(style, 'fillRef'))[0], colors);
  const ln = child(spPr, 'ln');
  if (ln) {
    if (ln.getAttribute('w')) width = px(ln.getAttribute('w'));
    stroke = child(ln, 'noFill') ? null : solid(child(ln, 'solidFill')) || (style && child(style, 'lnRef') ? color(els(child(style, 'lnRef'))[0], colors) : null);
  } else if (lnIdx) stroke = color(els(child(style, 'lnRef'))[0], colors);
  return { fill, stroke, width, ln };
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
/** Width of `text` in px: measured in a browser, else estimated at half an em per character. */
function textWidth(text, sizePx, bold, font) {
  if (canvas == null) canvas = typeof document !== 'undefined' && document.createElement('canvas').getContext?.('2d') || false;
  if (!canvas) return text.length * sizePx * 0.5;
  canvas.fontKerning = 'none';
  canvas.font = `${bold ? 'bold ' : ''}${sizePx}px "${font}", sans-serif`;
  return canvas.measureText(text).width;
}

/**
 * Text box paragraphs as SVG text: indents, tab stops, spacing and exact line heights from the
 * paragraph, size, colour and weight per run. Lines wrap at the box's inner width (unless the box
 * says wrap="none"), word by word as in Word; a tab that would pass the right edge starts a new line.
 */
function textBox(txbx, bodyPr, w, h, ctx, line = 0) {
  const font0 = ctx.font;
  // The insets are from the inside of the box's outline (sample S078's cover box: its text 1px lower in Word's
  // PDF than from the shape's edge, under a 0.75pt line).
  const ins = (k, d) => (bodyPr && bodyPr.getAttribute(k) != null ? px(bodyPr.getAttribute(k)) : d) + line;
  const l = ins('lIns', 9.6);
  const t = ins('tIns', 4.8);
  const r = ins('rIns', 9.6);
  const wrap = bodyPr?.getAttribute('wrap') !== 'none';
  const out = [];
  let y = t;
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
    const indL = wa(ind, 'left') ?? wa(ind, 'start') ?? st.pPr.indLeft ?? 0;
    const indR = wa(ind, 'right') ?? wa(ind, 'end') ?? st.pPr.indRight ?? 0;
    const indF = wa(ind, 'firstLine') != null || wa(ind, 'hanging') != null ? (Number(wa(ind, 'firstLine')) || 0) - (Number(wa(ind, 'hanging')) || 0) : st.pPr.indFirst ?? 0;
    const left = l + twipsToPx(indL);
    const right = R - twipsToPx(indR);
    const first = left + twipsToPx(indF);
    // Lines of placed pieces: { x, ink (where its last visible glyph ends), text, open (its tspan's start tag) }.
    const lines = [{ pieces: [], x: first }];
    const newLine = () => lines.push({ pieces: [], x: left });
    let maxSize = 0;
    let lineF = lineFactor(st.font);
    for (const run of [...p.getElementsByTagName('*')].filter((n) => n.localName === 'r')) {
      const rPr = child(run, 'rPr');
      // Its character style (w:rStyle) between the paragraph style and its own formatting.
      const cs = { ...st.rPr, ...(ctx.char ? ctx.char(wa(child(rPr, 'rStyle'), 'val')) : {}) };
      const size = Number(wa(child(rPr, 'sz'), 'val')) / 2 || (cs.size ? cs.size / 2 : markSz);
      const sizePx = (size * 4) / 3;
      const c = wa(child(rPr, 'color'), 'val') ?? cs.color;
      const on = (name, key) => { const e = child(rPr, name); return e ? !/^(0|false)$/.test(wa(e, 'val') || '') : !!cs[key]; };
      const bold = on('b', 'bold');
      const italic = on('i', 'italic');
      const letter = (Number(wa(child(rPr, 'spacing'), 'val')) || 0) / TWIPS_PER_PX;
      const rf = child(rPr, 'rFonts');
      const theme = wa(rf, 'asciiTheme');
      const font = wa(rf, 'ascii') || wa(rf, 'hAnsi') || (theme && ctx.theme?.[theme.startsWith('major') ? 'major' : 'minor']) || cs.font || st.font || font0;
      const open = `<tspan font-family="${esc(font)}, Calibri, Arial, sans-serif" font-size="${sizePx.toFixed(2)}" fill="#${c && c !== 'auto' ? c : ctx.textColor || '000000'}"${bold ? ' font-weight="bold"' : ''}${italic ? ' font-style="italic"' : ''}${letter ? ` letter-spacing="${letter.toFixed(2)}"` : ''}>`;
      for (const n of els(run)) {
        const line = lines[lines.length - 1];
        if (n.localName === 'tab') {
          const stop = tabs.find((v) => v > line.x - l + 0.5);
          const to = l + (stop ?? (Math.floor((line.x - l) / 48) + 1) * 48);
          if (wrap && to >= right && line.pieces.length) newLine();
          else line.x = to;
        } else if (n.localName === 'br' || n.localName === 'cr') {
          newLine();
        } else if (n.localName === 't' && n.textContent) {
          if (sizePx * lineFactor(font) > maxSize * lineF) lineF = lineFactor(font);
          maxSize = Math.max(maxSize, sizePx);
          // Words keep their trailing spaces; a word that doesn't fit starts the next line.
          for (const word of n.textContent.match(/\S*\s*/g).filter(Boolean)) {
            let cur = lines[lines.length - 1];
            const width = textWidth(word, sizePx, bold, font) + word.length * letter;
            const inked = textWidth(word.trimEnd(), sizePx, bold, font) + word.trimEnd().length * letter;
            if (wrap && cur.x + inked > right + 0.5 && cur.pieces.length) { newLine(); cur = lines[lines.length - 1]; }
            const last = cur.pieces[cur.pieces.length - 1];
            cur.one = Math.max(cur.one || 0, sizePx * lineFactor(font)); // (each line as tall as its own text: Word's)
            if (last && last.open === open && last.end === cur.x) last.text += word;
            else cur.pieces.push({ x: cur.x, text: word, open });
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
    const before = (Number(sp('before')) || 0) / TWIPS_PER_PX;
    const ctx2 = wa(child(pPr, 'contextualSpacing'), 'val');
    const contextual = child(pPr, 'contextualSpacing') ? !/^(0|false|off)$/.test(ctx2 || '') : !!st.pPr.contextualSpacing;
    if (contextual && prevStyle === (sid || '')) { y -= prevAfter; }
    else y += Math.max(0, before - prevAfter);
    for (const ln of lines) {
      const one = ln.one || sizePx * lineF;
      const lineH = heightOf(one);
      // Baseline: the font's ascent (0.778 of its line: 0.95em of Calibri's 1.22) below the line top, centred in an exact
      // line's extra room; a multiple line's extra goes below its text, as in the body (Word's PDF of a text box at
      // 1.08 lines: its first baseline 0.7px higher than centred).
      const base = y + (rule === 'exact' ? (lineH - one) / 2 : rule === 'atLeast' ? lineH - one : 0) + one * 0.778;
      y += lineH;
      if (!ln.pieces.length) continue;
      // Centred and right-aligned lines move as a whole (trailing spaces don't count).
      const start = ln === lines[0] ? first : left;
      const ink = ln.pieces[ln.pieces.length - 1].ink;
      const shift = align === 'center' ? (right - ink) / 2 - (ln.pieces[0].x - start) / 2 : align === 'right' || align === 'end' ? right - ink : 0;
      for (const pc of ln.pieces) {
        out.push(`<text x="${(pc.x + shift).toFixed(1)}" y="${base.toFixed(1)}" xml:space="preserve">${pc.open}${esc(pc.text)}</tspan></text>`);
      }
    }
    const after = (Number(sp('after')) || 0) / TWIPS_PER_PX;
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
    rows.forEach((tr, j) => {
      const top = y;
      rowTops.push(top);
      let bottom = top;
      for (const { tc, c, span, merge } of cellsOf[j]) {
        if (merge === 'cont') continue; // (its text is the cell above's)
        const x0 = xs[Math.min(c, grid.length)], x1 = xs[Math.min(c + span, grid.length)];
        y = top + mT;
        prevAfter = 0; prevStyle = null;
        layBlocks(tc, x0 + mL, x1 - mR, ts || {});
        const k = merge === 'restart' ? last(j, c) : j;
        if (k > j) owed.set(k, Math.max(owed.get(k) ?? 0, y + mB));
        else bottom = Math.max(bottom, y + mB);
      }
      y = Math.max(bottom, owed.get(j) ?? 0);
    });
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
  ctx.textEnd = y + ins('bIns', 4.8); // (where the text ends, for a box that grows to fit it: shapeSvg)
  return out.join('');
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
      cur.pieces.push({ x: cur.x, text: bu.getAttribute('char') || '\u2022', open: `<tspan font-family="${esc(buFont)}, Arial, sans-serif" font-size="${sizePx.toFixed(2)}" fill="#${ctx.textColor || '000000'}">`, ink: cur.x + sizePx * 0.5 });
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
      const open = `<tspan font-family="${esc(font)}, Calibri, Arial, sans-serif" font-size="${sizePx.toFixed(2)}" fill="#${c || ctx.textColor || '000000'}"${bold ? ' font-weight="bold"' : ''}${italic ? ' font-style="italic"' : ''}>`;
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

function shapeSvg(el, ctx) {
  const name = el.localName;
  if (name === 'wgp' || name === 'grpSp' || name === 'spTree') {
    const xfrm = child(child(el, 'grpSpPr'), 'xfrm');
    const inner = els(el).filter((n) => ['wsp', 'sp', 'pic', 'grpSp'].includes(n.localName)).map((n) => shapeSvg(n, ctx)).join('');
    if (!xfrm) return inner;
    const [ext, chOff, chExt] = ['ext', 'chOff', 'chExt'].map((k) => child(xfrm, k));
    const off = name === 'wgp' ? null : child(xfrm, 'off'); // the outer group's frame is the drawing itself
    const sx = chExt && Number(chExt.getAttribute('cx')) ? Number(ext.getAttribute('cx')) / Number(chExt.getAttribute('cx')) : 1;
    const sy = chExt && Number(chExt.getAttribute('cy')) ? Number(ext.getAttribute('cy')) / Number(chExt.getAttribute('cy')) : 1;
    const tx = px(off?.getAttribute('x')) - px(chOff?.getAttribute('x')) * sx;
    const ty = px(off?.getAttribute('y')) - px(chOff?.getAttribute('y')) * sy;
    return `<g transform="matrix(${sx} 0 0 ${sy} ${tx.toFixed(2)} ${ty.toFixed(2)})">${inner}</g>`;
  }
  const spPr = child(el, 'spPr');
  const xfrm = child(spPr, 'xfrm');
  if (!xfrm) return '';
  const off = child(xfrm, 'off');
  const ext = child(xfrm, 'ext');
  const x = px(off?.getAttribute('x'));
  const y = px(off?.getAttribute('y'));
  const w = px(ext?.getAttribute('cx'));
  const h = px(ext?.getAttribute('cy'));
  const rot = Number(xfrm.getAttribute('rot') || 0) / 60000;
  const flip = `${xfrm.getAttribute('flipH') === '1' ? -1 : 1} 0 0 ${xfrm.getAttribute('flipV') === '1' ? -1 : 1}`;
  const t = `translate(${x.toFixed(2)} ${y.toFixed(2)})${rot ? ` rotate(${rot} ${w / 2} ${h / 2})` : ''}${flip !== '1 0 0 1' ? ` matrix(${flip} ${xfrm.getAttribute('flipH') === '1' ? w : 0} ${xfrm.getAttribute('flipV') === '1' ? h : 0})` : ''}`;
  if (name === 'pic') {
    const blip = desc(el, 'blip');
    const href = blip && ctx.image(blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed'));
    return href ? `<image transform="${t}" width="${w.toFixed(2)}" height="${h.toFixed(2)}" preserveAspectRatio="none" href="${href}"/>` : '';
  }
  if (name !== 'wsp' && name !== 'sp') return '';
  const style = child(el, 'style');
  const pt = paints(spPr, style, ctx.colors);
  const txbx = child(el, 'txbx');
  const ln = child(spPr, 'ln');
  const line = ln && !child(ln, 'noFill') && child(ln, 'solidFill') ? px(ln.getAttribute('w') || EMU_PER_PX) : 0;
  // (Text without a colour of its own takes the shape's font colour, wps:style's fontRef: white on Word's default shapes.)
  const fontRef = style && child(style, 'fontRef');
  const textColor = fontRef && els(fontRef)[0] ? color(els(fontRef)[0], ctx.colors)?.hex : null;
  const tctx = textColor ? { ...ctx, textColor } : ctx;
  const bodyPr = child(el, 'bodyPr');
  const text = txbx ? textBox(txbx, bodyPr, w, h, tctx, line) : '';
  // "Resize shape to fit text" (a:spAutoFit): Word makes the box as tall as its text when it lays the page out, whatever
  // height the file gives it (a text box of four paragraphs in 4.2in: Word's body text wraps beside it to its end).
  let hh = h;
  if (txbx && child(bodyPr, 'spAutoFit') && tctx.textEnd > h + 0.5) { hh = tctx.textEnd; ctx.grown = Math.max(ctx.grown || 0, hh - h); }
  const paths = geometry(spPr, w, hh);
  const strokeAttrs = pt.stroke ? ` stroke-width="${+pt.width.toFixed(2)}" stroke-miterlimit="8"` : '';
  let body = paths.map((q) => `<path d="${q.d}" ${paint('fill', modeFill(pt.fill, q.fill))} ${q.stroke ? paint('stroke', pt.stroke) + strokeAttrs : 'stroke="none"'} vector-effect="non-scaling-stroke"/>`).join('');
  // Arrowheads, where the outline has them (on a line, or any open path).
  const open = paths.find((q) => q.ends && !/Z$/.test(q.d));
  if (pt.stroke && open && pt.ln) body += arrowhead(child(pt.ln, 'headEnd'), open.ends[0], pt.width, pt.stroke) + arrowhead(child(pt.ln, 'tailEnd'), open.ends[1], pt.width, pt.stroke);
  // DrawingML text (a:txBody: SmartArt's shapes), in its own rectangle (dsp:txXfrm) where it has one, unrotated.
  const txBody = child(el, 'txBody');
  let dml = '';
  if (txBody) {
    const tx = child(child(el, 'txXfrm'), 'off') ? child(el, 'txXfrm') : null;
    const box = tx ? { x: px(child(tx, 'off').getAttribute('x')) - x, y: px(child(tx, 'off').getAttribute('y')) - y, w: px(child(tx, 'ext').getAttribute('cx')), h: px(child(tx, 'ext').getAttribute('cy')) } : { x: 0, y: 0, w, h };
    dml = `<g transform="translate(${x.toFixed(2)} ${y.toFixed(2)})">${dmlText(txBody, box, tctx)}</g>`;
  }
  return `<g transform="${t}">${body}${text}</g>${dml}`;
}

/**
 * SVG for a drawing whose graphic is a shape or a group of shapes, sized to its extent.
 * @param {Element} drawing w:drawing
 * @param {{ image: (rId: string) => string|null, colors?: object, font?: string }} ctx image() returns a data: URL
 */
export function shapesSvg(drawing, ctx, width, height, pad = 0) {
  const data = desc(drawing, 'graphicData');
  const root = data && els(data).find((n) => n.localName === 'wgp' || n.localName === 'wsp');
  if (!root || !width || !height) return null;
  const c = { colors: {}, font: 'Calibri', ...ctx };
  const inner = shapeSvg(root, c);
  if (!inner) return null;
  // A single text box that grows to fit its text (shapeSvg) is that much taller; ctx.grown tells the caller.
  if (root.localName === 'wsp' && c.grown) { height = Math.round((height + c.grown) * 100) / 100; ctx.grownTo = height; }
  const W = width + 2 * pad, H = height + 2 * pad;
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
 * How far a drawing's lines are drawn past its box (px) for their arrowheads, which Word draws there (an image can't:
 * the room is added around it, shapesSvg's pad). Only for arrowheads: room for every outline moved lines on the page
 * by a fraction of a pixel (sample S129, 103 inline shapes: 118 lines 0.7px off), so outlines stay inside the box.
 */
export function shapesPad(drawing) {
  let pad = 0;
  for (const ln of [...drawing.getElementsByTagName('*')].filter((n) => n.localName === 'ln' && n.parentNode?.localName === 'spPr')) {
    if (child(ln, 'noFill')) continue;
    const w = ln.getAttribute('w') ? px(ln.getAttribute('w')) : 1.33;
    const ends = ['headEnd', 'tailEnd'].map((k) => child(ln, k)).filter((e) => e && e.getAttribute('type') && e.getAttribute('type') !== 'none');
    if (!ends.length) continue;
    const k = { sm: 2, med: 3, lg: 5 };
    pad = Math.max(pad, w / 2 + Math.max(...ends.map((e) => Math.max(w, 8 / 3) * Math.max(k[e.getAttribute('w')] || 3, k[e.getAttribute('len')] || 3))));
  }
  return Math.ceil(pad);
}
