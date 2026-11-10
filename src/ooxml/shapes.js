// DrawingML shapes and groups (wps:wsp, wpg:wgp, pic:pic inside them) drawn as one SVG:
// preset and custom geometry, solid fills and outlines, pictures, and text boxes.
// Display only; the drawing's XML is saved back untouched.
import { esc, elementKids as els, childNamed as child } from './xml.js';
import { lineFactor } from './props.js';
import { EMU_PER_PX, TWIPS_PER_PX, emuToPx, twipsToPx } from './units.js';

/** Office's default theme colours, for a drawing whose theme doesn't say. */
export const OFFICE = { accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47', dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6' };
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
    else if (m.localName === 'lumMod' || m.localName === 'shade') hex = mix(hex, v, '000000');
    else if (m.localName === 'lumOff' || m.localName === 'tint') hex = mix(hex, m.localName === 'tint' ? v : 1 - v, 'FFFFFF');
  }
  return { hex, alpha };
}

function mix(hex, keep, other) {
  const a = parseInt(hex, 16);
  const b = parseInt(other, 16);
  return [16, 8, 0].map((s) => Math.round(((a >> s) & 255) * keep + ((b >> s) & 255) * (1 - keep)).toString(16).padStart(2, '0')).join('');
}

/** fill="..." stroke="..." attributes from spPr (and the shape's style references as a fallback). */
function paints(spPr, style, colors) {
  const solid = (el) => el && color(els(el)[0], colors);
  let fill = null;
  let stroke = null;
  let width = 0.75 * (4 / 3);
  const f = els(spPr).find((n) => /Fill$/.test(n.localName));
  if (f) fill = f.localName === 'solidFill' ? solid(f) : f.localName === 'gradFill' ? color(desc(f, 'gs') && els(desc(f, 'gs'))[0], colors) : null;
  else if (style && child(style, 'fillRef') && child(style, 'fillRef').getAttribute('idx') !== '0') fill = color(els(child(style, 'fillRef'))[0], colors);
  const ln = child(spPr, 'ln');
  if (ln) {
    if (ln.getAttribute('w')) width = px(ln.getAttribute('w'));
    stroke = child(ln, 'noFill') ? null : solid(child(ln, 'solidFill')) || (style && child(style, 'lnRef') ? color(els(child(style, 'lnRef'))[0], colors) : null);
  } else if (style && child(style, 'lnRef') && child(style, 'lnRef').getAttribute('idx') !== '0') stroke = color(els(child(style, 'lnRef'))[0], colors);
  const attrs = (kind, c) => (c ? `${kind}="#${c.hex}"${c.alpha < 1 ? ` ${kind}-opacity="${+c.alpha.toFixed(3)}"` : ''}` : `${kind}="none"`);
  return `${attrs('fill', fill)} ${attrs('stroke', stroke)}${stroke ? ` stroke-width="${+width.toFixed(2)}"` : ''}`;
}

/** SVG path data for a shape's geometry, in a w x h box. */
function geometry(spPr, w, h) {
  const cust = child(spPr, 'custGeom');
  if (cust) {
    let d = '';
    for (const path of els(child(cust, 'pathLst'))) {
      const pw = Number(path.getAttribute('w')) || w * EMU_PER_PX;
      const ph = Number(path.getAttribute('h')) || h * EMU_PER_PX;
      const sx = w / pw;
      const sy = h / ph;
      const pt = (p) => `${(Number(p.getAttribute('x')) * sx).toFixed(2)} ${(Number(p.getAttribute('y')) * sy).toFixed(2)}`;
      for (const c of els(path)) {
        const ps = els(c).filter((n) => n.localName === 'pt');
        if (c.localName === 'moveTo') d += `M${pt(ps[0])}`;
        else if (c.localName === 'lnTo') d += `L${pt(ps[0])}`;
        else if (c.localName === 'cubicBezTo') d += `C${ps.map(pt).join(' ')}`;
        else if (c.localName === 'quadBezTo') d += `Q${ps.map(pt).join(' ')}`;
        else if (c.localName === 'close') d += 'Z';
      }
    }
    return d;
  }
  const prst = child(spPr, 'prstGeom')?.getAttribute('prst') || 'rect';
  if (prst === 'ellipse') return `M0 ${h / 2}A${w / 2} ${h / 2} 0 1 0 ${w} ${h / 2}A${w / 2} ${h / 2} 0 1 0 0 ${h / 2}Z`;
  if (prst === 'line' || /Connector/.test(prst)) return `M0 0L${w} ${h}`;
  if (prst === 'roundRect') {
    const r = Math.min(w, h) * 0.1667;
    return `M${r} 0H${w - r}Q${w} 0 ${w} ${r}V${h - r}Q${w} ${h} ${w - r} ${h}H${r}Q0 ${h} 0 ${h - r}V${r}Q0 0 ${r} 0Z`;
  }
  if (prst === 'triangle') return `M${w / 2} 0L${w} ${h}L0 ${h}Z`;
  if (prst === 'rtTriangle') return `M0 0L${w} ${h}L0 ${h}Z`;
  if (prst === 'diamond') return `M${w / 2} 0L${w} ${h / 2}L${w / 2} ${h}L0 ${h / 2}Z`;
  return `M0 0H${w}V${h}H0Z`;
}

// A w: attribute (w:val, w:left, ...).
const wa = (el, name) => (el ? el.getAttribute(`w:${name}`) : null);

let canvas = null;
/** Width of `text` in px: measured in a browser, else estimated at half an em per character. */
function textWidth(text, sizePx, bold, font) {
  if (canvas == null) canvas = typeof document !== 'undefined' && document.createElement('canvas').getContext?.('2d') || false;
  if (!canvas) return text.length * sizePx * 0.5;
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
  // (Paragraphs in block content controls too: a cover page's title and author boxes hold theirs in w:sdt.)
  const paras = (el) => els(el).flatMap((n) => (n.localName === 'p' ? [n] : n.localName === 'sdt' ? paras(child(n, 'sdtContent')) : []));
  for (const p of paras(desc(txbx, 'txbxContent'))) {
    const pPr = child(p, 'pPr');
    // The paragraph's style (and the document defaults) under its own formatting.
    const st = ctx.para ? ctx.para(wa(child(pPr, 'pStyle'), 'val')) : { pPr: {}, rPr: {}, font: font0 };
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
    const right = w - r - twipsToPx(indR);
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
      const open = `<tspan font-family="${esc(font)}" font-size="${sizePx.toFixed(2)}" fill="#${c && c !== 'auto' ? c : '000000'}"${bold ? ' font-weight="bold"' : ''}${italic ? ' font-style="italic"' : ''}${letter ? ` letter-spacing="${letter.toFixed(2)}"` : ''}>`;
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
    // Each font's own single line (Calibri 1.22em, Times New Roman 1.15em), as in the body.
    const one = sizePx * lineF;
    const lineH = line && rule === 'exact' ? line / TWIPS_PER_PX : line && rule === 'atLeast' ? Math.max(line / TWIPS_PER_PX, one) : one * (line ? line / 240 : 1);
    y += (Number(sp('before')) || 0) / TWIPS_PER_PX;
    for (const ln of lines) {
      // Baseline: the font's ascent (0.778 of its line: 0.95em of Calibri's 1.22) below the line top, centred in an exact line's extra room.
      const base = y + (lineH - one) / 2 + one * 0.778;
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
    y += (Number(sp('after')) || 0) / TWIPS_PER_PX;
  }
  return out.join('');
}

function shapeSvg(el, ctx) {
  const name = el.localName;
  if (name === 'wgp' || name === 'grpSp') {
    const xfrm = child(child(el, 'grpSpPr'), 'xfrm');
    const inner = els(el).filter((n) => ['wsp', 'pic', 'grpSp'].includes(n.localName)).map((n) => shapeSvg(n, ctx)).join('');
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
  if (name !== 'wsp') return '';
  const style = child(el, 'style');
  const body = `<path d="${geometry(spPr, w, h)}" ${paints(spPr, style, ctx.colors)} vector-effect="non-scaling-stroke"/>`;
  const txbx = child(el, 'txbx');
  const ln = child(spPr, 'ln');
  const line = ln && !child(ln, 'noFill') && child(ln, 'solidFill') ? px(ln.getAttribute('w') || EMU_PER_PX) : 0;
  return `<g transform="${t}">${body}${txbx ? textBox(txbx, child(el, 'bodyPr'), w, h, ctx, line) : ''}</g>`;
}

/**
 * SVG for a drawing whose graphic is a shape or a group of shapes, sized to its extent.
 * @param {Element} drawing w:drawing
 * @param {{ image: (rId: string) => string|null, colors?: object, font?: string }} ctx image() returns a data: URL
 */
export function shapesSvg(drawing, ctx, width, height) {
  const data = desc(drawing, 'graphicData');
  const root = data && els(data).find((n) => n.localName === 'wgp' || n.localName === 'wsp');
  if (!root || !width || !height) return null;
  const inner = shapeSvg(root, { colors: {}, font: 'Calibri', ...ctx });
  if (!inner) return null;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" overflow="visible">${inner}</svg>`;
}
