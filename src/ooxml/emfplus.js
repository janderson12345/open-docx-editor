// EMF+ (GDI+) records as SVG. An EMF+ file carries its drawing in comment records of an ordinary EMF; a dual file adds a
// GDI copy, an EMF+ only file has none (so the GDI reader alone drew it blank). This reads the EMF+ records: objects
// (brushes, pens, paths, images, fonts, string formats, regions), fills and outlines, beziers and curves, text, images,
// clipping, and the world, page and container transforms. Each drawing is an SVG element in world units with the
// matrix that takes it to the picture's device units, so pens, text and gradients turn and scale as GDI+ draws them.
//
// GDI records between a GetDC record and the next EMF+ record belong to the drawing too: the result says which, so
// the GDI reader draws those (and only those) in order with these.

const n = (v) => +(+v).toFixed(3);
const hex = (v) => `#${[(v >> 16) & 255, (v >> 8) & 255, v & 255].map((k) => k.toString(16).padStart(2, '0')).join('')}`;
const alpha = (v) => ((v >>> 24) & 255) / 255;
const mul = (A, B) => [A[0] * B[0] + A[1] * B[2], A[0] * B[1] + A[1] * B[3], A[2] * B[0] + A[3] * B[2], A[2] * B[1] + A[3] * B[3],
  A[4] * B[0] + A[5] * B[2] + B[4], A[4] * B[1] + A[5] * B[3] + B[5]];
const apply = (M, x, y) => [M[0] * x + M[2] * y + M[4], M[1] * x + M[3] * y + M[5]];
const ID = [1, 0, 0, 1, 0, 0];

/** True when the EMF in `dv` holds EMF+ records (its first comment is an EMF+ header). */
export function hasEmfPlus(dv) {
  for (let off = 0; off + 8 <= dv.byteLength;) {
    const type = dv.getUint32(off, true), size = dv.getUint32(off + 4, true);
    if (size < 8) return false;
    if (type === 70) return size >= 28 && dv.getUint32(off + 12, true) === 0x2b464d45 && dv.getUint16(off + 16, true) === 0x4001;
    if (type !== 1) return false; // the header comment comes straight after the EMF header
    off += size;
  }
  return false;
}

/**
 * The EMF+ drawing of the EMF in `bytes`: { items: [{ at, s }], defs: [], gdi: [[from, to]], unsupported } — SVG
 * elements in device units less (bx, by), each with the byte offset of the record it came from; `gdi` the GDI record
 * ranges to draw with them; `unsupported` true when a drawing record was skipped (a dual file is then better drawn
 * from its GDI copy).
 */
export function emfPlusToSvg(input, { bx, by }, h) {
  // (Rebound while an object split over several records is read from a joined copy of its parts.)
  let bytes = input;
  let dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const items = [], defs = [], gdi = [];
  let unsupported = false, defId = 0;
  const id = (p) => `${p}${++defId}`;
  let dpiX = 96, dpiY = 96;
  const objects = [];
  let pending = null; // an object split over several records: { id, total, parts, got }
  let world = ID.slice(), pageUnit = 2, pageScale = 1;
  let clip = null; // clipPath id
  let interp = 0;
  const saved = new Map();
  let gdiFrom = -1;

  const unitPx = (u) => ({ 2: 1, 3: dpiX / 72, 4: dpiX, 5: dpiX / 300, 6: dpiX / 25.4 }[u] ?? 1);
  const k = () => unitPx(pageUnit) * pageScale;
  const toDevice = () => { const s = k(); return [world[0] * s, world[1] * s, world[2] * s, world[3] * s, world[4] * s - bx, world[5] * s - by]; };
  const mat = (M) => `matrix(${M.map((v) => +v.toPrecision(7)).join(' ')})`; // (significant digits: EMU-scaled worlds are tiny)
  const tf = () => ` transform="${mat(toDevice())}"`;
  let at = 0;
  const emit = (s) => items.push({ at, s: clip ? `<g clip-path="url(#${clip})">${s}</g>` : s });

  // ----- reading -----
  const rd = (o) => ({
    o,
    u8() { return dv.getUint8(this.o++); },
    u16() { const v = dv.getUint16(this.o, true); this.o += 2; return v; },
    i16() { const v = dv.getInt16(this.o, true); this.o += 2; return v; },
    u32() { const v = dv.getUint32(this.o, true); this.o += 4; return v; },
    i32() { const v = dv.getInt32(this.o, true); this.o += 4; return v; },
    f32() { const v = dv.getFloat32(this.o, true); this.o += 4; return v; },
    m() { return [this.f32(), this.f32(), this.f32(), this.f32(), this.f32(), this.f32()]; },
  });
  const points = (r, count, flags) => {
    const pts = [];
    if (flags & 0x0800) { // relative, 7 or 15 bits each
      let x = 0, y = 0;
      const one = () => { const b = r.u8(); if (!(b & 0x80)) return b & 0x40 ? b - 0x80 : b; const v = ((b & 0x7f) << 8) | r.u8(); return v & 0x4000 ? v - 0x8000 : v; };
      for (let i = 0; i < count; i++) { x += one(); y += one(); pts.push([x, y]); }
    } else if (flags & 0x4000) for (let i = 0; i < count; i++) pts.push([r.i16(), r.i16()]);
    else for (let i = 0; i < count; i++) pts.push([r.f32(), r.f32()]);
    return pts;
  };
  const rects = (r, count, compressed) => {
    const out = [];
    for (let i = 0; i < count; i++) out.push(compressed ? [r.i16(), r.i16(), r.i16(), r.i16()] : [r.f32(), r.f32(), r.f32(), r.f32()]);
    return out;
  };
  const utf16 = (r, len) => { let s = ''; for (let i = 0; i < len; i++) s += String.fromCharCode(r.u16()); return s; };

  // ----- objects -----
  const readPath = (r) => {
    r.u32(); // version
    const count = r.u32(), flags = r.u32();
    const pts = points(r, count, flags & 0x4000 ? 0x4000 : flags & 0x0800 ? 0x0800 : 0);
    const types = [];
    if (flags & 0x1000) { // run-length encoded types
      while (types.length < count) {
        const a = r.u8(), b = r.u8();
        const run = a & 0x3f;
        for (let i = 0; i < run && types.length < count; i++) types.push(b);
      }
    } else for (let i = 0; i < count; i++) types.push(r.u8());
    let d = '';
    for (let i = 0; i < count; i++) {
      const t = types[i] & 7;
      if (t === 0) d += `M${n(pts[i][0])} ${n(pts[i][1])}`;
      else if (t === 3 && i + 2 < count) {
        d += ` C${n(pts[i][0])} ${n(pts[i][1])} ${n(pts[i + 1][0])} ${n(pts[i + 1][1])} ${n(pts[i + 2][0])} ${n(pts[i + 2][1])}`;
        i += 2;
      } else d += ` L${n(pts[i][0])} ${n(pts[i][1])}`;
      if (types[i] & 0x80) d += 'Z';
    }
    return { kind: 'path', d };
  };
  const readImage = (r, end) => {
    r.u32(); // version
    const type = r.u32();
    if (type === 1) { // bitmap
      const w = r.i32(), hh = r.i32(), stride = r.i32(), format = r.u32(), compressed = r.u32();
      if (compressed === 1) {
        const data = bytes.subarray(r.o, end);
        const mime = data[0] === 0x89 && data[1] === 0x50 ? 'image/png' : data[0] === 0xff && data[1] === 0xd8 ? 'image/jpeg'
          : data[0] === 0x47 && data[1] === 0x49 ? 'image/gif' : data[0] === 0x42 && data[1] === 0x4d ? 'image/bmp' : null;
        if (!mime) return null;
        const [iw, ih] = w > 0 && hh > 0 ? [w, hh] : imageSize(data); // (GDI+ leaves a compressed bitmap's size 0)
        return { kind: 'image', w: iw || 1, h: ih || 1, href: h.dataUrl(mime, data) };
      }
      const bpp = (format >> 8) & 255;
      let palette = null;
      if (format & 0x00010000) { // indexed: a palette comes first
        r.u32();
        const count = r.u32();
        palette = [];
        for (let i = 0; i < count; i++) palette.push(r.u32());
      }
      if (!(w > 0 && hh > 0) || r.o + stride * hh > end) return null;
      const rgba = new Uint8Array(w * hh * 4);
      const premul = format === 0x000e200b;
      for (let y = 0; y < hh; y++) {
        const row = r.o + y * stride;
        for (let x = 0; x < w; x++) {
          const d = (y * w + x) * 4;
          let c;
          if (bpp === 32) { c = dv.getUint32(row + x * 4, true); if (format === 0x00022009) c |= 0xff000000; }
          else if (bpp === 24) c = 0xff000000 | (bytes[row + x * 3 + 2] << 16) | (bytes[row + x * 3 + 1] << 8) | bytes[row + x * 3];
          else if (bpp === 16 && format === 0x00021005) { const v = dv.getUint16(row + x * 2, true); c = 0xff000000 | (((v >> 11) * 255 / 31) << 16) | ((((v >> 5) & 63) * 255 / 63) << 8) | ((v & 31) * 255 / 31); }
          else if (palette && (bpp === 8 || bpp === 4 || bpp === 1)) {
            const bit = x * bpp, byte = bytes[row + (bit >> 3)];
            const idx = bpp === 8 ? byte : (byte >> (8 - bpp - (bit & 7))) & ((1 << bpp) - 1);
            c = palette[idx] ?? 0;
          } else return null;
          const a = (c >>> 24) & 255;
          const f = premul && a ? 255 / a : 1;
          rgba[d] = Math.min(255, ((c >> 16) & 255) * f); rgba[d + 1] = Math.min(255, ((c >> 8) & 255) * f); rgba[d + 2] = Math.min(255, (c & 255) * f); rgba[d + 3] = a;
        }
      }
      return { kind: 'image', w, h: hh, href: h.png(w, hh, rgba) };
    }
    if (type === 2) { // a metafile inside: drawn by the EMF reader, as SVG
      const mtype = r.u32(), size = r.u32();
      if (mtype < 3) return null; // WMF
      const svg = h.emfToSvg(bytes.subarray(r.o, Math.min(end, r.o + size)));
      if (!svg) return null;
      const m = /viewBox="0 0 ([\d.]+) ([\d.]+)"/.exec(svg);
      return { kind: 'image', w: m ? +m[1] : 1, h: m ? +m[2] : 1, href: `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(svg)))}` };
    }
    return null;
  };
  const readBrush = (r, end) => {
    r.u32(); // version
    const type = r.u32();
    if (type === 0) return { kind: 'brush', type: 'solid', color: r.u32() };
    if (type === 1) return { kind: 'brush', type: 'hatch', style: r.u32(), fore: r.u32(), back: r.u32() };
    if (type === 2) { // texture
      const flags = r.u32(); r.u32();
      const m = flags & 2 ? r.m() : null;
      const img = readImage(r, end);
      return img ? { kind: 'brush', type: 'texture', img, m } : { kind: 'brush', type: 'solid', color: 0x80808080 };
    }
    if (type === 4) { // linear gradient
      const flags = r.u32(); r.u32();
      const rect = [r.f32(), r.f32(), r.f32(), r.f32()];
      const c1 = r.u32(), c2 = r.u32();
      r.u32(); r.u32();
      const m = flags & 2 ? r.m() : null;
      let stops = [[0, c1], [1, c2]];
      if (flags & 4) { // preset colours
        const cnt = r.u32(); const pos = []; for (let i = 0; i < cnt; i++) pos.push(r.f32());
        stops = pos.map((p) => [p, r.u32()]);
      } else if (flags & 8) { // blend factors (horizontal)
        const cnt = r.u32(); const pos = []; for (let i = 0; i < cnt; i++) pos.push(r.f32());
        const fac = []; for (let i = 0; i < cnt; i++) fac.push(r.f32());
        stops = pos.map((p, i) => [p, lerp(c1, c2, fac[i])]);
      }
      return { kind: 'brush', type: 'linear', rect, m, stops };
    }
    if (type === 3) { // path gradient: drawn as a radial one from the centre colour to the first surrounding colour
      const flags = r.u32(); r.u32();
      const center = r.u32(), cx = r.f32(), cy = r.f32();
      const cnt = r.u32(); const surround = []; for (let i = 0; i < cnt; i++) surround.push(r.u32());
      let pts = [];
      if (flags & 1) { const size = r.u32(); const p = readPath(r); r.o = Math.min(end, r.o); void size; pts = (p.d.match(/-?[\d.]+ -?[\d.]+/g) || []).map((s) => s.split(' ').map(Number)); }
      else { const pc = r.u32(); pts = points(r, pc, 0); }
      const rad = Math.max(1, ...pts.map(([x, y]) => Math.hypot(x - cx, y - cy)));
      return { kind: 'brush', type: 'radial', cx, cy, r: rad, c1: center, c2: surround[0] ?? center };
    }
    return { kind: 'brush', type: 'solid', color: 0xff808080 };
  };
  const lerp = (a, b, t) => {
    const ch = (s) => Math.round(((a >>> s) & 255) * (1 - t) + ((b >>> s) & 255) * t);
    return ((ch(24) << 24) | (ch(16) << 16) | (ch(8) << 8) | ch(0)) >>> 0;
  };
  const readPen = (r, end) => {
    r.u32(); r.u32(); // version, type
    const flags = r.u32(), unit = r.u32(), width = r.f32();
    const pen = { kind: 'pen', unit, width, style: 0, cap: 0, join: 0, miter: 10, dashes: null };
    if (flags & 0x01) pen.m = r.m();
    if (flags & 0x02) pen.cap = r.i32();
    if (flags & 0x04) r.i32();
    if (flags & 0x08) pen.join = r.i32();
    if (flags & 0x10) pen.miter = r.f32();
    if (flags & 0x20) pen.style = r.i32();
    if (flags & 0x40) r.i32();
    if (flags & 0x80) r.f32();
    if (flags & 0x100) { const c = r.u32(); pen.dashes = []; for (let i = 0; i < c; i++) pen.dashes.push(r.f32()); }
    if (flags & 0x200) r.i32();
    if (flags & 0x400) { const c = r.u32(); r.o += c * 4; }
    if (flags & 0x800) { const s = r.u32(); r.o += s; }
    if (flags & 0x1000) { const s = r.u32(); r.o += s; }
    pen.brush = r.o < end ? readBrush(r, end) : { kind: 'brush', type: 'solid', color: 0xff000000 };
    return pen;
  };
  const readFont = (r) => {
    r.u32();
    const em = r.f32(), unit = r.u32(), style = r.i32(); r.u32();
    const len = r.u32();
    let face = utf16(r, len).replace(/\0.*$/, '');
    // (GDI+ may write the family in capitals, "ARIAL": font names match whatever their case, the stand-ins don't.)
    if (face === face.toUpperCase()) face = face.toLowerCase().replace(/\b[a-z]/g, (c) => c.toUpperCase());
    return { kind: 'font', em, unit, bold: !!(style & 1), italic: !!(style & 2), underline: !!(style & 4), strike: !!(style & 8), face };
  };
  const readFormat = (r) => {
    r.u32();
    const flags = r.u32(); r.u32();
    return { kind: 'format', flags, align: r.u32(), lineAlign: r.u32() };
  };
  const readRegion = (r) => {
    r.u32(); r.u32(); // version, node count
    const node = () => {
      const type = r.u32();
      if (type >= 1 && type <= 5) { const a = node(); const b = node(); return { op: type, a, b }; }
      if (type === 0x10000000) { const [x, y, w, hh] = [r.f32(), r.f32(), r.f32(), r.f32()]; return { d: `M${n(x)} ${n(y)}h${n(w)}v${n(hh)}h${n(-w)}Z` }; }
      if (type === 0x10000001) { const size = r.u32(); const start = r.o; const p = readPath(r); r.o = start + size; return { d: p.d }; }
      if (type === 0x10000002) return { d: '' };
      return { d: 'M-1e6 -1e6h2e6v2e6h-2e6Z', infinite: true };
    };
    return { kind: 'region', node: node() };
  };
  const readObject = (type, r, end) => {
    switch (type) {
      case 1: return readBrush(r, end);
      case 2: return readPen(r, end);
      case 3: return readPath(r);
      case 4: return readRegion(r);
      case 5: return readImage(r, end);
      case 6: return readFont(r);
      case 7: return readFormat(r);
      default: return null;
    }
  };

  // ----- painting -----
  const paint = (brush) => { // [paint, opacity] for fill or stroke, in the current world units
    if (!brush) return ['none', 1];
    if (typeof brush === 'number') return [hex(brush), alpha(brush)];
    if (brush.type === 'solid') return [hex(brush.color), alpha(brush.color)];
    if (brush.type === 'hatch') {
      const g = id('h');
      const lines = { 0: 'M0 4h8', 1: 'M4 0v8', 2: 'M0 0l8 8', 3: 'M8 0l-8 8', 4: 'M0 4h8M4 0v8', 5: 'M0 0l8 8M8 0l-8 8' }[brush.style];
      if (!lines) return [hex(lerp(brush.fore, brush.back, 0.5)), 1];
      defs.push(`<pattern id="${g}" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="${hex(brush.back)}" fill-opacity="${n(alpha(brush.back))}"/><path d="${lines}" stroke="${hex(brush.fore)}" stroke-opacity="${n(alpha(brush.fore))}" stroke-width="1"/></pattern>`);
      return [`url(#${g})`, 1];
    }
    if (brush.type === 'texture') {
      const g = id('t');
      defs.push(`<pattern id="${g}" width="${brush.img.w}" height="${brush.img.h}" patternUnits="userSpaceOnUse"${brush.m ? ` patternTransform="${mat(brush.m)}"` : ''}><image width="${brush.img.w}" height="${brush.img.h}" href="${brush.img.href}"/></pattern>`);
      return [`url(#${g})`, 1];
    }
    if (brush.type === 'linear') {
      const g = id('g');
      const [x, y, w, hh] = brush.rect;
      let p1 = [x, y + hh / 2], p2 = [x + w, y + hh / 2];
      if (brush.m) { p1 = apply(brush.m, ...p1); p2 = apply(brush.m, ...p2); }
      const stops = brush.stops.map(([o, c]) => `<stop offset="${n(o)}" stop-color="${hex(c)}" stop-opacity="${n(alpha(c))}"/>`).join('');
      defs.push(`<linearGradient id="${g}" gradientUnits="userSpaceOnUse" x1="${n(p1[0])}" y1="${n(p1[1])}" x2="${n(p2[0])}" y2="${n(p2[1])}">${stops}</linearGradient>`);
      return [`url(#${g})`, 1];
    }
    if (brush.type === 'radial') {
      const g = id('g');
      defs.push(`<radialGradient id="${g}" gradientUnits="userSpaceOnUse" cx="${n(brush.cx)}" cy="${n(brush.cy)}" r="${n(brush.r)}"><stop offset="0" stop-color="${hex(brush.c1)}" stop-opacity="${n(alpha(brush.c1))}"/><stop offset="1" stop-color="${hex(brush.c2)}" stop-opacity="${n(alpha(brush.c2))}"/></radialGradient>`);
      return [`url(#${g})`, 1];
    }
    return ['#808080', 1];
  };
  const brushOf = (flags, v) => (flags & 0x8000 ? v : objects[v & 255]);
  const fillAttr = (brush, rule = 'evenodd') => { const [p, o] = paint(brush); return `fill="${p}"${o < 1 ? ` fill-opacity="${n(o)}"` : ''} fill-rule="${rule}" stroke="none"`; };
  const strokeAttr = (pen) => {
    if (!pen || pen.kind !== 'pen') return 'stroke="#000" stroke-width="1" vector-effect="non-scaling-stroke"';
    const [p, o] = paint(pen.brush);
    let w = pen.width;
    if (pen.unit && pen.unit !== pageUnit && pen.unit > 1) w = (w * unitPx(pen.unit)) / k();
    const thin = !(w > 0);
    const dash = pen.style === 5 && pen.dashes ? pen.dashes : { 1: [3, 1], 2: [1, 1], 3: [3, 1, 1, 1], 4: [3, 1, 1, 1, 1, 1] }[pen.style];
    const dw = thin ? 1 : w;
    return `fill="none" stroke="${p}"${o < 1 ? ` stroke-opacity="${n(o)}"` : ''} stroke-width="${n(dw)}"${thin ? ' vector-effect="non-scaling-stroke"' : ''}`
      + `${dash ? ` stroke-dasharray="${dash.map((v) => n(v * dw)).join(' ')}"` : ''}`
      + ` stroke-linecap="${['butt', 'square', 'round'][pen.cap] || 'butt'}" stroke-linejoin="${['miter', 'bevel', 'round', 'miter'][pen.join] || 'miter'}" stroke-miterlimit="${n(Math.max(1, pen.miter))}"`;
  };
  const rectD = ([x, y, w, hh]) => `M${n(x)} ${n(y)}h${n(w)}v${n(hh)}h${n(-w)}Z`;
  const ellipseD = ([x, y, w, hh]) => `M${n(x)} ${n(y + hh / 2)}a${n(w / 2)} ${n(hh / 2)} 0 1 0 ${n(w)} 0a${n(w / 2)} ${n(hh / 2)} 0 1 0 ${n(-w)} 0Z`;
  const arcD = ([x, y, w, hh], start, sweep, pie) => {
    const cx = x + w / 2, cy = y + hh / 2, rx = w / 2, ry = hh / 2;
    const at2 = (deg) => { const t = (deg * Math.PI) / 180; return [cx + rx * Math.cos(t), cy + ry * Math.sin(t)]; };
    if (Math.abs(sweep) >= 360) return ellipseD([x, y, w, hh]);
    const [x1, y1] = at2(start), [x2, y2] = at2(start + sweep);
    const arc = `A${n(rx)} ${n(ry)} 0 ${Math.abs(sweep) > 180 ? 1 : 0} ${sweep > 0 ? 1 : 0} ${n(x2)} ${n(y2)}`;
    return pie ? `M${n(cx)} ${n(cy)}L${n(x1)} ${n(y1)}${arc}Z` : `M${n(x1)} ${n(y1)}${arc}`;
  };
  const polyD = (pts, close) => (pts.length ? `M${pts.map(([x, y]) => `${n(x)} ${n(y)}`).join('L')}${close ? 'Z' : ''}` : '');
  const curveD = (pts, tension, close) => { // a cardinal spline as beziers
    if (pts.length < 2) return '';
    const t = tension / 3, P = (i) => pts[close ? (i + pts.length) % pts.length : Math.max(0, Math.min(pts.length - 1, i))];
    let d = `M${n(pts[0][0])} ${n(pts[0][1])}`;
    const segs = close ? pts.length : pts.length - 1;
    for (let i = 0; i < segs; i++) {
      const p0 = P(i - 1), p1 = P(i), p2 = P(i + 1), p3 = P(i + 2);
      d += `C${n(p1[0] + t * (p2[0] - p0[0]))} ${n(p1[1] + t * (p2[1] - p0[1]))} ${n(p2[0] - t * (p3[0] - p1[0]))} ${n(p2[1] - t * (p3[1] - p1[1]))} ${n(p2[0])} ${n(p2[1])}`;
    }
    return close ? `${d}Z` : d;
  };
  const shape = (d, attrs) => { if (d) emit(`<path d="${d}" ${attrs}${tf()}/>`); };

  // Clipping: a clip path in device units (its shape carries the transform it was set under), chained to the clip
  // before it for an intersection.
  const setClip = (d, mode, infinite) => {
    if (mode === 0 && infinite) { clip = null; return; }
    if (mode !== 0 && mode !== 1) return; // union, xor, exclude, complement: the clip is left as it was
    const c = id('p');
    const chain = mode === 1 && clip ? ` clip-path="url(#${clip})"` : '';
    defs.push(`<clipPath id="${c}"${chain}><path d="${d}" clip-rule="nonzero"${tf()}/></clipPath>`);
    clip = c;
  };
  const regionD = (node) => { // a region as one path: unions joined; for an intersection or exclusion, the first part
    if (!node) return '';
    if (node.d !== undefined) return node.d;
    if (node.op === 2) return `${regionD(node.a)} ${regionD(node.b)}`;
    return regionD(node.a);
  };

  const image = (img, src, dest) => { // dest: the image's top left, top right and bottom left
    if (!img) { unsupported = true; return; }
    const [p0, p1, p2] = dest;
    const [sx, sy, sw, sh] = src[2] > 0 && src[3] > 0 ? src : [0, 0, img.w, img.h];
    const A = [(p1[0] - p0[0]) / sw, (p1[1] - p0[1]) / sw, (p2[0] - p0[0]) / sh, (p2[1] - p0[1]) / sh, p0[0], p0[1]];
    const M = mul([1, 0, 0, 1, -sx, -sy], mul(A, toDevice()));
    const crop = sx || sy || sw !== img.w || sh !== img.h;
    const smooth = interp === 5 ? ' style="image-rendering:pixelated"' : '';
    const el = `<image width="${img.w}" height="${img.h}" preserveAspectRatio="none" href="${img.href}"${smooth}/>`;
    emit(`<g transform="${mat(M)}">${crop ? `<svg x="${n(sx)}" y="${n(sy)}" width="${n(sw)}" height="${n(sh)}" viewBox="${n(sx)} ${n(sy)} ${n(sw)} ${n(sh)}" overflow="hidden">${el}</svg>` : el}</g>`);
  };

  const text = (str, font, brush, fmt, layout) => {
    if (!font || font.kind !== 'font' || !str) return;
    const em = font.unit === 0 || font.unit === pageUnit ? font.em : (font.em * unitPx(font.unit)) / k();
    const [asc, desc] = h.win(font.face);
    const [x, y, w, hh] = layout;
    const typographic = fmt && fmt.flags & 0x4000;
    const pad = typographic ? 0 : em / 6;
    const align = fmt ? fmt.align : 0, valign = fmt ? fmt.lineAlign : 0;
    const lines = str.replace(/\r\n?/g, '\n').split('\n');
    const lh = (asc + desc) * em;
    let top = y;
    if (hh > 0 && valign === 1) top = y + (hh - lh * lines.length) / 2;
    else if (hh > 0 && valign === 2) top = y + hh - lh * lines.length;
    const anchor = w > 0 && align === 1 ? 'middle' : w > 0 && align === 2 ? 'end' : 'start';
    const tx = anchor === 'middle' ? x + w / 2 : anchor === 'end' ? x + w - pad : x + pad;
    const [p, o] = paint(brush);
    const deco = [font.underline && 'underline', font.strike && 'line-through'].filter(Boolean).join(' ');
    const spans = lines.map((s, i) => `<tspan x="${n(tx)}" y="${n(top + asc * em + i * lh)}">${h.esc(s)}</tspan>`).join('');
    emit(`<text font-family="${h.esc(h.fontStack(font.face))}" font-size="${n(em)}"${font.bold ? ' font-weight="bold"' : ''}${font.italic ? ' font-style="italic"' : ''}${deco ? ` text-decoration="${deco}"` : ''} fill="${p}"${o < 1 ? ` fill-opacity="${n(o)}"` : ''} text-anchor="${anchor}" xml:space="preserve"${tf()}>${spans}</text>`);
  };

  // ----- the records -----
  const record = (type, flags, r, end) => {
    const lo = flags & 255;
    switch (type) {
      case 0x4001: r.u32(); r.u32(); dpiX = r.u32() || 96; dpiY = r.u32() || 96; void dpiY; break; // header
      case 0x4002: case 0x4003: break; // end of file, comment
      case 0x4004: gdiFrom = -2; break; // GetDC: GDI records follow
      case 0x4008: { // object
        const otype = (flags >> 8) & 0x7f;
        // An object too big for one record is split over several, each flagged as continued and starting with the
        // object's whole size (the last one too, as GDI+ writes them): read it once all its bytes are in.
        const part = pending && pending.id === lo;
        if (flags & 0x8000 || part) {
          const total = flags & 0x8000 ? r.u32() : pending.total;
          if (!part) pending = { id: lo, total, parts: [], got: 0 };
          const piece = bytes.subarray(r.o, end);
          pending.parts.push(piece); pending.got += piece.length;
          if (pending.got < pending.total && flags & 0x8000) break;
          const all = new Uint8Array(pending.got);
          let q = 0;
          for (const p of pending.parts) { all.set(p, q); q += p.length; }
          pending = null;
          const [b0, d0] = [bytes, dv];
          bytes = all; dv = new DataView(all.buffer);
          try { objects[lo] = readObject(otype, rd(0), all.length); } finally { bytes = b0; dv = d0; }
          break;
        }
        objects[lo] = readObject(otype, r, end);
        break;
      }
      case 0x4009: { const c = r.u32(); emit(`<rect x="-1e5" y="-1e5" width="2e5" height="2e5" fill="${hex(c)}"${alpha(c) < 1 ? ` fill-opacity="${n(alpha(c))}"` : ''}/>`); break; } // clear
      case 0x400a: { const b = r.u32(), c = r.u32(); shape(rects(r, c, flags & 0x4000).map(rectD).join(''), fillAttr(brushOf(flags, b))); break; } // FillRects
      case 0x400b: { const c = r.u32(); shape(rects(r, c, flags & 0x4000).map(rectD).join(''), strokeAttr(objects[lo])); break; } // DrawRects
      case 0x400c: { const b = r.u32(), c = r.u32(); shape(polyD(points(r, c, flags), true), fillAttr(brushOf(flags, b))); break; } // FillPolygon
      case 0x400d: { const c = r.u32(); shape(polyD(points(r, c, flags), flags & 0x2000), strokeAttr(objects[lo])); break; } // DrawLines
      case 0x400e: { const b = r.u32(); shape(ellipseD(rects(r, 1, flags & 0x4000)[0]), fillAttr(brushOf(flags, b))); break; } // FillEllipse
      case 0x400f: shape(ellipseD(rects(r, 1, flags & 0x4000)[0]), strokeAttr(objects[lo])); break; // DrawEllipse
      case 0x4010: { const b = r.u32(), s = r.f32(), w = r.f32(); shape(arcD(rects(r, 1, flags & 0x4000)[0], s, w, true), fillAttr(brushOf(flags, b))); break; } // FillPie
      case 0x4011: case 0x4012: { const s = r.f32(), w = r.f32(); shape(arcD(rects(r, 1, flags & 0x4000)[0], s, w, type === 0x4011), strokeAttr(objects[lo])); break; } // DrawPie, DrawArc
      case 0x4013: { const b = r.u32(); const reg = objects[lo]; if (reg && reg.kind === 'region') shape(regionD(reg.node), fillAttr(brushOf(flags, b), 'nonzero')); break; } // FillRegion
      case 0x4014: { const b = r.u32(); const p = objects[lo]; if (p && p.kind === 'path') shape(p.d, fillAttr(brushOf(flags, b))); break; } // FillPath
      case 0x4015: { const pen = objects[r.u32() & 255]; const p = objects[lo]; if (p && p.kind === 'path') shape(p.d, strokeAttr(pen)); break; } // DrawPath
      case 0x4016: { const b = r.u32(), t = r.f32(), c = r.u32(); shape(curveD(points(r, c, flags), t, true), fillAttr(brushOf(flags, b), flags & 0x2000 ? 'nonzero' : 'evenodd')); break; } // FillClosedCurve
      case 0x4017: { const t = r.f32(), c = r.u32(); shape(curveD(points(r, c, flags), t, true), strokeAttr(objects[lo])); break; } // DrawClosedCurve
      case 0x4018: { // DrawCurve
        const t = r.f32(), offset = r.u32(), segs = r.u32(), c = r.u32();
        const pts = points(r, c, flags);
        const d = curveD(pts, t, false);
        if (offset === 0 && segs >= c - 1) shape(d, strokeAttr(objects[lo]));
        else { // only part of it: the segments from `offset`
          const parts = d.split('C'); shape(`M${n(pts[offset][0])} ${n(pts[offset][1])}C${parts.slice(1 + offset, 1 + offset + segs).join('C')}`, strokeAttr(objects[lo]));
        }
        break;
      }
      case 0x4019: { // DrawBeziers
        const c = r.u32(); const pts = points(r, c, flags);
        let d = pts.length ? `M${n(pts[0][0])} ${n(pts[0][1])}` : '';
        for (let i = 1; i + 2 < pts.length; i += 3) {
          d +=`C${n(pts[i][0])} ${n(pts[i][1])} ${n(pts[i + 1][0])} ${n(pts[i + 1][1])} ${n(pts[i + 2][0])} ${n(pts[i + 2][1])}`;
        }
        shape(d, strokeAttr(objects[lo]));
        break;
      }
      case 0x401a: { // DrawImage
        r.u32(); r.u32();
        const src = [r.f32(), r.f32(), r.f32(), r.f32()];
        const [x, y, w, hh] = rects(r, 1, flags & 0x4000)[0];
        image(objects[lo], src, [[x, y], [x + w, y], [x, y + hh]]);
        break;
      }
      case 0x401b: { // DrawImagePoints
        r.u32(); r.u32();
        const src = [r.f32(), r.f32(), r.f32(), r.f32()];
        const c = r.u32();
        const pts = points(r, c, flags);
        if (pts.length >= 3) image(objects[lo], src, pts);
        break;
      }
      case 0x401c: { // DrawString
        const b = r.u32(), f = r.u32(), len = r.u32();
        const layout = [r.f32(), r.f32(), r.f32(), r.f32()];
        text(utf16(r, len), objects[lo], brushOf(flags, b), objects[f & 255], layout);
        break;
      }
      case 0x4036: { // DrawDriverString: drawn only when its glyphs are characters
        const b = r.u32(), opts = r.u32(); r.u32();
        const c = r.u32();
        if (!(opts & 1)) { unsupported = true; break; }
        let s = ''; for (let i = 0; i < c; i++) s += String.fromCharCode(r.u16());
        const pts = []; for (let i = 0; i < c; i++) pts.push([r.f32(), r.f32()]);
        const font = objects[lo];
        if (!font || font.kind !== 'font' || !c) break;
        const em = font.unit === 0 || font.unit === pageUnit ? font.em : (font.em * unitPx(font.unit)) / k();
        const [p, o] = paint(brushOf(flags, b));
        emit(`<text font-family="${h.esc(h.fontStack(font.face))}" font-size="${n(em)}"${font.bold ? ' font-weight="bold"' : ''}${font.italic ? ' font-style="italic"' : ''} fill="${p}"${o < 1 ? ` fill-opacity="${n(o)}"` : ''} xml:space="preserve" x="${pts.map(([x]) => n(x)).join(' ')}" y="${pts.map(([, y]) => n(y)).join(' ')}"${tf()}>${h.esc(s)}</text>`);
        break;
      }
      case 0x4021: interp = lo; break; // SetInterpolationMode
      case 0x401d: case 0x401e: case 0x401f: case 0x4020: case 0x4022: case 0x4023: case 0x4024: break; // rendering hints
      case 0x4025: case 0x4028: saved.set(r.u32(), { world: world.slice(), pageUnit, pageScale, clip }); break; // Save, BeginContainerNoParams
      case 0x4027: { // BeginContainer: dest and source rectangles set a transform for the container
        const dst = [r.f32(), r.f32(), r.f32(), r.f32()], src = [r.f32(), r.f32(), r.f32(), r.f32()];
        saved.set(r.u32(), { world: world.slice(), pageUnit, pageScale, clip });
        const u = unitPx(lo) / k();
        const sx = src[2] ? (dst[2] * u) / src[2] : 1, sy = src[3] ? (dst[3] * u) / src[3] : 1;
        world = mul([sx, 0, 0, sy, dst[0] * u - src[0] * sx, dst[1] * u - src[1] * sy], world);
        break;
      }
      case 0x4026: case 0x4029: { const s = saved.get(r.u32()); if (s) ({ world, pageUnit, pageScale, clip } = { ...s, world: s.world.slice() }); break; } // Restore, EndContainer
      case 0x402a: world = r.m(); break;
      case 0x402b: world = ID.slice(); break;
      case 0x402c: { const m = r.m(); world = flags & 0x2000 ? mul(world, m) : mul(m, world); break; }
      case 0x402d: { const m = [1, 0, 0, 1, r.f32(), r.f32()]; world = flags & 0x2000 ? mul(world, m) : mul(m, world); break; }
      case 0x402e: { const sx = r.f32(), sy = r.f32(); const m = [sx, 0, 0, sy, 0, 0]; world = flags & 0x2000 ? mul(world, m) : mul(m, world); break; }
      case 0x402f: { const a = (r.f32() * Math.PI) / 180; const m = [Math.cos(a), Math.sin(a), -Math.sin(a), Math.cos(a), 0, 0]; world = flags & 0x2000 ? mul(world, m) : mul(m, world); break; }
      case 0x4030: pageUnit = lo; pageScale = r.f32() || 1; break; // SetPageTransform
      case 0x4031: clip = null; break; // ResetClip
      case 0x4032: setClip(rectD([r.f32(), r.f32(), r.f32(), r.f32()]), (flags >> 8) & 15); break; // SetClipRect
      case 0x4033: { const p = objects[lo]; if (p && p.kind === 'path') setClip(p.d, (flags >> 8) & 15); break; } // SetClipPath
      case 0x4034: { const reg = objects[lo]; if (reg && reg.kind === 'region') setClip(regionD(reg.node), (flags >> 8) & 15, reg.node.infinite); break; } // SetClipRegion
      case 0x4035: break; // OffsetClip (rare)
      case 0x4038: { // StrokeFillPath
        const p = objects[lo]; if (p && p.kind === 'path') shape(p.d, fillAttr(null)); break;
      }
      default: if (type >= 0x4009 && type <= 0x4040) unsupported = true;
    }
  };

  for (let off = 0; off + 8 <= bytes.length;) {
    const type = dv.getUint32(off, true), size = dv.getUint32(off + 4, true);
    if (size < 8 || off + size > bytes.length) break;
    if (type === 70 && size >= 16 && dv.getUint32(off + 12, true) === 0x2b464d45) {
      if (gdiFrom >= 0) { gdi.push([gdiFrom, off]); gdiFrom = -1; }
      const end = off + 12 + dv.getUint32(off + 8, true);
      for (let p = off + 16, sub = 0; p + 12 <= end; sub++) {
        const t = dv.getUint16(p, true), fl = dv.getUint16(p + 2, true), sz = dv.getUint32(p + 4, true), ds = dv.getUint32(p + 8, true);
        if (sz < 12 || p + sz > end) break;
        at = off + sub / 1000;
        try { record(t, fl, rd(p + 12), p + 12 + ds); } catch { unsupported = true; }
        p += sz;
      }
      if (gdiFrom === -2) gdiFrom = off + size;
    }
    if (type === 14) break;
    off += size;
  }
  if (gdiFrom >= 0) gdi.push([gdiFrom, bytes.length]);
  return { items, defs, gdi, unsupported };
}

/** The pixel size of a PNG, GIF, BMP or JPEG file, or [0, 0]. */
function imageSize(b) {
  const be = (o) => (b[o] << 8) | b[o + 1];
  if (b[0] === 0x89 && b.length > 24) return [(be(16) << 16) | be(18), (be(20) << 16) | be(22)];
  if (b[0] === 0x47 && b.length > 10) return [b[6] | (b[7] << 8), b[8] | (b[9] << 8)];
  if (b[0] === 0x42 && b.length > 26) { const dv = new DataView(b.buffer, b.byteOffset); return [dv.getInt32(18, true), Math.abs(dv.getInt32(22, true))]; }
  if (b[0] === 0xff) { // JPEG: the first start-of-frame marker
    for (let o = 2; o + 9 < b.length;) {
      if (b[o] !== 0xff) { o++; continue; }
      const m = b[o + 1];
      if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return [be(o + 7), be(o + 5)];
      o += 2 + be(o + 2);
    }
  }
  return [0, 0];
}
