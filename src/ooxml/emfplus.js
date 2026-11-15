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
const ID = [1, 0, 0, 1, 0, 0];
const inverse = ([a, b, c, d, e, f]) => { const k = a * d - b * c || 1; return [d / k, -b / k, -c / k, a / k, (c * f - d * e) / k, (b * e - a * f) / k]; };
// Line spacing (em) as GDI+ steps lines, its line gap included (FontFamily.GetLineSpacing / em height).
const LINE = { arial: 1.1499, 'times new roman': 1.1499, calibri: 1.2207, 'calibri light': 1.2207, cambria: 1.1724, 'segoe ui': 1.3301,
  tahoma: 1.2075, verdana: 1.2153, 'courier new': 1.1328, consolas: 1.1704, georgia: 1.1362, aptos: 1.2 };

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
  // (The page unit starts as Display: a pixel of a display reference, 1/100 inch of a printer's.)
  let world = ID.slice(), pageUnit = 1, pageScale = 1, refDisplay = true;
  // A container's own coordinates: its transforms are set relative to `base` (the world as it began), and its clip
  // limits everything inside it (`cClip`), as GDI+ plays containers.
  let base = ID.slice();
  let clip = null; // { id (clipPath), d, M (device matrix), parent }
  let cClip = null;
  let interp = 0;
  const saved = new Map();
  let gdiFrom = -1;
  let multi = 0; // multi-format: 1 in its first section, 2 in a later one

  const unitPx = (u) => ({ 1: refDisplay ? 1 : dpiX / 100, 2: 1, 3: dpiX / 72, 4: dpiX, 5: dpiX / 300, 6: dpiX / 25.4 }[u] ?? 1);
  // A font's em in world units: World and Pixel sizes are world units (they follow the page transform, as GDI+ draws
  // them); a physical size (points, inches, mm) is that size whatever the page scale.
  const fontEm = (font) => (font.unit === 0 || font.unit === 2 ? font.em : (font.em * unitPx(font.unit)) / k());
  const k = () => unitPx(pageUnit) * pageScale;
  const toDevice = () => { const s = k(); return [world[0] * s, world[1] * s, world[2] * s, world[3] * s, world[4] * s - bx, world[5] * s - by]; };
  const mat = (M) => `matrix(${M.map((v) => +v.toPrecision(7)).join(' ')})`; // (significant digits: EMU-scaled worlds are tiny)
  const tf = () => ` transform="${mat(toDevice())}"`;
  let at = 0;
  const emit = (s) => items.push({ at, s: clip ? `<g clip-path="url(#${clip.id})">${s}</g>` : s });

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
          else if (bpp === 16) {
            const v = dv.getUint16(row + x * 2, true), f5 = (q) => Math.round((q * 255) / 31);
            if (format === 0x00021006) c = 0xff000000 | (f5(v >> 11) << 16) | (Math.round((((v >> 5) & 63) * 255) / 63) << 8) | f5(v & 31); // RGB565
            else if (format === 0x00021005 || format === 0x00061007) c = ((format === 0x00061007 && !(v & 0x8000) ? 0 : 255) << 24 | (f5((v >> 10) & 31) << 16) | (f5((v >> 5) & 31) << 8) | f5(v & 31)) >>> 0; // RGB555, ARGB1555
            else if (format === 0x00101004) { const g = v >> 8; c = (0xff000000 | (g << 16) | (g << 8) | g) >>> 0; } // 16-bit grey
            else return null;
          }
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
      const inner = bytes.subarray(r.o, Math.min(end, r.o + size));
      const svg = h.emfToSvg(inner);
      if (!svg) return null;
      // (Its size from its frame, in its reference's pixels: the drawing's rounded bounds came out a pixel short.)
      const iv = new DataView(inner.buffer, inner.byteOffset, inner.byteLength);
      const pxm = iv.getInt32(72, true) / (iv.getInt32(80, true) || 1);
      const fw = ((iv.getInt32(32, true) - iv.getInt32(24, true)) / 100) * pxm, fh = ((iv.getInt32(36, true) - iv.getInt32(28, true)) / 100) * (iv.getInt32(76, true) / (iv.getInt32(84, true) || 1));
      const m = /viewBox="0 0 ([\d.]+) ([\d.]+)"/.exec(svg);
      return { kind: 'image', w: fw > 0 ? fw : m ? +m[1] : 1, h: fh > 0 ? fh : m ? +m[2] : 1, href: `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(svg)))}` };
    }
    return null;
  };
  const readBrush = (r, end) => {
    r.u32(); // version
    const type = r.u32();
    if (type === 0) return { kind: 'brush', type: 'solid', color: r.u32() };
    if (type === 1) return { kind: 'brush', type: 'hatch', style: r.u32(), fore: r.u32(), back: r.u32() };
    if (type === 2) { // texture
      const flags = r.u32(), wrap = r.u32();
      const m = flags & 2 ? r.m() : null;
      const img = readImage(r, end);
      return img ? { kind: 'brush', type: 'texture', img, m, wrap } : { kind: 'brush', type: 'solid', color: 0x80808080 };
    }
    if (type === 4) { // linear gradient
      const flags = r.u32(), wrap = r.u32();
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
      return { kind: 'brush', type: 'linear', rect, m, stops, wrap };
    }
    if (type === 3) { // path gradient: from the centre colour out to the boundary's
      const flags = r.u32(); r.u32();
      const center = r.u32(), cx = r.f32(), cy = r.f32();
      const cnt = r.u32(); const surround = []; for (let i = 0; i < cnt; i++) surround.push(r.u32());
      let pts = [], curved = false;
      if (flags & 1) {
        const size = r.u32(), start = r.o;
        const p = readPath(r);
        r.o = start + size;
        curved = /C/.test(p.d);
        pts = (p.d.match(/-?[\d.]+ -?[\d.]+/g) || []).map((q) => q.split(' ').map(Number));
      } else { const pc = r.u32(); pts = points(r, pc, 0); }
      const m = flags & 2 ? r.m() : null;
      const xs = pts.map((q) => q[0]), ys = pts.map((q) => q[1]);
      const box = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
      // A curved boundary (an ellipse): elliptical, its colour reached at the boundary all round. Straight sides: a fan
      // of triangles from the centre, each corner its own colour (the last one given for the corners after it).
      if (curved || pts.length < 3) return { kind: 'brush', type: 'radial', cx, cy, rx: Math.max(1, (box[2] - box[0]) / 2), ry: Math.max(1, (box[3] - box[1]) / 2), c1: center, c2: surround[0] ?? center, m };
      return { kind: 'brush', type: 'fan', cx, cy, pts, box, c1: center, cols: pts.map((_, i) => surround[Math.min(i, cnt - 1)] ?? center), m };
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
    const pen = { kind: 'pen', unit, width, style: 0, cap: 0, join: 0, miter: 10, dashes: null, dashOffset: 0 };
    if (flags & 0x01) pen.m = r.m();
    if (flags & 0x02) pen.cap = r.i32();
    pen.endCap = flags & 0x04 ? r.i32() : pen.cap;
    if (flags & 0x08) pen.join = r.i32();
    if (flags & 0x10) pen.miter = r.f32();
    if (flags & 0x20) pen.style = r.i32();
    if (flags & 0x40) r.i32();
    if (flags & 0x80) pen.dashOffset = r.f32();
    if (flags & 0x100) { const c = r.u32(); pen.dashes = []; for (let i = 0; i < c; i++) pen.dashes.push(r.f32()); }
    if (pen.dashes && !(flags & 0x20)) pen.style = 5; // (GDI+ writes a custom pattern with no line style)
    if (flags & 0x200) pen.inset = r.i32() === 1;
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
    const align = r.u32(), lineAlign = r.u32();
    r.u32(); r.u32(); r.f32(); r.i32(); // digit substitution and language, first tab offset, hotkey prefix
    return { kind: 'format', flags, align, lineAlign, lead: r.f32(), trail: r.f32() };
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
      // An 8 x 8 tile of the picture's pixels, upright and unscaled whatever the world and page (as GDI+ draws them):
      // its lines, or for the percent styles that share of its pixels.
      const g = id('h');
      const lines = { 0: 'M0 4.5h8', 1: 'M4.5 0v8', 2: 'M0 0l8 8', 3: 'M8 0l-8 8', 4: 'M0 4.5h8M4.5 0v8', 5: 'M0 0l8 8M8 0l-8 8' }[brush.style];
      const pct = { 6: 5, 7: 10, 8: 20, 9: 25, 10: 30, 11: 40, 12: 50, 13: 60, 14: 70, 15: 75, 16: 80, 17: 90 }[brush.style];
      // (The other named styles: their lines at their spacing and weight, or a rule for which pixels are inked.)
      const down = 'M-4 -4l16 16M-4 4l16 16M-4 -12l16 16', up = 'M12 -4l-16 16M12 4l-16 16M12 -12l-16 16', wide = (d) => [d, 3];
      const more = {
        18: [down, 1], 19: [up, 1], 20: [down, 2], 21: [up, 2], 22: wide('M0 0l8 8M-8 0l16 16M0 -8l16 16'), 23: wide('M8 0l-8 8M16 0l-16 16M8 -8l-16 16'),
        24: ['M0.5 0v8M4.5 0v8', 1], 25: ['M0 0.5h8M0 4.5h8', 1], 26: ['M0.5 0v8M2.5 0v8M4.5 0v8M6.5 0v8', 1], 27: ['M0 0.5h8M0 2.5h8M0 4.5h8M0 6.5h8', 1],
        28: ['M1 0v8M5 0v8', 2], 29: ['M0 1h8M0 5h8', 2], 30: ['M0 0l4 4M4 0l4 4', 1], 31: ['M4 0l-4 4M8 0l-4 4', 1], 32: ['M0 0.5h4M4 4.5h4', 1],
        33: ['M0.5 0v4M4.5 4v4', 1], 36: ['M0 4l4-4l4 4', 1], 37: ['M0 4q2-3 4 0t4 0', 1], 38: ['M0 8l8-8M4 4l4 4', 1], 39: ['M0 0.5h8M0.5 0v4M0 4.5h8M4.5 4v4', 1],
        46: ['M0 0l8 8M8 0l-8 8M0 4h8', 1], 48: ['M0 0.5h8M0 4.5h8M0.5 0v8M4.5 0v8', 1], 51: ['M0 0l8 8M8 0l-8 8', 1],
      }[brush.style];
      const dots = {
        34: (x, y) => (x * 5 + y * 3) % 8 === 0, 35: (x, y) => ((x >> 1) * 5 + (y >> 1) * 3) % 4 === 0, 43: (x, y) => (x === 0 && y % 2 === 0) || (y === 0 && x % 2 === 0),
        44: (x, y) => (x + y) % 4 === 0 && (x - y) % 4 === 0, 49: (x, y) => ((x >> 1) + (y >> 1)) % 2 === 0, 50: (x, y) => ((x >> 2) + (y >> 2)) % 2 === 0,
        52: (x, y) => Math.abs(x - 3.5) + Math.abs(y - 3.5) <= 4,
        40: (x, y) => (x + y) % 4 === 0 || (x - y + 8) % 8 === 0, 41: (x, y) => y < 4 ? (x + y) % 2 === 0 : x < 4 && y < 6, 42: (x, y) => (x === 2 && y === 1) || (x === 3 && y === 2) || (x === 6 && y === 5) || (x === 5 && y === 6),
        45: (x, y) => y === 0 || (y < 4 && x === 0) || (y === 4) || (y > 4 && x === 4), 47: (x, y) => x % 2 === 0 && y % 2 === 0,
      }[brush.style];
      let ink = '';
      const cellsOf = (on) => { const c = []; for (let y = 0; y < 8; y++) for (let x = 0; x < 8; x++) if (on(x, y)) c.push(`M${x} ${y}h1v1h-1Z`); return `<path d="${c.join('')}" fill="${hex(brush.fore)}" fill-opacity="${n(alpha(brush.fore))}"/>`; };
      if (lines || more) ink = `<path d="${lines || more[0]}" fill="none" stroke="${hex(brush.fore)}" stroke-opacity="${n(alpha(brush.fore))}" stroke-width="${lines ? 1 : more[1]}"/>`;
      else if (pct) ink = cellsOf((x, y) => (pct === 50 ? (x + y) % 2 === 0 : ((x * 37 + y * 21) % 64) < (pct * 64) / 100));
      else if (dots) ink = cellsOf(dots);
      else return [hex(lerp(brush.fore, brush.back, 0.5)), 1];
      defs.push(`<pattern id="${g}" width="8" height="8" patternUnits="userSpaceOnUse" patternTransform="${mat(inverse(toDevice()))}" shape-rendering="crispEdges" overflow="hidden"><rect width="8" height="8" fill="${hex(brush.back)}" fill-opacity="${n(alpha(brush.back))}"/>${ink}</pattern>`);
      return [`url(#${g})`, 1];
    }
    if (brush.type === 'texture') {
      // (Tiled; mirrored at every other tile for the flip modes, a tile of two by two. Clamped as tiled: the same over the
      // picture itself, which is where such a brush fills, and a pattern too big for its scale draws nothing.)
      const g = id('t');
      const { w, h: ih, href } = brush.img;
      const fx = brush.wrap === 1 || brush.wrap === 3, fy = brush.wrap === 2 || brush.wrap === 3;
      const tw = fx ? 2 * w : w, th = fy ? 2 * ih : ih;
      const im = (sx, sy) => `<image width="${w}" height="${ih}" href="${href}" transform="matrix(${sx} 0 0 ${sy} ${sx < 0 ? 2 * w : 0} ${sy < 0 ? 2 * ih : 0})"/>`;
      const tiles = [im(1, 1), fx && im(-1, 1), fy && im(1, -1), fx && fy && im(-1, -1)].filter(Boolean).join('');
      defs.push(`<pattern id="${g}" width="${tw}" height="${th}" patternUnits="userSpaceOnUse"${brush.m ? ` patternTransform="${mat(brush.m)}"` : ''}>${tiles}</pattern>`);
      return [`url(#${g})`, 1];
    }
    if (brush.type === 'linear') {
      const g = id('g');
      const [x, y, w, hh] = brush.rect;
      // (Across its rectangle, through the brush's matrix; repeated past it as GDI+ tiles by default, mirrored, or
      // padded only when clamped.)
      const spread = brush.wrap === 4 ? '' : ` spreadMethod="${brush.wrap === 0 ? 'repeat' : 'reflect'}"`;
      const stops = brush.stops.map(([o, c]) => `<stop offset="${n(o)}" stop-color="${hex(c)}" stop-opacity="${n(alpha(c))}"/>`).join('');
      defs.push(`<linearGradient id="${g}" gradientUnits="userSpaceOnUse" x1="${n(x)}" y1="${n(y + hh / 2)}" x2="${n(x + w)}" y2="${n(y + hh / 2)}"${spread}${brush.m ? ` gradientTransform="${mat(brush.m)}"` : ''}>${stops}</linearGradient>`);
      return [`url(#${g})`, 1];
    }
    if (brush.type === 'radial') {
      const g = id('g');
      const { cx, cy, rx, ry } = brush;
      const squash = mul(mul([1, 0, 0, 1, -cx, -cy], [1, 0, 0, ry / rx, 0, 0]), [1, 0, 0, 1, cx, cy]);
      defs.push(`<radialGradient id="${g}" gradientUnits="userSpaceOnUse" cx="${n(cx)}" cy="${n(cy)}" r="${n(rx)}" gradientTransform="${mat(brush.m ? mul(squash, brush.m) : squash)}"><stop offset="0" stop-color="${hex(brush.c1)}" stop-opacity="${n(alpha(brush.c1))}"/><stop offset="1" stop-color="${hex(brush.c2)}" stop-opacity="${n(alpha(brush.c2))}"/></radialGradient>`);
      return [`url(#${g})`, 1];
    }
    if (brush.type === 'fan') {
      // (A pattern the size of the boundary holding the triangles. Each is the centre colour, under its side's colours
      // (one corner's to the next's, along the side) faded in from the centre out to the side: a corner gets its own
      // colour, and the colours meet along the lines to the corners, as GDI+ blends them.)
      const g = id('f');
      const { cx, cy, pts, box, c1, cols } = brush;
      const stop = (o, c) => `<stop offset="${o}" stop-color="${hex(c)}" stop-opacity="${n(alpha(c))}"/>`;
      const tris = pts.map((a, i) => {
        const b = pts[(i + 1) % pts.length];
        const ex = b[0] - a[0], ey = b[1] - a[1], len2 = ex * ex + ey * ey || 1;
        const t = ((cx - a[0]) * ex + (cy - a[1]) * ey) / len2; // the foot of the perpendicular from the centre
        const fx = a[0] + t * ex, fy = a[1] + t * ey;
        const side = id('g'), fade = id('g'), mask = id('k');
        const d = `M${n(cx)} ${n(cy)}L${n(a[0])} ${n(a[1])}L${n(b[0])} ${n(b[1])}Z`;
        defs.push(`<linearGradient id="${side}" gradientUnits="userSpaceOnUse" x1="${n(a[0])}" y1="${n(a[1])}" x2="${n(b[0])}" y2="${n(b[1])}">${stop(0, cols[i])}${stop(1, cols[(i + 1) % pts.length])}</linearGradient>`);
        defs.push(`<linearGradient id="${fade}" gradientUnits="userSpaceOnUse" x1="${n(cx)}" y1="${n(cy)}" x2="${n(fx)}" y2="${n(fy)}"><stop offset="0" stop-color="#000"/><stop offset="1" stop-color="#fff"/></linearGradient>`);
        defs.push(`<mask id="${mask}" maskUnits="userSpaceOnUse" x="-1e6" y="-1e6" width="2e6" height="2e6"><path d="${d}" fill="url(#${fade})" stroke="url(#${fade})" stroke-width="0.5"/></mask>`);
        return `<path d="${d}" fill="${hex(c1)}" fill-opacity="${n(alpha(c1))}" stroke="${hex(c1)}" stroke-width="0.5"/><path d="${d}" fill="url(#${side})" stroke="url(#${side})" stroke-width="0.5" mask="url(#${mask})"/>`;
      }).join('');
      defs.push(`<pattern id="${g}" x="${n(box[0])}" y="${n(box[1])}" width="${n(box[2] - box[0])}" height="${n(box[3] - box[1])}" patternUnits="userSpaceOnUse"${brush.m ? ` patternTransform="${mat(brush.m)}"` : ''}><g transform="translate(${n(-box[0])} ${n(-box[1])})">${tris}</g></pattern>`);
      return [`url(#${g})`, 1];
    }
    return ['#808080', 1];
  };
  const brushOf = (flags, v) => (flags & 0x8000 ? v : objects[v & 255]);
  const fillAttr = (brush, rule = 'evenodd') => { const [p, o] = paint(brush); return `fill="${p}"${o < 1 ? ` fill-opacity="${n(o)}"` : ''} fill-rule="${rule}" stroke="none"`; };
  const strokeAttr = (pen, widen = 1) => {
    if (!pen || pen.kind !== 'pen') return 'stroke="#000" stroke-width="1" vector-effect="non-scaling-stroke"';
    const [p, o] = paint(pen.brush);
    let w = pen.width;
    if (pen.unit && pen.unit !== pageUnit && pen.unit > 1) w = (w * unitPx(pen.unit)) / k();
    const thin = !(w > 0);
    const dash = pen.style === 5 && pen.dashes ? pen.dashes : { 1: [3, 1], 2: [1, 1], 3: [3, 1, 1, 1], 4: [3, 1, 1, 1, 1, 1] }[pen.style];
    const dw = (thin ? 1 : w) * widen;
    // (A different cap at each end: the line ends flat, each end's cap a marker, as SVG has one line cap.)
    const capName = (c) => ['butt', 'square', 'round'][c] || 'butt';
    const end = pen.endCap ?? pen.cap;
    let caps = ` stroke-linecap="${capName(pen.cap)}"`;
    if (end !== pen.cap) {
      const marker = (c) => {
        if (c !== 1 && c !== 2) return null;
        const mid = id('m');
        const shapeEl = c === 2 ? '<circle r="0.5"/>' : '<rect x="0" y="-0.5" width="0.5" height="1"/>';
        defs.push(`<marker id="${mid}" markerUnits="strokeWidth" markerWidth="1" markerHeight="1" viewBox="-0.5 -0.5 1 1" orient="auto-start-reverse" overflow="visible"><g fill="${p}"${o < 1 ? ` fill-opacity="${n(o)}"` : ''}>${shapeEl}</g></marker>`);
        return mid;
      };
      const ms = marker(pen.cap), me = marker(end);
      caps = ` stroke-linecap="butt"${ms ? ` marker-start="url(#${ms})"` : ''}${me ? ` marker-end="url(#${me})"` : ''}`;
    }
    return `fill="none" stroke="${p}"${o < 1 ? ` stroke-opacity="${n(o)}"` : ''} stroke-width="${n(dw)}"${thin ? ' vector-effect="non-scaling-stroke"' : ''}`
      + `${dash ? ` stroke-dasharray="${dash.map((v) => n(v * dw)).join(' ')}"${pen.dashOffset ? ` stroke-dashoffset="${n(pen.dashOffset * dw)}"` : ''}` : ''}`
      + `${caps} stroke-linejoin="${['miter', 'bevel', 'round', 'miter'][pen.join] || 'miter'}" stroke-miterlimit="${n(Math.max(1, pen.miter))}"`;
  };
  // A closed figure's outline: an inset pen inside it (twice as wide, clipped to the figure), as GDI+ draws it.
  // A pen with a transform of its own (thicker one way than another): the figure taken through the transform's inverse,
  // stroked there, and drawn through the transform, so the pen's width turns and stretches with it.
  const mapD = (d, M) => {
    let cx = 0, cy = 0, out = '';
    const pt = (x, y) => { const q = [M[0] * x + M[2] * y + M[4], M[1] * x + M[3] * y + M[5]]; return `${n(q[0])} ${n(q[1])}`; };
    for (const [, c, args] of d.matchAll(/([MLCZHVhv])([^MLCZHVhv]*)/g)) {
      const a = (args.match(/-?[\d.]+(?:e-?\d+)?/g) || []).map(Number);
      if (c === 'Z') { out += 'Z'; continue; }
      if (c === 'H' || c === 'h') { cx = c === 'H' ? a[0] : cx + a[0]; out += `L${pt(cx, cy)}`; continue; }
      if (c === 'V' || c === 'v') { cy = c === 'V' ? a[0] : cy + a[0]; out += `L${pt(cx, cy)}`; continue; }
      const ps = [];
      for (let i = 0; i + 1 < a.length; i += 2) { ps.push(pt(a[i], a[i + 1])); cx = a[i]; cy = a[i + 1]; }
      out += c + ps.join(' ');
    }
    return out;
  };
  const outline = (d, pen, closed = true, plain = null) => {
    if (!d) return;
    if (pen && pen.kind === 'pen' && pen.m && !pen.inset && plain !== false) {
      const src = plain || d;
      if (!/[aA]/.test(src)) {
        const Pm = [pen.m[0], pen.m[1], pen.m[2], pen.m[3], 0, 0];
        emit(`<path d="${mapD(src, inverse(Pm))}" ${strokeAttr(pen)} transform="${mat(mul(Pm, toDevice()))}"/>`);
        return;
      }
    }
    if (!(closed && pen && pen.kind === 'pen' && pen.inset)) { shape(d, strokeAttr(pen)); return; }
    const was = clip;
    clip = clipPath(d, toDevice(), clip);
    shape(d, strokeAttr(pen, 2));
    clip = was;
  };
  const rectD = ([x, y, w, hh]) => `M${n(x)} ${n(y)}h${n(w)}v${n(hh)}h${n(-w)}Z`;
  // An ellipse as four beziers (for a figure that has to go through a transform as points).
  const ellipseBez = ([x, y, w, hh]) => {
    const cx = x + w / 2, cy = y + hh / 2, rx = w / 2, ry = hh / 2, k = 0.5522847;
    return `M${n(cx + rx)} ${n(cy)}C${n(cx + rx)} ${n(cy + k * ry)} ${n(cx + k * rx)} ${n(cy + ry)} ${n(cx)} ${n(cy + ry)}C${n(cx - k * rx)} ${n(cy + ry)} ${n(cx - rx)} ${n(cy + k * ry)} ${n(cx - rx)} ${n(cy)}`
      + `C${n(cx - rx)} ${n(cy - k * ry)} ${n(cx - k * rx)} ${n(cy - ry)} ${n(cx)} ${n(cy - ry)}C${n(cx + k * rx)} ${n(cy - ry)} ${n(cx + rx)} ${n(cy - k * ry)} ${n(cx + rx)} ${n(cy)}Z`;
  };
  const ellipseD = ([x, y, w, hh]) => `M${n(x)} ${n(y + hh / 2)}a${n(w / 2)} ${n(hh / 2)} 0 1 0 ${n(w)} 0a${n(w / 2)} ${n(hh / 2)} 0 1 0 ${n(-w)} 0Z`;
  const arcD = ([x, y, w, hh], start, sweep, pie) => {
    const cx = x + w / 2, cy = y + hh / 2, rx = w / 2, ry = hh / 2;
    if (Math.abs(sweep) >= 360) return ellipseD([x, y, w, hh]);
    // (GDI+ takes the angles along rays from the centre, not as the ellipse's parameter: its 45 degrees on a 2:1
    // ellipse is where the 45-degree ray meets it.)
    const param = (deg) => { const a = (deg * Math.PI) / 180; return Math.atan2(rx * Math.sin(a), ry * Math.cos(a)); };
    const t1 = param(start);
    let t2 = param(start + sweep);
    if (sweep > 0) while (t2 <= t1) t2 += 2 * Math.PI;
    else while (t2 >= t1) t2 -= 2 * Math.PI;
    const [x1, y1] = [cx + rx * Math.cos(t1), cy + ry * Math.sin(t1)], [x2, y2] = [cx + rx * Math.cos(t2), cy + ry * Math.sin(t2)];
    const arc = `A${n(rx)} ${n(ry)} 0 ${Math.abs(t2 - t1) > Math.PI ? 1 : 0} ${sweep > 0 ? 1 : 0} ${n(x2)} ${n(y2)}`;
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
  // before it for an intersection. Inside a container a new clip still lies within the container's.
  const clipPath = (d, M, parent) => {
    const c = { id: id('p'), d, M, parent };
    defs.push(`<clipPath id="${c.id}"${parent ? ` clip-path="url(#${parent.id})"` : ''}><path d="${d}" clip-rule="nonzero" transform="${mat(M)}"/></clipPath>`);
    return c;
  };
  // A region as clip paths: a shape, shapes joined (union), or one inside another (intersection). (Exclusion, xor and
  // complement keep their first part.)
  const regionClip = (node, parent) => {
    if (!node || node.infinite) return parent;
    if (node.d !== undefined) return clipPath(node.d || 'M0 0Z', toDevice(), parent);
    if (node.op === 1) return regionClip(node.a, regionClip(node.b, parent));
    if (node.op === 2 && node.a.d !== undefined && node.b.d !== undefined) {
      if (node.a.infinite || node.b.infinite) return parent;
      return clipPath(`${node.a.d} ${node.b.d}`, toDevice(), parent);
    }
    return regionClip(node.a, parent);
  };
  const setClip = (d, mode, infinite) => {
    if (mode !== 0 && mode !== 1) return; // union, xor, exclude, complement: the clip is left as it was
    const parent = mode === 1 ? clip : cClip;
    clip = infinite ? parent : clipPath(d, toDevice(), parent);
  };
  const setClipRegion = (node, mode) => {
    if (mode !== 0 && mode !== 1) return;
    clip = regionClip(node, mode === 1 ? clip : cClip);
  };
  const fillRegion = (node, attrs) => { // a region filled: its shapes, an intersection drawn clipped to its other part
    if (!node) return;
    if (node.d !== undefined) { shape(node.d, attrs); return; }
    if (node.op === 2) { fillRegion(node.a, attrs); fillRegion(node.b, attrs); return; }
    if (node.op === 1) {
      const was = clip;
      clip = regionClip(node.b, clip);
      fillRegion(node.a, attrs);
      clip = was;
      return;
    }
    fillRegion(node.a, attrs);
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
    const em = fontEm(font);
    const [asc, desc] = h.win(font.face);
    const [x, y, w, hh] = layout;
    // (The format's margins: 1/6 em before and after the text by default, none in GenericTypographic's.)
    const lead = (fmt ? fmt.lead : 1 / 6) * em, trail = (fmt ? fmt.trail : 1 / 6) * em;
    const align = fmt ? fmt.align : 0, valign = fmt ? fmt.lineAlign : 0;
    const flags = fmt ? fmt.flags : 0;
    // Its lines: at its breaks, and, in a layout box of a width, wrapped at spaces to fit it (unless NoWrap).
    let lines = str.replace(/\r\n?/g, '\n').split('\n');
    const room = w - lead - trail;
    if (w > 0 && !(flags & 0x1000) && h.measure) {
      const css = `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${em}px ${h.fontStack(font.face)}`;
      lines = lines.flatMap((line) => {
        const out = [];
        let cur = '';
        for (const word of line.split(/(?<= )/)) {
          if (cur && h.measure(cur + word.trimEnd(), css) > room + 0.5) { out.push(cur.trimEnd()); cur = word; } else cur += word;
        }
        out.push(cur.trimEnd());
        return out;
      });
    }
    // (NoWrap in a box it's cut off at: the characters that don't wholly fit are left out, as GDI+'s character trimming.)
    if (w > 0 && hh > 0 && flags & 0x1000 && !(flags & 0x4000) && h.measure) {
      const css = `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${em}px ${h.fontStack(font.face)}`;
      lines = lines.map((line) => { let t = line; while (t && h.measure(t, css) > room + 0.5) t = t.slice(0, -1); return t; });
    }
    const lh = (LINE[font.face.toLowerCase()] ?? asc + desc) * em;
    // (Drawn at a point, a layout of no size: centred or ended on the point. A centred or bottom block is a little
    // taller than its lines, as GDI+ measures it.)
    const span = hh > 0 ? hh : 0;
    const block = lh * lines.length + (valign ? 0.09 * em : 0);
    const top = valign === 1 ? y + (span - block) / 2 : valign === 2 ? y + span - block : y;
    const anchor = align === 1 ? 'middle' : align === 2 ? 'end' : 'start';
    const tx = anchor === 'middle' ? x + w / 2 : anchor === 'end' ? x + w - trail : x + lead;
    const [p, o] = paint(brush);
    const deco = [font.underline && 'underline', font.strike && 'line-through'].filter(Boolean).join(' ');
    const spans = lines.map((s, i) => `<tspan x="${n(tx)}" y="${n(top + asc * em + i * lh)}">${h.esc(s)}</tspan>`).join('');
    // (Cut off at its layout box, unless NoClip.)
    const was = clip;
    if (w > 0 && hh > 0 && !(flags & 0x4000)) clip = clipPath(rectD([x, y, w, hh]), toDevice(), clip);
    emit(`<text font-family="${h.esc(h.fontStack(font.face))}" font-size="${n(em)}"${font.bold ? ' font-weight="bold"' : ''}${font.italic ? ' font-style="italic"' : ''}${deco ? ` text-decoration="${deco}"` : ''} fill="${p}"${o < 1 ? ` fill-opacity="${n(o)}"` : ''} text-anchor="${anchor}" xml:space="preserve"${tf()}>${spans}</text>`);
    clip = was;
  };

  // ----- the records -----
  const record = (type, flags, r, end) => {
    const lo = flags & 255;
    switch (type) {
      case 0x4001: r.u32(); refDisplay = !!(r.u32() & 1); dpiX = r.u32() || 96; dpiY = r.u32() || 96; void dpiY; break; // header
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
      case 0x400b: { const c = r.u32(); for (const q of rects(r, c, flags & 0x4000)) outline(rectD(q), objects[lo]); break; } // DrawRects
      case 0x400c: { const b = r.u32(), c = r.u32(); shape(polyD(points(r, c, flags), true), fillAttr(brushOf(flags, b))); break; } // FillPolygon
      case 0x400d: { const c = r.u32(); outline(polyD(points(r, c, flags), flags & 0x2000), objects[lo], !!(flags & 0x2000)); break; } // DrawLines
      case 0x400e: { const b = r.u32(); shape(ellipseD(rects(r, 1, flags & 0x4000)[0]), fillAttr(brushOf(flags, b))); break; } // FillEllipse
      case 0x400f: { const q = rects(r, 1, flags & 0x4000)[0]; outline(ellipseD(q), objects[lo], true, ellipseBez(q)); break; } // DrawEllipse
      case 0x4010: { const b = r.u32(), s = r.f32(), w = r.f32(); shape(arcD(rects(r, 1, flags & 0x4000)[0], s, w, true), fillAttr(brushOf(flags, b))); break; } // FillPie
      case 0x4011: case 0x4012: { const s = r.f32(), w = r.f32(); shape(arcD(rects(r, 1, flags & 0x4000)[0], s, w, type === 0x4011), strokeAttr(objects[lo])); break; } // DrawPie, DrawArc
      case 0x4013: { const b = r.u32(); const reg = objects[lo]; if (reg && reg.kind === 'region') fillRegion(reg.node, fillAttr(brushOf(flags, b), 'nonzero')); break; } // FillRegion
      case 0x4014: { const b = r.u32(); const p = objects[lo]; if (p && p.kind === 'path') shape(p.d, fillAttr(brushOf(flags, b))); break; } // FillPath
      case 0x4015: { const pen = objects[r.u32() & 255]; const p = objects[lo]; if (p && p.kind === 'path') outline(p.d, pen, /Z$/.test(p.d)); break; } // DrawPath
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
        const b = r.u32(), opts = r.u32(), hasMatrix = r.u32();
        const c = r.u32();
        if (!(opts & 1)) { unsupported = true; break; }
        let s = ''; for (let i = 0; i < c; i++) s += String.fromCharCode(r.u16());
        const pts = []; for (let i = 0; i < c; i++) pts.push([r.f32(), r.f32()]);
        const m = hasMatrix ? r.m() : null; // (turned or skewed text: charts' labels)
        const font = objects[lo];
        if (!font || font.kind !== 'font' || !c) break;
        const em = fontEm(font);
        const [p, o] = paint(brushOf(flags, b));
        // (With realized advances only the first position counts: the font's own advances place the rest.)
        const xy = opts & 4 ? [pts[0]] : pts;
        emit(`<text font-family="${h.esc(h.fontStack(font.face))}" font-size="${n(em)}"${font.bold ? ' font-weight="bold"' : ''}${font.italic ? ' font-style="italic"' : ''} fill="${p}"${o < 1 ? ` fill-opacity="${n(o)}"` : ''} xml:space="preserve" x="${xy.map(([x]) => n(x)).join(' ')}" y="${xy.map(([, y]) => n(y)).join(' ')}" transform="${mat(m ? mul(m, toDevice()) : toDevice())}">${h.esc(s)}</text>`);
        break;
      }
      case 0x4021: interp = lo; break; // SetInterpolationMode
      case 0x401d: case 0x401e: case 0x401f: case 0x4020: case 0x4022: case 0x4023: case 0x4024: break; // rendering hints
      case 0x4025: saved.set(r.u32(), { world: world.slice(), base, pageUnit, pageScale, clip, cClip }); break; // Save
      case 0x4028: // BeginContainerNoParams: a container in the same coordinates
        saved.set(r.u32(), { world: world.slice(), base, pageUnit, pageScale, clip, cClip });
        base = world.slice(); cClip = clip;
        break;
      case 0x4027: { // BeginContainer: its source rectangle (in the record's unit) mapped onto the destination (world units)
        const dst = [r.f32(), r.f32(), r.f32(), r.f32()], src = [r.f32(), r.f32(), r.f32(), r.f32()].map((v) => (v * unitPx(lo)) / k());
        saved.set(r.u32(), { world: world.slice(), base, pageUnit, pageScale, clip, cClip });
        const sx = src[2] ? dst[2] / src[2] : 1, sy = src[3] ? dst[3] / src[3] : 1;
        world = mul([sx, 0, 0, sy, dst[0] - src[0] * sx, dst[1] - src[1] * sy], world);
        base = world.slice(); cClip = clip;
        break;
      }
      case 0x4026: case 0x4029: { const s = saved.get(r.u32()); if (s) ({ world, base, pageUnit, pageScale, clip, cClip } = { ...s, world: s.world.slice() }); break; } // Restore, EndContainer
      case 0x402a: world = mul(r.m(), base); break; // (relative to the container)
      case 0x402b: world = base.slice(); break;
      case 0x402c: { const m = r.m(); world = flags & 0x2000 ? mul(world, m) : mul(m, world); break; }
      case 0x402d: { const m = [1, 0, 0, 1, r.f32(), r.f32()]; world = flags & 0x2000 ? mul(world, m) : mul(m, world); break; }
      case 0x402e: { const sx = r.f32(), sy = r.f32(); const m = [sx, 0, 0, sy, 0, 0]; world = flags & 0x2000 ? mul(world, m) : mul(m, world); break; }
      case 0x402f: { const a = (r.f32() * Math.PI) / 180; const m = [Math.cos(a), Math.sin(a), -Math.sin(a), Math.cos(a), 0, 0]; world = flags & 0x2000 ? mul(world, m) : mul(m, world); break; }
      case 0x4030: pageUnit = lo; pageScale = r.f32() || 1; break; // SetPageTransform
      case 0x4031: clip = cClip; break; // ResetClip (to the container's clip)
      case 0x4032: setClip(rectD([r.f32(), r.f32(), r.f32(), r.f32()]), (flags >> 8) & 15); break; // SetClipRect
      case 0x4033: { const p = objects[lo]; if (p && p.kind === 'path') setClip(p.d, (flags >> 8) & 15); break; } // SetClipPath
      case 0x4034: { const reg = objects[lo]; if (reg && reg.kind === 'region') setClipRegion(reg.node, (flags >> 8) & 15); break; } // SetClipRegion
      case 0x4035: { // OffsetClip: the clip moved by (dx, dy) world units
        const dx = r.f32(), dy = r.f32();
        if (clip && clip !== cClip) {
          const D = toDevice(); // (the move in device units, through the current world)
          const [a, b, c, d, e, f] = clip.M;
          clip = clipPath(clip.d, [a, b, c, d, e + D[0] * dx + D[2] * dy, f + D[1] * dx + D[3] * dy], clip.parent);
        }
        break;
      }
      case 0x4037: case 0x4038: case 0x4039: case 0x403a: break; // StrokeFillPath (GDI's current pen and brush), image effects, terminal-server records
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
        // (Multi-format sections hold the same drawing in other forms: the first is drawn, the others skipped.)
        if (t === 0x4005) multi = 1;
        else if (t === 0x4006) multi = 2;
        else if (t === 0x4007) multi = 0;
        else if (multi !== 2) { try { record(t, fl, rd(p + 12), p + 12 + ds); } catch { unsupported = true; } }
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
