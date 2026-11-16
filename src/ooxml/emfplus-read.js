// EMF+ records read (emfplus.js draws them): a little-endian reader over the bytes, and the objects the records define
// (brushes, pens, paths, regions, images, fonts, string formats). Each reader takes its bytes from `r`, so an object
// split over several records is read from the joined copy by a reader of its own. `h`: the picture helpers emfplus.js is
// given (png, dataUrl, emfToSvg) for the images.

export const n = (v) => +(+v).toFixed(3);

/** A little-endian reader over `bytes` from offset o: r.bytes and r.dv are what it reads. */
export const reader = (bytes) => {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return (o) => ({
    o,
    bytes,
    dv,
    u8() { return dv.getUint8(this.o++); },
    u16() { const v = dv.getUint16(this.o, true); this.o += 2; return v; },
    i16() { const v = dv.getInt16(this.o, true); this.o += 2; return v; },
    u32() { const v = dv.getUint32(this.o, true); this.o += 4; return v; },
    i32() { const v = dv.getInt32(this.o, true); this.o += 4; return v; },
    f32() { const v = dv.getFloat32(this.o, true); this.o += 4; return v; },
    m() { return [this.f32(), this.f32(), this.f32(), this.f32(), this.f32(), this.f32()]; },
  });
};
export const points = (r, count, flags) => {
  const pts = [];
  if (flags & 0x0800) { // relative, 7 or 15 bits each
    let x = 0, y = 0;
    const one = () => { const b = r.u8(); if (!(b & 0x80)) return b & 0x40 ? b - 0x80 : b; const v = ((b & 0x7f) << 8) | r.u8(); return v & 0x4000 ? v - 0x8000 : v; };
    for (let i = 0; i < count; i++) { x += one(); y += one(); pts.push([x, y]); }
  } else if (flags & 0x4000) for (let i = 0; i < count; i++) pts.push([r.i16(), r.i16()]);
  else for (let i = 0; i < count; i++) pts.push([r.f32(), r.f32()]);
  return pts;
};
export const rects = (r, count, compressed) => {
  const out = [];
  for (let i = 0; i < count; i++) out.push(compressed ? [r.i16(), r.i16(), r.i16(), r.i16()] : [r.f32(), r.f32(), r.f32(), r.f32()]);
  return out;
};
export const utf16 = (r, len) => { let s = ''; for (let i = 0; i < len; i++) s += String.fromCharCode(r.u16()); return s; };

// ----- objects -----
export const readPath = (r) => {
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
export const readImage = (r, end, h) => {
  r.u32(); // version
  const type = r.u32();
  if (type === 1) { // bitmap
    const w = r.i32(), hh = r.i32(), stride = r.i32(), format = r.u32(), compressed = r.u32();
    if (compressed === 1) {
      const data = r.bytes.subarray(r.o, end);
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
        if (bpp === 32) { c = r.dv.getUint32(row + x * 4, true); if (format === 0x00022009) c |= 0xff000000; }
        else if (bpp === 24) c = 0xff000000 | (r.bytes[row + x * 3 + 2] << 16) | (r.bytes[row + x * 3 + 1] << 8) | r.bytes[row + x * 3];
        else if (bpp === 16) {
          const v = r.dv.getUint16(row + x * 2, true), f5 = (q) => Math.round((q * 255) / 31);
          if (format === 0x00021006) c = 0xff000000 | (f5(v >> 11) << 16) | (Math.round((((v >> 5) & 63) * 255) / 63) << 8) | f5(v & 31); // RGB565
          else if (format === 0x00021005 || format === 0x00061007) c = ((format === 0x00061007 && !(v & 0x8000) ? 0 : 255) << 24 | (f5((v >> 10) & 31) << 16) | (f5((v >> 5) & 31) << 8) | f5(v & 31)) >>> 0; // RGB555, ARGB1555
          else if (format === 0x00101004) { const g = v >> 8; c = (0xff000000 | (g << 16) | (g << 8) | g) >>> 0; } // 16-bit grey
          else return null;
        }
        else if (palette && (bpp === 8 || bpp === 4 || bpp === 1)) {
          const bit = x * bpp, byte = r.bytes[row + (bit >> 3)];
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
    const inner = r.bytes.subarray(r.o, Math.min(end, r.o + size));
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
export const readBrush = (r, end, h) => {
  r.u32(); // version
  const type = r.u32();
  if (type === 0) return { kind: 'brush', type: 'solid', color: r.u32() };
  if (type === 1) return { kind: 'brush', type: 'hatch', style: r.u32(), fore: r.u32(), back: r.u32() };
  if (type === 2) { // texture
    const flags = r.u32(), wrap = r.u32();
    const m = flags & 2 ? r.m() : null;
    const img = readImage(r, end, h);
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
export const lerp = (a, b, t) => {
  const ch = (s) => Math.round(((a >>> s) & 255) * (1 - t) + ((b >>> s) & 255) * t);
  return ((ch(24) << 24) | (ch(16) << 16) | (ch(8) << 8) | ch(0)) >>> 0;
};
export const readPen = (r, end, h) => {
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
  pen.brush = r.o < end ? readBrush(r, end, h) : { kind: 'brush', type: 'solid', color: 0xff000000 };
  return pen;
};
export const readFont = (r) => {
  r.u32();
  const em = r.f32(), unit = r.u32(), style = r.i32(); r.u32();
  const len = r.u32();
  let face = utf16(r, len).replace(/\0.*$/, '');
  // (GDI+ may write the family in capitals, "ARIAL": font names match whatever their case, the stand-ins don't.)
  if (face === face.toUpperCase()) face = face.toLowerCase().replace(/\b[a-z]/g, (c) => c.toUpperCase());
  return { kind: 'font', em, unit, bold: !!(style & 1), italic: !!(style & 2), underline: !!(style & 4), strike: !!(style & 8), face };
};
export const readFormat = (r) => {
  r.u32();
  const flags = r.u32(); r.u32();
  const align = r.u32(), lineAlign = r.u32();
  r.u32(); r.u32(); r.f32(); r.i32(); // digit substitution and language, first tab offset, hotkey prefix
  return { kind: 'format', flags, align, lineAlign, lead: r.f32(), trail: r.f32() };
};
export const readRegion = (r) => {
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
export const readObject = (type, r, end, h) => {
  switch (type) {
    case 1: return readBrush(r, end, h);
    case 2: return readPen(r, end, h);
    case 3: return readPath(r);
    case 4: return readRegion(r);
    case 5: return readImage(r, end, h);
    case 6: return readFont(r);
    case 7: return readFormat(r);
    default: return null;
  }
};

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
