// EMF (Windows enhanced metafile) pictures as SVG. Word documents carry them for pasted drawings, Visio and Office
// screenshots; a browser can't show them. This draws the GDI records (an EMF+ file's own records ride in comments and
// are skipped: Word writes a GDI copy alongside): shapes and lines with their pens and brushes, text with its font and
// character advances, bitmaps, clip rectangles, and the window/viewport and world transforms.

import { fontStack } from './props.js';
import { esc } from './xml.js';

const rgb = (v) => `#${[v & 255, (v >> 8) & 255, (v >> 16) & 255].map((n) => n.toString(16).padStart(2, '0')).join('')}`;
const n = (v) => +v.toFixed(3);
// Windows ascent and descent (em) of common fonts: GDI puts top-aligned text's baseline an ascent below the point,
// and a positive LOGFONT height is the whole cell (ascent + descent).
const WIN = { arial: [0.905, 0.212], calibri: [0.952, 0.269], cambria: [0.95, 0.222], 'segoe ui': [1.079, 0.251], tahoma: [1.0, 0.207],
  verdana: [1.005, 0.21], 'times new roman': [0.891, 0.216], 'courier new': [0.833, 0.3], aptos: [1.01, 0.275], 'aptos narrow': [1.01, 0.275],
  'aptos display': [1.01, 0.275], 'arial narrow': [0.936, 0.212], consolas: [0.743, 0.257], georgia: [0.917, 0.219] };
const win = (face) => WIN[face.toLowerCase()] || [0.92, 0.23];

/** True when `bytes` is an EMF (record 1, with the " EMF" signature). */
export function isEmf(bytes) {
  return bytes && bytes.length > 88 && bytes[0] === 1 && bytes[1] === 0 && bytes[2] === 0 && bytes[3] === 0 &&
    bytes[40] === 0x20 && bytes[41] === 0x45 && bytes[42] === 0x4d && bytes[43] === 0x46;
}

const dataUrl = (mime, out) => {
  let s = '';
  for (let i = 0; i < out.length; i += 0x8000) s += String.fromCharCode(...out.subarray(i, i + 0x8000));
  return `data:${mime};base64,${btoa(s)}`;
};

let crcTable = null;
const crc32 = (buf) => {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let k = 0; k < 256; k++) { let c = k; for (let j = 0; j < 8; j++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; crcTable[k] = c >>> 0; }
  }
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = crcTable[(c ^ buf[i]) & 255] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};

/** A PNG of RGBA rows (deflate in stored blocks: no compression, nothing to depend on). */
function png(w, h, rgba) {
  const raw = new Uint8Array(h * (w * 4 + 1));
  for (let y = 0; y < h; y++) raw.set(rgba.subarray(y * w * 4, (y + 1) * w * 4), y * (w * 4 + 1) + 1);
  const blocks = Math.ceil(raw.length / 65535) || 1;
  const z = new Uint8Array(2 + raw.length + blocks * 5 + 4);
  z[0] = 0x78; z[1] = 0x01;
  let p = 2;
  for (let b = 0; b < blocks; b++) {
    const part = raw.subarray(b * 65535, Math.min(raw.length, (b + 1) * 65535));
    z[p++] = b === blocks - 1 ? 1 : 0;
    z[p++] = part.length & 255; z[p++] = part.length >> 8; z[p++] = ~part.length & 255; z[p++] = (~part.length >> 8) & 255;
    z.set(part, p); p += part.length;
  }
  let a = 1, bsum = 0;
  for (let i = 0; i < raw.length; i++) { a = (a + raw[i]) % 65521; bsum = (bsum + a) % 65521; }
  new DataView(z.buffer).setUint32(p, ((bsum << 16) | a) >>> 0);
  const chunk = (type, data) => {
    const c = new Uint8Array(12 + data.length);
    const dv = new DataView(c.buffer);
    dv.setUint32(0, data.length);
    for (let i = 0; i < 4; i++) c[4 + i] = type.charCodeAt(i);
    c.set(data, 8);
    dv.setUint32(8 + data.length, crc32(c.subarray(4, 8 + data.length)));
    return c;
  };
  const ihdr = new Uint8Array(13);
  const hv = new DataView(ihdr.buffer);
  hv.setUint32(0, w); hv.setUint32(4, h); ihdr[8] = 8; ihdr[9] = 6;
  const parts = [new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', z), chunk('IEND', new Uint8Array(0))];
  const out = new Uint8Array(parts.reduce((s, q) => s + q.length, 0));
  let at = 0;
  for (const q of parts) { out.set(q, at); at += q.length; }
  return dataUrl('image/png', out);
}

/** A BMP file from a DIB header and bits (a browser shows BMP data URLs); 32-bit DIBs as PNG, keeping their alpha. */
function bmp(bytes, offBmi, cbBmi, offBits, cbBits) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const bits = dv.getUint16(offBmi + 14, true), compression = dv.getUint32(offBmi + 16, true);
  if (bits === 32 && (compression === 0 || compression === 3)) {
    const w = dv.getInt32(offBmi + 4, true), hh = dv.getInt32(offBmi + 8, true), h = Math.abs(hh);
    if (w > 0 && h > 0 && cbBits >= w * h * 4) {
      const rgba = new Uint8Array(w * h * 4);
      let anyAlpha = false;
      for (let i = 3; i < w * h * 4; i += 4) if (bytes[offBits + i]) { anyAlpha = true; break; }
      for (let y = 0; y < h; y++) {
        const src = offBits + (hh > 0 ? h - 1 - y : y) * w * 4;
        for (let x = 0; x < w; x++) {
          const s = src + x * 4, d = (y * w + x) * 4;
          rgba[d] = bytes[s + 2]; rgba[d + 1] = bytes[s + 1]; rgba[d + 2] = bytes[s]; rgba[d + 3] = anyAlpha ? bytes[s + 3] : 255;
        }
      }
      return png(w, h, rgba);
    }
  }
  const out = new Uint8Array(14 + cbBmi + cbBits);
  const ov = new DataView(out.buffer);
  out[0] = 0x42; out[1] = 0x4d;
  ov.setUint32(2, out.length, true);
  ov.setUint32(10, 14 + cbBmi, true);
  out.set(bytes.subarray(offBmi, offBmi + cbBmi), 14);
  out.set(bytes.subarray(offBits, offBits + cbBits), 14 + cbBmi);
  return dataUrl('image/bmp', out);
}

/** The EMF in `bytes` as an SVG document (string), or null when it can't be read. */
export function emfToSvg(bytes) {
  if (!isEmf(bytes)) return null;
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const i32 = (o) => dv.getInt32(o, true);
  const u32 = (o) => dv.getUint32(o, true);
  const i16 = (o) => dv.getInt16(o, true);
  const u16 = (o) => dv.getUint16(o, true);
  const f32 = (o) => dv.getFloat32(o, true);
  // The picture's frame (rclFrame, 0.01mm) in device units is the SVG's view box: Word fits the frame, not the drawing's
  // tight bounds, to the picture's size. (Device pixels per mm from the reference device's size.)
  const pxmX = i32(72) / (i32(80) || 1), pxmY = i32(76) / (i32(84) || 1);
  let bx = (i32(24) / 100) * pxmX, by = (i32(28) / 100) * pxmY;
  let bw = ((i32(32) - i32(24)) / 100) * pxmX, bh = ((i32(36) - i32(28)) / 100) * pxmY;
  if (!(bw > 0 && bh > 0 && Number.isFinite(bw + bh))) { bx = i32(8); by = i32(12); bw = i32(16) - bx + 1; bh = i32(20) - by + 1; }
  if (bw <= 0 || bh <= 0) return null;

  const objects = [];
  const stock = (k) => {
    if (k === 0x80000000) return { kind: 'brush', style: 0, color: 0xffffff }; // WHITE_BRUSH
    if (k === 0x80000004) return { kind: 'brush', style: 0, color: 0 }; // BLACK_BRUSH
    if (k === 0x80000005) return { kind: 'brush', style: 1 }; // NULL_BRUSH
    if (k === 0x80000001 || k === 0x80000002 || k === 0x80000003) return { kind: 'brush', style: 0, color: [0xc0c0c0, 0x808080, 0x404040][k - 0x80000001] };
    if (k === 0x80000006) return { kind: 'pen', style: 0, width: 0, color: 0xffffff }; // WHITE_PEN
    if (k === 0x80000007) return { kind: 'pen', style: 0, width: 0, color: 0 }; // BLACK_PEN
    if (k === 0x80000008) return { kind: 'pen', style: 5 }; // NULL_PEN
    if (k >= 0x8000000a && k <= 0x80000011) return { kind: 'font', height: -12, weight: 400, italic: false, face: 'Segoe UI' };
    return null;
  };
  let dc = {
    win: [0, 0], winExt: [1, 1], vp: [0, 0], vpExt: [1, 1], mapMode: 1,
    world: [1, 0, 0, 1, 0, 0],
    pen: { kind: 'pen', style: 0, width: 0, color: 0 },
    brush: { kind: 'brush', style: 0, color: 0xffffff },
    font: { kind: 'font', height: -12, weight: 400, italic: false, face: 'Segoe UI' },
    textColor: 0, bkColor: 0xffffff, bkMode: 2, textAlign: 0, fill: 'evenodd',
    pos: [0, 0], clip: null,
  };
  const stack = [];
  const out = [];
  let path = null; // the figure being built between BEGINPATH and ENDPATH: SVG path data
  let clipId = 0;
  const clips = [];

  // Logical point -> device (the view box's units): world transform, then window to viewport.
  const scale = () => (dc.mapMode === 1 ? [1, 1] : [dc.vpExt[0] / (dc.winExt[0] || 1), dc.vpExt[1] / (dc.winExt[1] || 1)]);
  const pt = (x, y) => {
    const [a, b, c, d, e, f] = dc.world;
    const px = a * x + c * y + e, py = b * x + d * y + f;
    const [sx, sy] = scale();
    return [(px - dc.win[0]) * sx + dc.vp[0] - bx, (py - dc.win[1]) * sy + dc.vp[1] - by];
  };
  const len = (v) => { const [sx] = scale(); return Math.abs(v * sx * Math.hypot(dc.world[0], dc.world[1])); };
  const clipAttr = () => (dc.clip ? ` clip-path="url(#c${dc.clip})"` : '');
  const stroke = () => {
    const p = dc.pen;
    if (!p || p.style === 5) return 'stroke="none"';
    const w = Math.max(len(p.width || 0), 0.75);
    const dash = { 1: '4 2', 2: '1 2', 3: '4 2 1 2', 4: '4 2 1 2 1 2' }[p.style];
    return `stroke="${rgb(p.color)}" stroke-width="${n(w)}"${dash ? ` stroke-dasharray="${dash.split(' ').map((v) => n(v * w)).join(' ')}"` : ''} stroke-linejoin="miter"`;
  };
  const fill = () => {
    const b = dc.brush;
    if (!b || b.style === 1) return 'fill="none"';
    if (b.style === 3 && b.image) return `fill="${rgb(0xc0c0c0)}"`;
    return `fill="${rgb(b.color ?? 0)}"`;
  };
  const poly = (o, count, small, close, drawFill) => {
    const pts = [];
    for (let k = 0; k < count; k++) pts.push(small ? pt(i16(o + k * 4), i16(o + k * 4 + 2)) : pt(i32(o + k * 8), i32(o + k * 8 + 4)));
    return { pts, d: `M${pts.map(([x, y]) => `${n(x)} ${n(y)}`).join(' L')}${close ? 'Z' : ''}`, drawFill };
  };
  const shape = (d, filled) => {
    if (path) { path.push(d); return; }
    out.push(`<path d="${d}" ${filled ? fill() : 'fill="none"'} ${stroke()} fill-rule="${dc.fill}"${clipAttr()}/>`);
  };

  let off = 0;
  while (off + 8 <= bytes.length) {
    const type = u32(off), size = u32(off + 4);
    if (size < 8 || off + size > bytes.length) break;
    const o = off + 8; // the record's data
    switch (type) {
      case 9: dc.winExt = [i32(o), i32(o + 4)]; break; // SETWINDOWEXTEX
      case 10: dc.win = [i32(o), i32(o + 4)]; break; // SETWINDOWORGEX
      case 11: dc.vpExt = [i32(o), i32(o + 4)]; break; // SETVIEWPORTEXTEX
      case 12: dc.vp = [i32(o), i32(o + 4)]; break; // SETVIEWPORTORGEX
      case 17: dc.mapMode = u32(o); break; // SETMAPMODE
      case 18: dc.bkMode = u32(o); break; // SETBKMODE
      case 19: dc.fill = u32(o) === 2 ? 'nonzero' : 'evenodd'; break; // SETPOLYFILLMODE
      case 22: dc.textAlign = u32(o); break; // SETTEXTALIGN
      case 24: dc.textColor = u32(o); break; // SETTEXTCOLOR
      case 25: dc.bkColor = u32(o); break; // SETBKCOLOR
      case 27: dc.pos = [i32(o), i32(o + 4)]; break; // MOVETOEX
      case 33: stack.push({ ...dc }); break; // SAVEDC
      case 34: { // RESTOREDC (relative: -1 is the last saved)
        const k = i32(o);
        const idx = k < 0 ? stack.length + k : k - 1;
        if (idx >= 0 && idx < stack.length) { dc = stack[idx]; stack.length = idx; }
        break;
      }
      case 35: dc.world = [f32(o), f32(o + 4), f32(o + 8), f32(o + 12), f32(o + 16), f32(o + 20)]; break; // SETWORLDTRANSFORM
      case 36: { // MODIFYWORLDTRANSFORM
        const m = [f32(o), f32(o + 4), f32(o + 8), f32(o + 12), f32(o + 16), f32(o + 20)];
        const mode = u32(o + 24);
        if (mode === 1) dc.world = [1, 0, 0, 1, 0, 0];
        else {
          const [a, b, c, d, e, f] = mode === 2 ? m : dc.world; // left multiply: m * world
          const [A, B, C, D, E, F] = mode === 2 ? dc.world : m;
          if (mode === 4) dc.world = m;
          else dc.world = [a * A + b * C, a * B + b * D, c * A + d * C, c * B + d * D, e * A + f * C + E, e * B + f * D + F];
        }
        break;
      }
      case 37: { // SELECTOBJECT
        const k = u32(o);
        const obj = k & 0x80000000 ? stock(k) : objects[k];
        if (obj) dc = { ...dc, [obj.kind]: obj };
        break;
      }
      case 38: objects[u32(o)] = { kind: 'pen', style: u32(o + 4) & 15, width: i32(o + 8), color: u32(o + 16) }; break; // CREATEPEN
      case 95: objects[u32(o)] = { kind: 'pen', style: u32(o + 20) & 15, width: u32(o + 24), color: u32(o + 32) }; break; // EXTCREATEPEN
      case 39: objects[u32(o)] = { kind: 'brush', style: u32(o + 4), color: u32(o + 8) }; break; // CREATEBRUSHINDIRECT
      case 93: case 94: objects[u32(o)] = { kind: 'brush', style: 3, image: true }; break; // pattern brushes: grey
      case 40: objects[u32(o)] = undefined; break; // DELETEOBJECT
      case 82: { // EXTCREATEFONTINDIRECTW: a LOGFONTW
        const lf = o + 4;
        let face = '';
        for (let k = 0; k < 32; k++) { const c = u16(lf + 28 + k * 2); if (!c) break; face += String.fromCharCode(c); }
        objects[u32(o)] = { kind: 'font', height: i32(lf), escapement: i32(lf + 8), weight: i32(lf + 16), italic: !!bytes[lf + 20], underline: !!bytes[lf + 21], strike: !!bytes[lf + 22], face: face || 'Segoe UI' };
        break;
      }
      case 30: case 75: { // INTERSECTCLIPRECT / EXCLUDECLIPRECT (an excluded area is left drawn)
        if (type === 75) break;
        const [x1, y1] = pt(i32(o), i32(o + 4)), [x2, y2] = pt(i32(o + 8), i32(o + 12));
        clips.push(`<clipPath id="c${++clipId}"><rect x="${n(Math.min(x1, x2))}" y="${n(Math.min(y1, y2))}" width="${n(Math.abs(x2 - x1))}" height="${n(Math.abs(y2 - y1))}"/></clipPath>`);
        dc.clip = clipId;
        break;
      }
      case 28: case 67: dc.clip = null; break; // SETMETARGN, SELECTCLIPPATH: back to the whole picture
      case 59: path = []; break; // BEGINPATH
      case 60: break; // ENDPATH
      case 61: if (path) path.push('Z'); break; // CLOSEFIGURE
      case 62: case 63: case 64: { // FILLPATH, STROKEANDFILLPATH, STROKEPATH
        if (!path) break;
        const d = path.join(' ');
        path = null;
        out.push(`<path d="${d}" ${type === 64 ? 'fill="none"' : fill()} ${type === 62 ? 'stroke="none"' : stroke()} fill-rule="${dc.fill}"${clipAttr()}/>`);
        break;
      }
      case 68: path = null; break; // ABORTPATH
      case 54: { // LINETO
        const [x1, y1] = pt(...dc.pos), [x2, y2] = pt(i32(o), i32(o + 4));
        dc.pos = [i32(o), i32(o + 4)];
        shape(`M${n(x1)} ${n(y1)} L${n(x2)} ${n(y2)}`, false);
        break;
      }
      case 43: case 42: case 44: { // RECTANGLE, ELLIPSE, ROUNDRECT
        const [x1, y1] = pt(i32(o), i32(o + 4)), [x2, y2] = pt(i32(o + 8), i32(o + 12));
        const x = Math.min(x1, x2), y = Math.min(y1, y2), w = Math.abs(x2 - x1), h = Math.abs(y2 - y1);
        if (type === 42) shape(`M${n(x)} ${n(y + h / 2)} a${n(w / 2)} ${n(h / 2)} 0 1 0 ${n(w)} 0 a${n(w / 2)} ${n(h / 2)} 0 1 0 ${n(-w)} 0Z`, true);
        else if (type === 44) {
          const rx = Math.min(len(i32(o + 16)) / 2, w / 2), ry = Math.min(len(i32(o + 20)) / 2, h / 2);
          shape(`M${n(x + rx)} ${n(y)} H${n(x + w - rx)} A${n(rx)} ${n(ry)} 0 0 1 ${n(x + w)} ${n(y + ry)} V${n(y + h - ry)} A${n(rx)} ${n(ry)} 0 0 1 ${n(x + w - rx)} ${n(y + h)} H${n(x + rx)} A${n(rx)} ${n(ry)} 0 0 1 ${n(x)} ${n(y + h - ry)} V${n(y + ry)} A${n(rx)} ${n(ry)} 0 0 1 ${n(x + rx)} ${n(y)}Z`, true);
        } else shape(`M${n(x)} ${n(y)} h${n(w)} v${n(h)} h${n(-w)}Z`, true);
        break;
      }
      case 3: case 4: case 86: case 87: { // POLYGON, POLYLINE (and 16-bit)
        const small = type > 80;
        const p = poly(o + 20, u32(o + 16), small, type === 3 || type === 86);
        shape(p.d, type === 3 || type === 86);
        break;
      }
      case 6: case 89: { // POLYLINETO
        const small = type === 89;
        const p = poly(o + 20, u32(o + 16), small, false);
        const [x0, y0] = pt(...dc.pos);
        shape(`M${n(x0)} ${n(y0)} L${p.d.slice(1)}`, false);
        const last = small ? [i16(o + 20 + (u32(o + 16) - 1) * 4), i16(o + 22 + (u32(o + 16) - 1) * 4)] : [i32(o + 20 + (u32(o + 16) - 1) * 8), i32(o + 24 + (u32(o + 16) - 1) * 8)];
        dc.pos = last;
        break;
      }
      case 2: case 5: case 85: case 88: { // POLYBEZIER, POLYBEZIERTO (and 16-bit)
        const small = type > 80, to = type === 5 || type === 88;
        const cnt = u32(o + 16);
        const p = [];
        for (let k = 0; k < cnt; k++) p.push(small ? pt(i16(o + 20 + k * 4), i16(o + 22 + k * 4)) : pt(i32(o + 20 + k * 8), i32(o + 24 + k * 8)));
        const start = to ? pt(...dc.pos) : p.shift();
        let d = `M${n(start[0])} ${n(start[1])}`;
        for (let k = 0; k + 2 < p.length; k += 3) d += ` C${n(p[k][0])} ${n(p[k][1])} ${n(p[k + 1][0])} ${n(p[k + 1][1])} ${n(p[k + 2][0])} ${n(p[k + 2][1])}`;
        if (to && cnt) dc.pos = small ? [i16(o + 20 + (cnt - 1) * 4), i16(o + 22 + (cnt - 1) * 4)] : [i32(o + 20 + (cnt - 1) * 8), i32(o + 24 + (cnt - 1) * 8)];
        shape(d, false);
        break;
      }
      case 7: case 8: case 90: case 91: { // POLYPOLYLINE, POLYPOLYGON (and 16-bit)
        const small = type > 80, closed = type === 8 || type === 91;
        const polys = u32(o + 16), total = u32(o + 20);
        let at = o + 24 + polys * 4;
        const parts = [];
        for (let k = 0; k < polys; k++) {
          const c = u32(o + 24 + k * 4);
          parts.push(poly(at, c, small, closed).d);
          at += c * (small ? 4 : 8);
        }
        if (total) shape(parts.join(' '), closed);
        break;
      }
      case 76: case 77: case 81: case 114: { // BITBLT, STRETCHBLT, STRETCHDIBITS, ALPHABLEND (laid out as STRETCHBLT)
        let x, y, w, h, offBmi, cbBmi, offBits, cbBits;
        if (type === 81) {
          x = i32(o + 16); y = i32(o + 20); w = i32(o + 64); h = i32(o + 68);
          offBmi = u32(o + 40); cbBmi = u32(o + 44); offBits = u32(o + 48); cbBits = u32(o + 52);
        } else {
          x = i32(o + 16); y = i32(o + 20); w = i32(o + 24); h = i32(o + 28);
          offBmi = u32(o + 76); cbBmi = u32(o + 80); offBits = u32(o + 84); cbBits = u32(o + 88);
        }
        const [x1, y1] = pt(x, y), [x2, y2] = pt(x + w, y + h);
        const rx = Math.min(x1, x2), ry = Math.min(y1, y2), rw = Math.abs(x2 - x1), rh = Math.abs(y2 - y1);
        if (!cbBmi || !cbBits) { // no bitmap: a fill by its raster operation (others, like the no-op EMF+ files end with, draw nothing)
          const rop = type === 81 ? u32(o + 60) : u32(o + 32);
          const color = rop === 0x00f00021 && dc.brush && dc.brush.style !== 1 ? fill() : rop === 0x00000042 ? 'fill="#000000"' : rop === 0x00ff0062 ? 'fill="#ffffff"' : null;
          if (color) out.push(`<rect x="${n(rx)}" y="${n(ry)}" width="${n(rw)}" height="${n(rh)}" ${color}${clipAttr()}/>`);
          break;
        }
        const href = bmp(bytes, off + offBmi, cbBmi, off + offBits, cbBits);
        out.push(`<image x="${n(rx)}" y="${n(ry)}" width="${n(rw)}" height="${n(rh)}" preserveAspectRatio="none" href="${href}"${clipAttr()}/>`);
        break;
      }
      case 84: case 83: { // EXTTEXTOUTW (and A)
        const wide = type === 84;
        const t = o + 28; // the EMRTEXT
        const rx = i32(t), ry = i32(t + 4), count = u32(t + 8), offString = u32(t + 12), options = u32(t + 16), offDx = u32(t + 36);
        if (!count) break;
        let s = '';
        for (let k = 0; k < count; k++) s += String.fromCharCode(wide ? u16(off + offString + k * 2) : bytes[off + offString + k]);
        const f = dc.font;
        const [asc, desc] = win(f.face);
        const fh = len(Math.abs(f.height || 12)) / (f.height > 0 ? asc + desc : 1); // a positive height is the cell's
        const align = dc.textAlign;
        const anchor = (align & 6) === 6 ? 'middle' : align & 2 ? 'end' : 'start';
        const lift = (align & 24) === 24 ? 0 : align & 8 ? -desc * fh : asc * fh; // to the baseline, along the text's y
        // Each character at its own advance (the dx array), so the text runs exactly as wide as Word drew it.
        const xs = [];
        let ax = rx;
        const step = (options & 0x2000) ? 8 : 4; // ETO_PDY: dx and dy pairs
        for (let k = 0; k < count; k++) { xs.push(ax); if (offDx) ax += i32(off + offDx + k * step); }
        const pts = xs.map((v) => pt(v, ry)).map(([px, py]) => [px, py + lift]);
        const [sx0, sy0] = pts[0];
        const xAttr = offDx && anchor === 'start' ? pts.map(([px]) => n(px)).join(' ') : n(sx0);
        if ((options & 2) && dc.bkMode === 2) { /* ETO_OPAQUE: the background is filled by the record's rectangle */
          const [ox1, oy1] = pt(i32(t + 20), i32(t + 24)), [ox2, oy2] = pt(i32(t + 28), i32(t + 32));
          out.push(`<rect x="${n(Math.min(ox1, ox2))}" y="${n(Math.min(oy1, oy2))}" width="${n(Math.abs(ox2 - ox1))}" height="${n(Math.abs(oy2 - oy1))}" fill="${rgb(dc.bkColor)}"${clipAttr()}/>`);
        }
        const deco = [f.underline && 'underline', f.strike && 'line-through'].filter(Boolean).join(' ');
        const rot = f.escapement ? ` transform="rotate(${n(-f.escapement / 10)} ${n(sx0)} ${n(sy0)})"` : '';
        out.push(`<text x="${xAttr}" y="${n(sy0)}" font-family="${esc(fontStack(f.face))}" font-size="${n(fh)}" font-weight="${f.weight >= 600 ? 'bold' : 'normal'}"${f.italic ? ' font-style="italic"' : ''}${deco ? ` text-decoration="${deco}"` : ''} fill="${rgb(dc.textColor)}" text-anchor="${anchor}" xml:space="preserve"${rot}${clipAttr()}>${esc(s)}</text>`);
        break;
      }
      case 14: off = bytes.length; break; // EOF
      default: break;
    }
    off += size;
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${n(bw)} ${n(bh)}" width="${n(bw)}" height="${n(bh)}" preserveAspectRatio="none"><defs>${clips.join('')}</defs>${out.join('')}</svg>`;
}
