// TIFF pictures (scans, faxes, older document systems), which Word shows but Chromium and Firefox can't: decoded with
// UTIF (MIT; uncompressed, LZW, PackBits, Deflate, JPEG and the CCITT fax compressions scanners use) and shown as a
// PNG. The document keeps the TIFF as it was; only what's drawn is the PNG.
import UTIF from 'utif';
import { zlibSync } from 'fflate';

/** Whether the bytes are a TIFF file ("II*\0" little-endian or "MM\0*" big-endian). */
export const isTiff = (b) => b && b.length > 8 && ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2a && b[3] === 0) || (b[0] === 0x4d && b[1] === 0x4d && b[2] === 0 && b[3] === 0x2a));

// (A picture larger than this, 40 million pixels, isn't decoded: a crafted file mustn't take all the memory.)
const MAX_PIXELS = 40e6;

/** The TIFF's first image as PNG bytes, or null if it can't be read. */
export function tiffToPng(bytes) {
  try {
    const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    const ifd = UTIF.decode(buf).find((d) => d.t256 && d.t257);
    if (!ifd) return null;
    const w = ifd.t256[0], h = ifd.t257[0];
    if (!(w > 0 && h > 0 && w * h <= MAX_PIXELS)) return null;
    UTIF.decodeImage(buf, ifd);
    return png(w, h, UTIF.toRGBA8(ifd));
  } catch {
    return null;
  }
}

/** RGBA pixels as PNG bytes. */
function png(w, h, rgba) {
  const raw = new Uint8Array(h * (1 + 4 * w));
  for (let y = 0; y < h; y++) raw.set(rgba.subarray(y * 4 * w, (y + 1) * 4 * w), y * (1 + 4 * w) + 1); // (filter 0 per row)
  const head = new Uint8Array(13);
  const hv = new DataView(head.buffer);
  hv.setUint32(0, w); hv.setUint32(4, h);
  head.set([8, 6, 0, 0, 0], 8); // 8 bits, RGBA
  const parts = [new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', head), chunk('IDAT', zlibSync(raw)), chunk('IEND', new Uint8Array(0))];
  const out = new Uint8Array(parts.reduce((t, p) => t + p.length, 0));
  let at = 0;
  for (const p of parts) { out.set(p, at); at += p.length; }
  return out;
}

let crcTable = null;
function chunk(type, data) {
  if (!crcTable) crcTable = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
  const out = new Uint8Array(12 + data.length);
  const dv = new DataView(out.buffer);
  dv.setUint32(0, data.length);
  for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
  out.set(data, 8);
  let c = 0xffffffff;
  for (let i = 4; i < 8 + data.length; i++) c = crcTable[(c ^ out[i]) & 0xff] ^ (c >>> 8);
  dv.setUint32(8 + data.length, (c ^ 0xffffffff) >>> 0);
  return out;
}
