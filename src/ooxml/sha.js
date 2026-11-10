// SHA-1 and SHA-512 (FIPS 180-4), synchronous: Word hashes a protection password 100,000 times, which in WebCrypto's
// one-promise-a-hash takes seconds; done here it takes a fraction of one.

const K512 = [
  0x428a2f98, 0xd728ae22, 0x71374491, 0x23ef65cd, 0xb5c0fbcf, 0xec4d3b2f, 0xe9b5dba5, 0x8189dbbc, 0x3956c25b, 0xf348b538,
  0x59f111f1, 0xb605d019, 0x923f82a4, 0xaf194f9b, 0xab1c5ed5, 0xda6d8118, 0xd807aa98, 0xa3030242, 0x12835b01, 0x45706fbe,
  0x243185be, 0x4ee4b28c, 0x550c7dc3, 0xd5ffb4e2, 0x72be5d74, 0xf27b896f, 0x80deb1fe, 0x3b1696b1, 0x9bdc06a7, 0x25c71235,
  0xc19bf174, 0xcf692694, 0xe49b69c1, 0x9ef14ad2, 0xefbe4786, 0x384f25e3, 0x0fc19dc6, 0x8b8cd5b5, 0x240ca1cc, 0x77ac9c65,
  0x2de92c6f, 0x592b0275, 0x4a7484aa, 0x6ea6e483, 0x5cb0a9dc, 0xbd41fbd4, 0x76f988da, 0x831153b5, 0x983e5152, 0xee66dfab,
  0xa831c66d, 0x2db43210, 0xb00327c8, 0x98fb213f, 0xbf597fc7, 0xbeef0ee4, 0xc6e00bf3, 0x3da88fc2, 0xd5a79147, 0x930aa725,
  0x06ca6351, 0xe003826f, 0x14292967, 0x0a0e6e70, 0x27b70a85, 0x46d22ffc, 0x2e1b2138, 0x5c26c926, 0x4d2c6dfc, 0x5ac42aed,
  0x53380d13, 0x9d95b3df, 0x650a7354, 0x8baf63de, 0x766a0abb, 0x3c77b2a8, 0x81c2c92e, 0x47edaee6, 0x92722c85, 0x1482353b,
  0xa2bfe8a1, 0x4cf10364, 0xa81a664b, 0xbc423001, 0xc24b8b70, 0xd0f89791, 0xc76c51a3, 0x0654be30, 0xd192e819, 0xd6ef5218,
  0xd6990624, 0x5565a910, 0xf40e3585, 0x5771202a, 0x106aa070, 0x32bbd1b8, 0x19a4c116, 0xb8d2d0c8, 0x1e376c08, 0x5141ab53,
  0x2748774c, 0xdf8eeb99, 0x34b0bcb5, 0xe19b48a8, 0x391c0cb3, 0xc5c95a63, 0x4ed8aa4a, 0xe3418acb, 0x5b9cca4f, 0x7763e373,
  0x682e6ff3, 0xd6b2b8a3, 0x748f82ee, 0x5defb2fc, 0x78a5636f, 0x43172f60, 0x84c87814, 0xa1f0ab72, 0x8cc70208, 0x1a6439ec,
  0x90befffa, 0x23631e28, 0xa4506ceb, 0xde82bde9, 0xbef9a3f7, 0xb2c67915, 0xc67178f2, 0xe372532b, 0xca273ece, 0xea26619c,
  0xd186b8c7, 0x21c0c207, 0xeada7dd6, 0xcde0eb1e, 0xf57d4f7f, 0xee6ed178, 0x06f067aa, 0x72176fba, 0x0a637dc5, 0xa2c898a6,
  0x113f9804, 0xbef90dae, 0x1b710b35, 0x131c471b, 0x28db77f5, 0x23047d84, 0x32caab7b, 0x40c72493, 0x3c9ebe0a, 0x15c9bebc,
  0x431d67c4, 0x9c100d4c, 0x4cc5d4be, 0xcb3e42b6, 0x597f299c, 0xfc657e2a, 0x5fcb6fab, 0x3ad6faec, 0x6c44198c, 0x4a475817,
];
const H512 = [0x6a09e667, 0xf3bcc908, 0xbb67ae85, 0x84caa73b, 0x3c6ef372, 0xfe94f82b, 0xa54ff53a, 0x5f1d36f1, 0x510e527f, 0xade682d1, 0x9b05688c, 0x2b3e6c1f, 0x1f83d9ab, 0xfb41bd6b, 0x5be0cd19, 0x137e2179];

/** The message padded to whole blocks: 0x80, zeros, its length in bits (big-endian) in the last `lenBytes`. */
function pad(bytes, block, lenBytes) {
  const n = Math.ceil((bytes.length + 1 + lenBytes) / block) * block;
  const out = new Uint8Array(n);
  out.set(bytes);
  out[bytes.length] = 0x80;
  const bits = bytes.length * 8;
  const v = new DataView(out.buffer);
  v.setUint32(n - 4, bits >>> 0);
  v.setUint32(n - 8, Math.floor(bits / 0x100000000));
  return v;
}

const W = new Int32Array(160);

export function sha512(bytes) {
  const v = pad(bytes, 128, 16);
  const h = Int32Array.from(H512);
  for (let off = 0; off < v.byteLength; off += 128) {
    for (let i = 0; i < 32; i++) W[i] = v.getInt32(off + i * 4);
    for (let i = 32; i < 160; i += 2) {
      // s0 = rotr1 ^ rotr8 ^ shr7 of W[i-30]; s1 = rotr19 ^ rotr61 ^ shr6 of W[i-4]
      let xh = W[i - 30], xl = W[i - 29];
      const s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
      const s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25));
      xh = W[i - 4]; xl = W[i - 3];
      const s1h = ((xh >>> 19) | (xl << 13)) ^ ((xl >>> 29) | (xh << 3)) ^ (xh >>> 6);
      const s1l = ((xl >>> 19) | (xh << 13)) ^ ((xh >>> 29) | (xl << 3)) ^ ((xl >>> 6) | (xh << 26));
      let lo = (s0l >>> 0) + (s1l >>> 0) + (W[i - 13] >>> 0) + (W[i - 31] >>> 0);
      W[i] = s0h + s1h + W[i - 14] + W[i - 32] + Math.floor(lo / 0x100000000);
      W[i + 1] = lo;
    }
    let [ah, al, bh, bl, ch, cl, dh, dl, eh, el, fh, fl, gh, gl, hh, hl] = h;
    for (let i = 0; i < 160; i += 2) {
      // S1 = rotr14 ^ rotr18 ^ rotr41 of e; ch = (e & f) ^ (~e & g)
      const S1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
      const S1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
      const chh = (eh & fh) ^ (~eh & gh), chl = (el & fl) ^ (~el & gl);
      let lo = (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + (K512[i + 1] >>> 0) + (W[i + 1] >>> 0);
      const t1h = (hh + S1h + chh + K512[i] + W[i] + Math.floor(lo / 0x100000000)) | 0, t1l = lo | 0;
      // S0 = rotr28 ^ rotr34 ^ rotr39 of a; maj = (a & b) ^ (a & c) ^ (b & c)
      const S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
      const S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
      const mh = (ah & bh) ^ (ah & ch) ^ (bh & ch), ml = (al & bl) ^ (al & cl) ^ (bl & cl);
      lo = (S0l >>> 0) + (ml >>> 0);
      const t2h = (S0h + mh + Math.floor(lo / 0x100000000)) | 0, t2l = lo | 0;
      hh = gh; hl = gl; gh = fh; gl = fl; fh = eh; fl = el;
      lo = (dl >>> 0) + (t1l >>> 0);
      eh = (dh + t1h + Math.floor(lo / 0x100000000)) | 0; el = lo | 0;
      dh = ch; dl = cl; ch = bh; cl = bl; bh = ah; bl = al;
      lo = (t1l >>> 0) + (t2l >>> 0);
      ah = (t1h + t2h + Math.floor(lo / 0x100000000)) | 0; al = lo | 0;
    }
    const add = [ah, al, bh, bl, ch, cl, dh, dl, eh, el, fh, fl, gh, gl, hh, hl];
    for (let i = 0; i < 16; i += 2) {
      const lo = (h[i + 1] >>> 0) + (add[i + 1] >>> 0);
      h[i] = h[i] + add[i] + Math.floor(lo / 0x100000000);
      h[i + 1] = lo;
    }
  }
  const out = new Uint8Array(64);
  const o = new DataView(out.buffer);
  for (let i = 0; i < 16; i++) o.setInt32(i * 4, h[i]);
  return out;
}

export function sha1(bytes) {
  const v = pad(bytes, 64, 8);
  const h = Int32Array.of(0x67452301, 0xefcdab89, 0x98badcfe, 0x10325476, 0xc3d2e1f0);
  for (let off = 0; off < v.byteLength; off += 64) {
    for (let i = 0; i < 16; i++) W[i] = v.getInt32(off + i * 4);
    for (let i = 16; i < 80; i++) { const x = W[i - 3] ^ W[i - 8] ^ W[i - 14] ^ W[i - 16]; W[i] = (x << 1) | (x >>> 31); }
    let [a, b, c, d, e] = h;
    for (let i = 0; i < 80; i++) {
      const f = i < 20 ? (b & c) | (~b & d) : i < 40 ? b ^ c ^ d : i < 60 ? (b & c) | (b & d) | (c & d) : b ^ c ^ d;
      const k = i < 20 ? 0x5a827999 : i < 40 ? 0x6ed9eba1 : i < 60 ? 0x8f1bbcdc : 0xca62c1d6;
      const t = (((a << 5) | (a >>> 27)) + f + e + k + W[i]) | 0;
      e = d; d = c; c = (b << 30) | (b >>> 2); b = a; a = t;
    }
    h[0] += a; h[1] += b; h[2] += c; h[3] += d; h[4] += e;
  }
  const out = new Uint8Array(20);
  const o = new DataView(out.buffer);
  for (let i = 0; i < 5; i++) o.setInt32(i * 4, h[i]);
  return out;
}
