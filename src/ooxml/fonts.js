// Fonts embedded in the document (word/fonts/*.odttf, listed in fontTable.xml): Word lays the text out in them, so
// the browser has to have them too. An embedded font is obfuscated: its first 32 bytes are XORed with its key, the
// GUID's 16 bytes in reverse order (ECMA-376 part 2, 17.8.1). The package keeps them as they are.

const FACES = { embedRegular: ['normal', 'normal'], embedBold: ['bold', 'normal'], embedItalic: ['normal', 'italic'], embedBoldItalic: ['bold', 'italic'] };

const familyMemo = new WeakMap();
/** The names (lowercase) of the fonts the document embeds a face of, from its font table. */
export function embeddedFamilies(pkg) {
  if (familyMemo.has(pkg)) return familyMemo.get(pkg);
  const out = new Set();
  const rel = pkg.rels?.find((r) => /\/fontTable$/.test(r.type));
  const doc = rel && pkg.xml(rel.target);
  for (const font of doc ? doc.getElementsByTagNameNS('*', 'font') : []) {
    const family = font.getAttributeNS(font.namespaceURI, 'name') || font.getAttribute('w:name');
    if (family && [...font.children].some((el) => FACES[el.localName])) out.add(family.toLowerCase());
  }
  familyMemo.set(pkg, out);
  return out;
}

/** [{ family, weight, style, bytes }] for the document's embedded fonts (de-obfuscated copies). */
export function embeddedFonts(pkg) {
  const rel = pkg.rels.find((r) => /\/fontTable$/.test(r.type));
  const doc = rel && pkg.xml(rel.target);
  if (!doc) return [];
  const rels = pkg.readRels(rel.target);
  const out = [];
  for (const font of doc.getElementsByTagNameNS('*', 'font')) {
    const family = font.getAttributeNS(font.namespaceURI, 'name') || font.getAttribute('w:name');
    if (!family) continue;
    for (const el of font.children) {
      const face = FACES[el.localName];
      if (!face) continue;
      const id = el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || el.getAttribute('r:id');
      const key = (el.getAttributeNS(el.namespaceURI, 'fontKey') || el.getAttribute('w:fontKey') || '').replace(/[{}-]/g, '');
      const target = rels.find((r) => r.id === id)?.target;
      const data = target && pkg.files[target];
      if (!data || data.length < 32) continue;
      const bytes = data.slice();
      if (key.length === 32) {
        const k = [];
        for (let i = 0; i < 16; i++) k.push(parseInt(key.substr(30 - i * 2, 2), 16));
        for (let i = 0; i < 32; i++) bytes[i] ^= k[i % 16];
      }
      out.push({ family, weight: face[0], style: face[1], bytes });
    }
  }
  return out;
}

let added = [];

/** Makes the document's embedded fonts available to the page (replacing the last document's); resolves when loaded. */
export async function loadEmbeddedFonts(pkg) {
  if (typeof FontFace === 'undefined' || !document.fonts) return;
  for (const f of added) document.fonts.delete(f);
  added = [];
  for (const { family, weight, style, bytes } of embeddedFonts(pkg)) {
    try {
      const face = new FontFace(family, bytes, { weight, style });
      document.fonts.add(face);
      added.push(face);
    } catch { /* a font the browser can't read: left to the fallbacks */ }
  }
  await Promise.all(added.map((f) => f.load().catch(() => null)));
}

// With no other name given, Word sets a missing font whose entry claims Armenian or Georgian letters (w:sig usb0 bits
// 10 and 26) in Sylfaen, whatever its family, panose or name: the Mac's "Helvetica Neue" (usb0 E50002FF) prints in
// Sylfaen, a serif, where the same name with only Latin bits prints in Arial. Word's PDFs of 40 generated documents,
// one missing font each, and of six samples. (Windows' own substitutes, Times and Helvetica, still win over it.)
function scriptAlt(font) {
  const sig = [...font.children].find((e) => e.localName === 'sig');
  const usb0 = parseInt(sig && (sig.getAttributeNS(sig.namespaceURI, 'usb0') || sig.getAttribute('w:usb0')), 16);
  return usb0 & ((1 << 10) | (1 << 26)) ? 'Sylfaen' : null;
}

/** The font table's stand-ins (w:altName): Word sets a font it doesn't have in the one named there (sample S063's
 * "CIDFont+F1", from a converted PDF, prints in Calibri). Map of font name -> alternative. */
export function fontAlts(pkg) {
  const rel = pkg.rels.find((r) => /\/fontTable$/.test(r.type));
  const doc = rel && pkg.xml(rel.target);
  const out = new Map();
  if (!doc) return out;
  for (const font of doc.getElementsByTagNameNS('*', 'font')) {
    const name = font.getAttributeNS(font.namespaceURI, 'name') || font.getAttribute('w:name');
    const alt = [...font.children].find((e) => e.localName === 'altName');
    const v = alt && (alt.getAttributeNS(alt.namespaceURI, 'val') || alt.getAttribute('w:val'));
    // (An altName that is no name, a run of U+FDFD before "iqua" where a converter lost "Book Antiqua", Word passes
    // over: a new test document naming a missing font so prints in its fallback, Calibri, and 14 of 15 samples' Palatino
    // so print in Book Antiqua, as Palatino with no altName does.)
    if (name && v && v !== name && !NO_NAME.test(v)) out.set(name, v.split(/[;,]/)[0].trim());
    else if (name && !/^sylfaen$/i.test(name) && scriptAlt(font)) out.set(name, scriptAlt(font)); // (never Sylfaen as itself)
  }
  return out;
}
// (Charsets Word falls back from by the family: ANSI, default, Mac, Central European, Cyrillic, Greek, Turkish, Baltic.)
const LATIN_CHARSETS = new Set(['00', '01', '4D', 'EE', 'CC', 'A1', 'A2', 'BA']);
const NO_NAME =/[\u0000-\u001f\ufdd0-\ufdef\ufdfd\ufffd-\uffff]/;

/**
 * The font Word falls back to for a font it has no other word for (no altName, no name it knows, not installed), by
 * its font table entry: Cambria for a roman family, Calibri for any other, whatever the document's default font. A
 * font with no entry is Cambria (fontStack). Word's PDFs of 30 new test documents, a missing font each with a Latin
 * w:sig, the family auto, swiss, roman, script, modern or decorative, fixed pitch or variable, under a Calibri, Times
 * New Roman or Arial default; and of samples: Montserrat Medium, ADLaM Display and a script font with no altName in
 * Calibri. The same in the Mac, Central European, Cyrillic and Greek charsets; in the Hebrew one Arial, or Times New
 * Roman for a roman family (a sample's Apple Chancery prints in Arial); in others (symbol, East Asian) the browser's
 * generic family stands (null). Map of font name -> the fallback, or null.
 */
export function fontFallbacks(pkg) {
  const rel = pkg.rels.find((r) => /\/fontTable$/.test(r.type));
  const doc = rel && pkg.xml(rel.target);
  const out = new Map();
  if (!doc) return out;
  const val = (font, tag) => {
    const e = [...font.children].find((c) => c.localName === tag);
    return e ? e.getAttributeNS(e.namespaceURI, 'val') || e.getAttribute('w:val') || '' : null;
  };
  for (const font of doc.getElementsByTagNameNS('*', 'font')) {
    const name = font.getAttributeNS(font.namespaceURI, 'name') || font.getAttribute('w:name');
    if (!name) continue;
    const charset = (val(font, 'charset') || '00').toUpperCase();
    const roman = /^roman$/i.test(val(font, 'family') || '');
    if (charset === 'B1') out.set(name, roman ? 'Times New Roman' : 'Arial'); // (Hebrew)
    else out.set(name, LATIN_CHARSETS.has(charset) ? (roman ? 'Cambria' : 'Calibri') : null);
  }
  return out;
}
