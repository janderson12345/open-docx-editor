// Fonts embedded in the document (word/fonts/*.odttf, listed in fontTable.xml): Word lays the text out in them, so
// the browser has to have them too. An embedded font is obfuscated: its first 32 bytes are XORed with its key, the
// GUID's 16 bytes in reverse order (ECMA-376 part 2, 17.8.1). The package keeps them as they are.

const FACES = { embedRegular: ['normal', 'normal'], embedBold: ['bold', 'normal'], embedItalic: ['normal', 'italic'], embedBoldItalic: ['bold', 'italic'] };

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
    if (name && v && v !== name) out.set(name, v.split(/[;,]/)[0].trim());
  }
  return out;
}
