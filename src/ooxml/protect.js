// Word's Restrict Editing, as it's kept in the file: w:documentProtection in settings.xml (what may be changed:
// nothing, comments, tracked changes or form fields; optionally only some styles; with a password's hash), the
// w:permStart / w:permEnd ranges that stay editable, w:formProt on sections of a protected form, the softer
// w:writeProtection ("Always open read-only", or a password to modify) and Mark as Final (docProps/custom.xml).
//
// The password is checked as Word checks it (ECMA-376 Part 4, 14.7.1 documentProtection): Word's old 16-bit key of
// the password, as hex text, hashed with the salt, then hashed again spinCount times. Like Word's own, the protection
// only steers the editor: it isn't encryption, and a file can always be changed by other means.

import { sha512, sha1 } from './sha.js';

const EDIT = new Set(['readOnly', 'comments', 'trackedChanges', 'forms']);
const on = (v) => v === '1' || v === 'true' || v === 'on';
const attr = (tag, name) => new RegExp(`\\bw:${name}="([^"]*)"`).exec(tag)?.[1] ?? null;
// w:cryptAlgorithmSid's hash algorithms (the others, MD2 / MD4 / MD5, browsers don't have).
const SIDS = { 4: 'SHA-1', 12: 'SHA-256', 13: 'SHA-384', 14: 'SHA-512' };
const NAMES = { 'SHA-1': 'SHA-1', 'SHA-256': 'SHA-256', 'SHA-384': 'SHA-384', 'SHA-512': 'SHA-512' };

/** The protection in settings.xml: { edit, enforced, formatting, password, writeProtection: { recommended, password } }. */
export function readProtection(settingsXml) {
  const tag = /<w:documentProtection\b[^>]*>/.exec(settingsXml || '')?.[0] || '';
  const wtag = /<w:writeProtection\b[^>]*>/.exec(settingsXml || '')?.[0] || '';
  const edit = attr(tag, 'edit');
  return {
    edit: EDIT.has(edit) ? edit : null,
    enforced: on(attr(tag, 'enforcement') || ''),
    formatting: on(attr(tag, 'formatting') || ''),
    password: !!hashOf(tag),
    writeProtection: wtag ? { recommended: on(attr(wtag, 'recommended') || ''), password: !!hashOf(wtag) } : null,
    tag,
    wtag,
  };
}

/** A protection tag's hash, salt, algorithm and spin count, in either of Word's ways of writing them. */
function hashOf(tag) {
  const hash = attr(tag, 'hash') ?? attr(tag, 'hashValue');
  if (!hash) return null;
  const algorithm = attr(tag, 'algorithmName') ? NAMES[attr(tag, 'algorithmName').toUpperCase()] : SIDS[attr(tag, 'cryptAlgorithmSid') || '4'];
  return { hash, salt: attr(tag, 'salt') ?? attr(tag, 'saltValue') ?? '', spin: +(attr(tag, 'cryptSpinCount') ?? attr(tag, 'spinCount') ?? 0), algorithm };
}

const INITIAL = [0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3];
const MATRIX = [
  [0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09], [0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF],
  [0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0], [0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40],
  [0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5], [0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A],
  [0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9], [0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0],
  [0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC], [0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10],
  [0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168], [0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C],
  [0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD], [0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC],
  [0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4],
];

/** Word's old 16-bit key of a password (its first 15 characters), as 8 hex digits in Word's byte order. */
export function legacyKey(password) {
  if (!password) return '00000000';
  const chars = [...password.slice(0, 15)].map((c) => { const u = c.charCodeAt(0); return u & 0xFF || (u >> 8) & 0xFF; });
  let high = INITIAL[chars.length - 1];
  chars.forEach((c, i) => { for (let bit = 0; bit < 7; bit++) if (c & (1 << bit)) high ^= MATRIX[15 - chars.length + i][bit]; });
  let low = 0;
  for (let i = chars.length - 1; i >= 0; i--) low = (((low >> 14) & 1) | ((low << 1) & 0x7FFF)) ^ chars[i];
  low = (((low >> 14) & 1) | ((low << 1) & 0x7FFF)) ^ chars.length ^ 0xCE4B;
  const key = ((high << 16) | low) >>> 0;
  return [0, 8, 16, 24].map((s) => ((key >>> s) & 0xFF).toString(16).toUpperCase().padStart(2, '0')).join('');
}

const b64 = (bytes) => btoa(String.fromCharCode(...bytes));
const unb64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
const utf16 = (s) => { const out = new Uint8Array(s.length * 2); for (let i = 0; i < s.length; i++) { out[i * 2] = s.charCodeAt(i) & 0xFF; out[i * 2 + 1] = s.charCodeAt(i) >> 8; } return out; };

/** The password's hash as Word keeps it (base64), for that salt, algorithm and spin count. */
export async function hashPassword(password, { salt, algorithm = 'SHA-512', spin = 100000 }) {
  // (SHA-512 and SHA-1, what Word uses, here; others with WebCrypto, slower.)
  const digest = SYNC[algorithm] || (async (b) => new Uint8Array(await globalThis.crypto.subtle.digest(algorithm, b)));
  const key = utf16(legacyKey(password));
  const s = typeof salt === 'string' ? unb64(salt) : salt;
  let h = await digest(new Uint8Array([...s, ...key]));
  const buf = new Uint8Array(h.length + 4);
  for (let i = 0; i < spin; i++) {
    buf.set(h);
    buf[h.length] = i & 0xFF; buf[h.length + 1] = (i >>> 8) & 0xFF; buf[h.length + 2] = (i >>> 16) & 0xFF; buf[h.length + 3] = i >>> 24;
    h = digest(buf);
    if (h instanceof Promise) h = await h;
  }
  return b64(h);
}
const SYNC = { 'SHA-512': sha512, 'SHA-1': sha1 };

/** True when `password` opens this protection tag's lock (always, with no password set). */
export async function checkPassword(tag, password) {
  const h = hashOf(tag);
  if (!h) return true;
  if (!h.algorithm) return false; // (an algorithm browsers lack: only Word can lift it)
  return (await hashPassword(password || '', h)) === h.hash;
}

/**
 * settings.xml with this protection: `edit` (or null: formatting only, or none), `formatting`, enforced, and with a
 * `password` its hash as Word 2013 and later write it (SHA-512, 100,000 rounds). Unenforced, Word's choices stay.
 */
export async function withProtection(settingsXml, { edit = null, formatting = false, enforced = true, password = '' } = {}) {
  let tag = '';
  if (edit || formatting) {
    tag = `<w:documentProtection${edit ? ` w:edit="${edit}"` : ''}${formatting ? ' w:formatting="1"' : ''} w:enforcement="${enforced ? 1 : 0}"`;
    if (enforced && password) {
      const salt = globalThis.crypto.getRandomValues(new Uint8Array(16));
      const hash = await hashPassword(password, { salt, algorithm: 'SHA-512', spin: 100000 });
      tag += ` w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="14" w:cryptSpinCount="100000" w:hash="${hash}" w:salt="${b64(salt)}"`;
    }
    tag += '/>';
  }
  const out = settingsXml.replace(/<w:documentProtection\b[^>]*\/>|<w:documentProtection\b[^>]*>[\s\S]*?<\/w:documentProtection>/g, '');
  if (!tag) return out;
  // Its place among w:settings' children: after w:trackRevisions, w:doNotTrackMoves and w:doNotTrackFormatting.
  let at = -1;
  for (const name of AFTER) {
    const m = new RegExp(`<w:${name}[\\s/>]`).exec(out);
    if (m && (at < 0 || m.index < at)) at = m.index;
  }
  if (at < 0) at = out.lastIndexOf('</w:settings>');
  return out.slice(0, at) + tag + out.slice(at);
}

/** Protection lifted, as Word does: the same choices kept, not enforced, no password. */
export function withoutEnforcement(settingsXml) {
  return settingsXml.replace(/<w:documentProtection\b[^>]*?\/?>/, (t) => {
    const edit = attr(t, 'edit'), formatting = attr(t, 'formatting');
    return `<w:documentProtection${edit ? ` w:edit="${edit}"` : ''}${formatting ? ` w:formatting="${formatting}"` : ''} w:enforcement="0"${t.endsWith('/>') ? '/>' : '>'}`;
  });
}

// w:settings children after w:documentProtection (ECMA-376 Part 1, 17.15.1.78).
const AFTER = ['autoFormatOverride', 'styleLockTheme', 'styleLockQFSet', 'defaultTabStop', 'autoHyphenation', 'consecutiveHyphenLimit', 'hyphenationZone', 'doNotHyphenateCaps', 'showEnvelope', 'summaryLength', 'clickAndTypeStyle', 'defaultTableStyle', 'evenAndOddHeaders', 'bookFoldRevPrinting', 'bookFoldPrinting', 'bookFoldPrintingSheets', 'drawingGridHorizontalSpacing', 'drawingGridVerticalSpacing', 'displayHorizontalDrawingGridEvery', 'displayVerticalDrawingGridEvery', 'doNotUseMarginsForDrawingGridOrigin', 'drawingGridHorizontalOrigin', 'drawingGridVerticalOrigin', 'doNotShadeFormData', 'noPunctuationKerning', 'characterSpacingControl', 'printTwoOnOne', 'strictFirstAndLastChars', 'noLineBreaksAfter', 'noLineBreaksBefore', 'savePreviewPicture', 'doNotValidateAgainstSchema', 'saveInvalidXml', 'ignoreMixedContent', 'alwaysShowPlaceholderText', 'doNotDemarcateInvalidXml', 'saveXmlDataOnly', 'useXSLTWhenSaving', 'saveThroughXslt', 'showXMLTags', 'alwaysMergeEmptyNamespace', 'updateFields', 'hdrShapeDefaults', 'footnotePr', 'endnotePr', 'compat', 'docVars', 'rsids', 'mathPr', 'attachedSchema', 'themeFontLang', 'clrSchemeMapping', 'doNotIncludeSubdocsInStats', 'doNotAutoCompressPictures', 'forceUpgrade', 'captions', 'readModeInkLockDown', 'smartTagType', 'schemaLibrary', 'shapeDefaults', 'doNotEmbedSmartTags', 'decimalSymbol', 'listSeparator'];

/**
 * The styles "Limit formatting to a selection of styles" leaves out, by style id and by lower-case name. As Word reads
 * them: a style with w:locked, and for a built-in style w:latentStyles locking it too (its w:lsdException's w:locked,
 * else w:defLockedState).
 */
export function lockedStyles(stylesXml) {
  const xml = stylesXml || '';
  const latent = /<w:latentStyles\b[^>]*>/.exec(xml)?.[0] || '';
  const defLocked = on(attr(latent, 'defLockedState') || '');
  const exceptions = new Map();
  for (const [t] of xml.matchAll(/<w:lsdException\b[^>]*>/g)) {
    const name = attr(t, 'name'), l = attr(t, 'locked');
    if (name) exceptions.set(name.toLowerCase(), l == null ? defLocked : on(l));
  }
  const out = new Set();
  for (const m of xml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
    const id = attr(m[1], 'styleId');
    const name = (/<w:name w:val="([^"]*)"/.exec(m[2])?.[1] || '').toLowerCase();
    const own = /<w:locked(?![^>]*w:val="(?:0|false|off)")/.test(m[2]);
    const locked = own && (on(attr(m[1], 'customStyle') || '') || (exceptions.get(name) ?? (latent ? defLocked : false)));
    if (locked) { if (id) out.add(id); if (name) out.add(name); }
  }
  return out;
}

/** True when docProps/custom.xml marks the document final (Word's Mark as Final). */
export const markedFinal = (customXml) => /<property\b[^>]*name="_MarkAsFinal"[^>]*>\s*<vt:bool>\s*(?:true|1)\s*<\/vt:bool>/i.test(customXml || '');

/** custom.xml without the mark (Word's Edit Anyway). */
export const withoutFinal = (customXml) => customXml.replace(/<property\b[^>]*name="_MarkAsFinal"[^>]*>[\s\S]*?<\/property>/g, '');
