// Small XML helpers shared by the reader and writer.
// The DOMParser / XMLSerializer implementations are injectable so the same code
// runs in the browser (native) and in Node (@xmldom/xmldom) for tests.

export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
};

// The prefixes Word declares on a part's root, for XML kept from a part and parsed again on its own (a table's w:tblPr,
// a run's w:rPr inside a tracked change): their declarations were left on the root (outerXml), and a browser won't
// parse a prefix nothing declares (Word 2023's w16du:dateUtc on a tracked table change stopped documents opening).
const PREFIX_NS = {
  w: NS.w, r: NS.r, a: NS.a, wp: NS.wp, pic: NS.pic,
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  v: 'urn:schemas-microsoft-com:vml',
  o: 'urn:schemas-microsoft-com:office:office',
  w10: 'urn:schemas-microsoft-com:office:word',
  wne: 'http://schemas.microsoft.com/office/word/2006/wordml',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
  w16: 'http://schemas.microsoft.com/office/word/2018/wordml',
  w16cex: 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
  w16cid: 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
  w16du: 'http://schemas.microsoft.com/office/word/2023/wordml/word16du',
  w16se: 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
  w16sdtdh: 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
  w16sdtfl: 'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock',
  wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
};

/**
 * Declarations (` xmlns:p="..."`) for the prefixes `xml` uses that `declared` (prefixes already declared where it will
 * be parsed) doesn't hold: Word's own namespaces, and a stand-in for any other, so the piece parses. (Read only: a
 * piece written back goes without them, under the part's root that declares them.)
 */
export function missingDecls(xml, declared = []) {
  const have = new Set(['xml', 'xmlns', ...declared]);
  // (Not again those its first tag declares: the declarations may be added to that tag, and twice is an error.)
  // (Its end found past quoted values: a ">" in one isn't the tag's end. The three branches start on different
  // characters, so the match is linear.)
  const first = /^\s*<(?:[^>"']|"[^"]*"|'[^']*')*>/.exec(String(xml))?.[0] || '';
  for (const m of first.matchAll(/\sxmlns:([A-Za-z_][\w.-]*)\s*=/g)) have.add(m[1]);
  let out = '';
  for (const m of String(xml).matchAll(/(?:<\/?|\s)([A-Za-z_][\w.-]*):[A-Za-z_]/g)) {
    const p = m[1];
    if (have.has(p)) continue;
    have.add(p);
    if (have.size > 300) break;
    out += ` xmlns:${p}="${Object.hasOwn(PREFIX_NS, p) ? PREFIX_NS[p] : `urn:odx:undeclared:${p}`}"`;
  }
  return out;
}

export const REL_TYPES = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  settings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  aFChunk: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk',
};

let impl = {
  DOMParser: typeof DOMParser !== 'undefined' ? DOMParser : null,
  XMLSerializer: typeof XMLSerializer !== 'undefined' ? XMLSerializer : null,
};

/** Override the XML implementation (used by Node tests). */
export function setXmlImpl({ DOMParser: P, XMLSerializer: S }) {
  impl = { DOMParser: P, XMLSerializer: S };
}

export function parseXml(text) {
  const doc = new impl.DOMParser().parseFromString(text, 'application/xml');
  const err = doc.getElementsByTagName('parsererror')[0];
  if (err) throw new Error('Invalid XML: ' + err.textContent);
  return doc;
}

/** HTML read as a browser reads it (errors recovered from), for imported content (altchunk.js). */
export function parseHtml(text) {
  return new impl.DOMParser().parseFromString(text, 'text/html');
}

export function serializeXml(node) {
  return new impl.XMLSerializer().serializeToString(node);
}

export const isEl = (n) => n && n.nodeType === 1;

/** Element children, optionally filtered by w: local name. */
export function kids(el, name, ns = NS.w) {
  const out = [];
  if (!el) return out;
  for (let c = el.firstChild; c; c = c.nextSibling) {
    if (!isEl(c)) continue;
    if (name && (c.localName !== name || c.namespaceURI !== ns)) continue;
    out.push(c);
  }
  return out;
}

export const kid = (el, name, ns = NS.w) => kids(el, name, ns)[0] || null;

/** Read a w:-namespaced attribute (falls back to the unprefixed name). */
export function attr(el, name, ns = NS.w) {
  if (!el) return null;
  if (el.hasAttributeNS && el.hasAttributeNS(ns, name)) return el.getAttributeNS(ns, name);
  return el.hasAttribute && el.hasAttribute(name) ? el.getAttribute(name) : null;
}

/** OOXML on/off properties: <w:b/> is on, <w:b w:val="0|false|off"/> is off. */
export function onOff(el) {
  if (!el) return null;
  const v = attr(el, 'val');
  return !(v === '0' || v === 'false' || v === 'off');
}

/** Characters XML can't hold (C0 controls but tab and newlines, U+FFFE and U+FFFF, lone surrogates). */
const XML_BAD = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;

/** Text with the characters XML can't hold left out (a file with one is one Word won't open). */
export const xmlChars = (s) => String(s).replace(XML_BAD, '');

export function esc(s) {
  return xmlChars(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Serialize an element and drop xmlns declarations that the document root already makes. */
export function outerXml(el, rootNs) {
  let s = serializeXml(el);
  s = s.replace(/\s+xmlns(?::([\w.-]+))?="([^"]*)"/g, (m, pfx, uri) =>
    rootNs[pfx || ''] === uri ? '' : m,
  );
  return s;
}

/** Map of prefix -> namespace URI declared on an element. */
export function nsDecls(el) {
  const out = {};
  for (const a of Array.from(el.attributes)) {
    if (a.name === 'xmlns') out[''] = a.value;
    else if (a.name.startsWith('xmlns:')) out[a.name.slice(6)] = a.value;
  }
  return out;
}

export const textOf = (el) => (el ? el.textContent || '' : '');

/** An element's child elements, whatever their namespace. */
export const elementKids = (el) => (el ? [...el.childNodes].filter((n) => n.nodeType === 1) : []);
/** Its first child element with this local name (any namespace), or null. */
export const childNamed = (el, name) => elementKids(el).find((n) => n.localName === name) || null;
/** Its child elements with this local name (any namespace). */
export const childrenNamed = (el, name) => elementKids(el).filter((n) => n.localName === name);

/**
 * An attribute value's or text's entities decoded (&quot; &apos; &lt; &gt; &#nn; &#xhh; and &amp; last, so "&amp;lt;"
 * stays "&lt;"), as the XML parser would. The one place XML read with regexes is unescaped: track.js, merge.js,
 * refs.js and svgfonts.js each had their own, and one had &amp; in the wrong place.
 */
const codePoint = (n) => (n >= 0 && n <= 0x10ffff ? String.fromCodePoint(n) : '\uFFFD');
export const unescapeXml = (s) => (s == null ? s : s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>')
  .replace(/&#(\d+);/g, (m, d) => codePoint(Number(d))).replace(/&#x([\da-f]+);/gi, (m, h) => codePoint(parseInt(h, 16))).replace(/&amp;/g, '&'));

// Start tags and their attributes read from XML text in one pass each, for the places that read a part or a run's XML
// as text. A regex reading two attributes of one tag, <x\b[^>]*a="..."[^>]*b="...", tries the second [^>]* at every
// place the first could stop: a crafted tag repeating a="1" inside another attribute's single-quoted value (valid XML)
// took seconds (140KB: 2s; a 300KB styles.xml: 3s), growing with the square of its length.

const NAME_END = /[\s/>]/;

/**
 * The first start tag named `name` (its qualified name, as written) at or after `from`: { tag (its text, to its '>'),
 * at, end (just past it), selfClosing }, or null. Quoted attribute values may hold '>'.
 */
export function startTag(xml, name, from = 0) {
  const open = `<${name}`;
  for (let i = xml.indexOf(open, from); i >= 0; i = xml.indexOf(open, i + 1)) {
    const after = xml[i + open.length];
    if (after !== undefined && !NAME_END.test(after)) continue; // (a longer name: <w:styles for <w:style)
    let q = null, j = i + open.length;
    for (; j < xml.length; j++) {
      const c = xml[j];
      if (q) { if (c === q) q = null; } else if (c === '"' || c === "'") q = c; else if (c === '>') break;
    }
    if (j >= xml.length) return null;
    return { tag: xml.slice(i, j + 1), at: i, end: j + 1, selfClosing: xml[j - 1] === '/' };
  }
  return null;
}

/** Every start tag named `name` in `xml`, in order (startTag's results). */
export function* startTags(xml, name) {
  for (let t = startTag(xml, name); t; t = startTag(xml, name, t.end)) yield t;
}

/** Where attribute `name` (its qualified name) is in a start tag's text: [value start, value end], or null. */
function attrSpan(tag, name) {
  let i = tag.search(NAME_END);
  if (i < 0) return null;
  while (i < tag.length) {
    while (i < tag.length && /\s/.test(tag[i])) i++;
    let k = i;
    while (k < tag.length && !/[\s=/>]/.test(tag[k])) k++;
    if (k === i) return null;
    const attr = tag.slice(i, k);
    while (k < tag.length && /\s/.test(tag[k])) k++;
    if (tag[k] !== '=') return null;
    k++;
    while (k < tag.length && /\s/.test(tag[k])) k++;
    const q = tag[k];
    if (q !== '"' && q !== "'") return null;
    const end = tag.indexOf(q, k + 1);
    if (end < 0) return null;
    if (attr === name) return [k + 1, end];
    i = end + 1;
  }
  return null;
}

/** Attribute `name`'s value in a start tag's text, as written (entities not decoded), or null. */
export const tagAttr = (tag, name) => { const s = tag && attrSpan(tag, name); return s ? tag.slice(s[0], s[1]) : null; };

/** The start tag with attribute `name`'s value replaced by `value` (written as given), or as it was without one. */
export const setTagAttr = (tag, name, value) => { const s = attrSpan(tag, name); return s ? tag.slice(0, s[0]) + value + tag.slice(s[1]) : tag; };

/** A start tag's element as text: itself when self-closing, else to the first `</name>` after it (null when none). */
export const elementText = (xml, t, name) => {
  if (t.selfClosing) return t.tag;
  const close = xml.indexOf(`</${name}>`, t.end);
  return close < 0 ? null : xml.slice(t.at, close + name.length + 3);
};
