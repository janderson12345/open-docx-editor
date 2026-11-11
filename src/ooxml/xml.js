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

export function esc(s) {
  return String(s)
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

/** An attribute value's entities decoded (&quot; &apos; &lt; &gt; &amp;), as the XML parser would. */
export const unescapeXml = (s) => (s == null ? s : s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&'));
