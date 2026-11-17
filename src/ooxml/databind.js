// Content controls bound to custom XML (w:dataBinding): a control shows the value of the node its XPath picks in one of
// the package's custom XML parts (customXml/itemN.xml), as Word does, and an edit to the control goes back into that
// node, which every control bound to it then shows. The XPath and the part's id come from the file, so they are read by
// a small resolver of the paths Word writes (/root/a, /ns0:root[1]/ns0:a[1], .../@attr), never a general XPath engine.
import { NS, kids, parseXml, serializeXml, startTag, tagAttr, unescapeXml } from './xml.js';

const REL_CUSTOM_XML = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';
const REL_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps';
const NS_DS = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml';
const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
// (Bounds on what a file can make this do: parts read, their size, an XPath's length and steps, the prefixes mapped.)
const MAX_PARTS = 64;
const MAX_PART_BYTES = 8 * 1024 * 1024;
const MAX_XPATH = 2048;
const MAX_STEPS = 64;
const MAX_MAPPINGS = 4096;
const MAX_CACHE = 2048;

// One step: an element (prefix optional) with an optional position, or, last, an attribute or text().
const NAME = '[\\p{L}_][\\p{L}\\p{N}_.\\-]*';
const STEP = new RegExp(`^(@)?(?:(${NAME}):)?(${NAME})(?:\\[([1-9][0-9]{0,5})\\])?$`, 'u');
const MAPPING = new RegExp(`xmlns:(${NAME})\\s*=\\s*(?:'([^']*)'|"([^"]*)")`, 'gu');

/** A part's id as Word compares them: without braces, upper case. */
const normId = (id) => (id ? String(id).trim().replace(/^\{|\}$/g, '').toUpperCase() : '');

/** The binding in a w:sdtPr's XML: { xpath, store, prefixes } or null. */
export function bindingOf(prXml) {
  const t = prXml && prXml.length < 1e6 ? startTag(prXml, 'w:dataBinding') : null;
  if (!t) return null;
  const get = (n) => unescapeXml(tagAttr(t.tag, n) ?? '');
  const xpath = get('w:xpath');
  return xpath ? { xpath, store: get('w:storeItemID'), prefixes: get('w:prefixMappings') } : null;
}

/** The same from a w:sdtPr element. */
function bindingOfEl(pr) {
  const b = kids(pr, 'dataBinding')[0];
  const xpath = b?.getAttributeNS(NS.w, 'xpath');
  return xpath ? { xpath, store: b.getAttributeNS(NS.w, 'storeItemID') || '', prefixes: b.getAttributeNS(NS.w, 'prefixMappings') || '' } : null;
}

/**
 * An XPath Word writes, read into steps: [{ ns, name, n, attr? , text? }] (n: the position, 1 when none is given),
 * or null for anything else (a wildcard, //, a predicate other than a position, too long).
 */
export function compileXPath(xpath, prefixes = '') {
  if (typeof xpath !== 'string' || xpath.length > MAX_XPATH || xpath[0] !== '/') return null;
  const map = new Map();
  if (prefixes) {
    if (prefixes.length > MAX_MAPPINGS) return null;
    for (const m of prefixes.matchAll(MAPPING)) map.set(m[1], m[2] ?? m[3]);
  }
  const parts = xpath.trim().slice(1).split('/');
  if (!parts.length || parts.length > MAX_STEPS) return null;
  const steps = [];
  for (let i = 0; i < parts.length; i++) {
    const s = parts[i].trim();
    const last = i === parts.length - 1;
    if (last && s === 'text()' && i > 0) { steps.push({ text: true }); continue; }
    const m = STEP.exec(s);
    if (!m) return null;
    const [, at, prefix, name, pos] = m;
    if (at && (!last || pos || i === 0)) return null;
    let ns = null;
    if (prefix) {
      if (!map.has(prefix)) return null;
      ns = map.get(prefix) || null;
    }
    steps.push(at ? { attr: true, ns, name } : { ns, name, n: pos ? Number(pos) : 1 });
  }
  return steps;
}

/** The node `steps` pick in a parsed part: { el, attr? } (attr: the attribute node), or null. */
function resolve(doc, steps) {
  let node = null;
  for (const s of steps) {
    if (s.text) break;
    if (s.attr) {
      const a = node && node.getAttributeNodeNS(s.ns, s.name);
      return a ? { el: node, attr: a } : null;
    }
    const pool = node ? kids(node) : [doc.documentElement];
    let k = 0;
    node = null;
    for (const c of pool) {
      if (c.localName === s.name && (c.namespaceURI || null) === s.ns && ++k === s.n) { node = c; break; }
    }
    if (!node) return null;
  }
  return node ? { el: node } : null;
}

/** A node's value as a bound control shows it. */
const valueOf = (hit) => (hit.attr ? hit.attr.value : hit.el.textContent);

/** The package's custom XML parts, and the values controls bound to them show. */
export class BindingStore {
  constructor(pkg) {
    this.pkg = pkg;
    this.parts = [];
    this._steps = new Map();
    const rels = pkg.mainPath ? pkg.readRels(pkg.mainPath).filter((r) => r.type === REL_CUSTOM_XML && !r.external) : [];
    for (const r of rels.slice(0, MAX_PARTS)) {
      const bytes = pkg.files[r.target];
      if (!bytes || bytes.length > MAX_PART_BYTES || this.parts.some((p) => p.path === r.target)) continue;
      const text = pkg.text(r.target);
      let doc;
      try { doc = parseXml(text.replace(/^\uFEFF/, '')); } catch { continue; }
      if (!doc.documentElement) continue;
      let id = '';
      const props = pkg.readRels(r.target).find((x) => x.type === REL_PROPS && !x.external);
      const pdoc = props && pkg.files[props.target]?.length < 1024 * 1024 ? (() => { try { return pkg.xml(props.target); } catch { return null; } })() : null;
      if (pdoc?.documentElement) id = normId(pdoc.documentElement.getAttributeNS(NS_DS, 'itemID'));
      this.parts.push({ path: r.target, id, doc, text, dirty: false });
    }
  }

  get empty() { return !this.parts.length; }

  _compiled(b) {
    const key = `${b.prefixes}\u0000${b.xpath}`;
    if (!this._steps.has(key)) {
      if (this._steps.size >= MAX_CACHE) this._steps.clear();
      this._steps.set(key, compileXPath(b.xpath, b.prefixes));
    }
    return this._steps.get(key);
  }

  /**
   * The node a binding picks, as Word finds it: in the part its storeItemID names, when one has that id (and nowhere
   * else); otherwise in the first part where the XPath finds one. { part, el, attr? } or null.
   */
  find(b) {
    if (!b || this.empty) return null;
    const steps = this._compiled(b);
    if (!steps) return null;
    const id = normId(b.store);
    const named = id && this.parts.find((p) => p.id === id);
    for (const part of named ? [named] : this.parts) {
      const hit = resolve(part.doc, steps);
      if (hit) return { part, ...hit };
    }
    return null;
  }

  /** The value a control's binding (its w:sdtPr XML or a binding) shows, or null when it finds nothing. */
  value(b) {
    const hit = this.find(typeof b === 'string' ? bindingOf(b) : b);
    return hit ? valueOf(hit) : null;
  }

  /** The node's identity (for controls bound to the same one), or null. */
  target(b) {
    const hit = this.find(typeof b === 'string' ? bindingOf(b) : b);
    return hit ? (hit.attr || hit.el) : null;
  }

  /** Set the value a binding finds; true when it changed. */
  set(b, value) {
    const hit = this.find(typeof b === 'string' ? bindingOf(b) : b);
    const v = String(value ?? '');
    if (!hit || valueOf(hit) === v) return false;
    if (hit.attr) hit.attr.value = v;
    else {
      while (hit.el.firstChild) hit.el.removeChild(hit.el.firstChild);
      if (v) hit.el.appendChild(hit.el.ownerDocument.createTextNode(v));
    }
    hit.part.dirty = true;
    return true;
  }

  /** Changed parts written back into the package (their XML declaration kept); the others are left as they were. */
  save() {
    for (const p of this.parts) {
      if (!p.dirty) continue;
      const decl = /^\uFEFF?(<\?xml[^?]*\?>)/.exec(p.text)?.[1] || '';
      this.pkg.setText(p.path, decl + serializeXml(p.doc.documentElement));
      p.dirty = false;
    }
  }
}

// --- What a bound control shows -------------------------------------------------------------------------------------

/** A date in the control's w:dateFormat (Word patterns such as M/d/yyyy, MMMM d, yyyy, dd-MMM-yy). */
export function formatDate(d, pr = '') {
  const fmt = (/<w:dateFormat\b[^>]*w:val="([^"]*)"/.exec(pr) || [])[1] || 'M/d/yyyy';
  const M = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const D = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const p2 = (n) => String(n).padStart(2, '0');
  const parts = { yyyy: d.getFullYear(), yy: p2(d.getFullYear() % 100), MMMM: M[d.getMonth()], MMM: M[d.getMonth()].slice(0, 3), MM: p2(d.getMonth() + 1), M: d.getMonth() + 1, dddd: D[d.getDay()], ddd: D[d.getDay()].slice(0, 3), dd: p2(d.getDate()), d: d.getDate() };
  return fmt.replace(/'([^']*)'|yyyy|yy|MMMM|MMM|MM|M|dddd|ddd|dd|d/g, (m, lit) => (lit != null ? lit : parts[m]));
}

/**
 * An XML date (xsd:date or xsd:dateTime) as a local Date, or null. One in UTC (Z) or with an offset is shown in local
 * time, as Word shows it (2012-10-03T00:00:00Z is October 2 in America); one without, as written.
 */
export function parseXmlDate(v) {
  const m = /^\s*(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.\d{1,9})?)?)?(Z|[+-]\d{2}:\d{2})?\s*$/.exec(String(v ?? ''));
  if (!m) return null;
  const [y, mo, d, h = 0, mi = 0, s = 0] = m.slice(1, 7).map((x) => Number(x || 0));
  if (mo < 1 || mo > 12 || d < 1 || d > 31 || h > 24 || mi > 59 || s > 60) return null;
  let date;
  if (m[7]) {
    const off = m[7] === 'Z' ? 0 : (m[7][0] === '-' ? -1 : 1) * (Number(m[7].slice(1, 3)) * 60 + Number(m[7].slice(4, 6)));
    date = new Date(Date.UTC(y, mo - 1, d, h, mi, s) - off * 60000);
  } else date = new Date(y, mo - 1, d, h, mi, s);
  return isNaN(date) ? null : date;
}

const fullDate = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}T00:00:00Z`;
const attrEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

/**
 * What a control bound to `value` shows, by its kind (reader.js fieldInfo): { text, checked?, date? }, or null when it
 * shows what it has (an empty value: its placeholder; a picture). Text keeps its line breaks ('\n') where the control
 * takes several lines (a multi-line plain text control, a rich text one); a single-line one shows each as a space.
 */
export function boundDisplay(prXml, kind, value) {
  if (value == null || value === '' || kind === 'picture') return null;
  const v = String(value).replace(/\r\n?/g, '\n');
  if (kind === 'checkbox') {
    const on = /^\s*(true|1)\s*$/i.test(v);
    const code = (re) => (re.exec(prXml) || [])[1];
    const hex = on ? code(/<w14:checkedState\b[^>]*w14:val="([0-9A-Fa-f]{1,6})"/) || '2612' : code(/<w14:uncheckedState\b[^>]*w14:val="([0-9A-Fa-f]{1,6})"/) || '2610';
    const cp = parseInt(hex, 16);
    return { text: String.fromCodePoint(cp > 0 && cp <= 0x10ffff ? cp : 0x2610), checked: on };
  }
  if (kind === 'dropdown' || kind === 'combobox') {
    for (const t of startTagsIn(prXml, 'w:listItem')) {
      if (unescapeXml(tagAttr(t, 'w:value') ?? '') === v) return { text: unescapeXml(tagAttr(t, 'w:displayText') ?? v) };
    }
    return { text: v };
  }
  if (kind === 'date') {
    const d = parseXmlDate(v);
    return d ? { text: formatDate(d, prXml), date: d } : { text: v };
  }
  const multi = kind === 'richText' || /<w:text\b[^>]*w:multiLine="(?:1|true|on)"/.test(prXml);
  return { text: multi ? v : v.replace(/\n/g, ' ') };
}

function* startTagsIn(xml, name) {
  let n = 0;
  for (let t = startTag(xml, name); t && n < 10000; t = startTag(xml, name, t.end), n++) yield t.tag;
}

/** A control's w:sdtPr XML once it shows a bound value: no placeholder; a check box's state, a date's w:fullDate set. */
export function boundPr(prXml, shown) {
  let pr = prXml.replace(/<w:showingPlcHdr\b[^>]*\/>/g, '');
  if (shown.checked != null) pr = pr.replace(/(<w14:checked\b[^>]*w14:val=")[^"]*"/, `$1${shown.checked ? 1 : 0}"`);
  if (shown.date) pr = pr.replace(/(<w:date\b)[^>]*?(\/?>)/, (m0, open, close) => `${open} w:fullDate="${attrEsc(fullDate(shown.date))}"${close}`);
  return pr;
}

/**
 * The value an edited control puts in its node: a check box's state as true or false, a list's item value, a date as
 * its storeMappedDataAs says (the date and time Word writes by default), otherwise its text.
 */
export function storeValue(prXml, kind, { text, checked, value }) {
  if (kind === 'checkbox') return checked ? 'true' : 'false';
  if (kind === 'dropdown' || kind === 'combobox') return value ?? text;
  if (kind === 'date' && text) {
    const as = (/<w:storeMappedDataAs\b[^>]*w:val="(\w+)"/.exec(prXml) || [])[1] || 'dateTime';
    if (as === 'text') return text;
    // The date picked (w:fullDate) while the text still shows it; a date typed in instead, read as one if it is.
    let full = (/w:fullDate="(\d{4}-\d\d-\d\d)/.exec(prXml) || [])[1];
    const picked = full && parseXmlDate(full);
    if (!picked || formatDate(picked, prXml) !== text) {
      const d = text.length < 64 ? new Date(text) : null;
      full = d && !isNaN(d) ? fullDate(d).slice(0, 10) : null;
    }
    if (!full) return text;
    return as === 'date' ? full : `${full}T00:00:00`;
  }
  return text;
}

// --- On opening: cached content made the bound values ---------------------------------------------------------------

const kindOf = (pr) => {
  const has = (ns, n) => pr.getElementsByTagNameNS(ns, n).length > 0;
  return has(NS_W14, 'checkbox') ? 'checkbox' : has(NS.w, 'dropDownList') ? 'dropdown' : has(NS.w, 'comboBox') ? 'combobox'
    : has(NS.w, 'date') ? 'date' : has(NS.w, 'picture') ? 'picture' : has(NS.w, 'text') ? 'text' : 'richText';
};

/**
 * Every bound control in a parsed part (document, header, footer) given the value it is bound to, as Word does when it
 * opens the file: its cached content replaced by the value, in the control's own run formatting (w:sdtPr's w:rPr), its
 * placeholder flag dropped. A control that finds no node, or whose node is empty, keeps what it has, as in Word; so
 * does one already showing its value. Inline controls, and ones around a paragraph or a table cell, are filled; others
 * (around several cells, a row) are left. Returns the number of controls changed.
 */
export function applyBindings(xmlDoc, store) {
  if (!store || store.empty) return 0;
  const W = NS.w;
  let changed = 0;
  const sdts = [...xmlDoc.getElementsByTagNameNS(W, 'sdt')];
  for (const sdt of sdts) {
    const pr = kids(sdt, 'sdtPr')[0];
    const content = kids(sdt, 'sdtContent')[0];
    const b = pr && content && bindingOfEl(pr);
    if (!b) continue;
    const value = store.value(b);
    const kind = kindOf(pr);
    const prXml = serializeXml(pr);
    const shown = boundDisplay(prXml, kind, value);
    if (!shown) continue;
    const plc = kids(pr, 'showingPlcHdr').some((e) => !/^(0|false|off)$/.test(e.getAttributeNS(W, 'val') || ''));
    const inner = kids(content);
    // Where the runs go: the control's own content (inline), or its paragraph's (around a paragraph or a cell).
    let host = content;
    const first = inner[0];
    if (first && first.namespaceURI === W && (first.localName === 'p' || first.localName === 'tc')) {
      if (inner.length !== 1) continue;
      const p = first.localName === 'tc' ? kids(first, 'p')[0] : first;
      if (!p) continue;
      host = p;
    } else if (inner.some((c) => c.namespaceURI === W && (c.localName === 'p' || c.localName === 'tc' || c.localName === 'tr' || c.localName === 'tbl' || c.localName === 'sdt'))) continue;
    const cached = textWithBreaks(host);
    if (!plc && cached === shown.text) continue;
    // Formatting: the control's run properties (a check box keeps its symbol font: its first run's).
    const firstRun = host.getElementsByTagNameNS(W, 'r')[0];
    const rPr = kind === 'checkbox' && firstRun ? kids(firstRun, 'rPr')[0] : kids(pr, 'rPr')[0];
    const doc = sdt.ownerDocument;
    const r = doc.createElementNS(W, 'w:r');
    if (rPr) r.appendChild(rPr.cloneNode(true));
    shown.text.split('\n').forEach((line, i) => {
      if (i) r.appendChild(doc.createElementNS(W, 'w:br'));
      if (!line) return;
      const t = doc.createElementNS(W, 'w:t');
      t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      t.appendChild(doc.createTextNode(line));
      r.appendChild(t);
    });
    // Out with the runs (and what goes with them: proofing marks, the placeholder's text); a paragraph keeps its pPr,
    // a cell its other paragraphs dropped (a plain text control holds one).
    if (host !== content && first.localName === 'tc') for (const extra of kids(first, 'p').slice(1)) first.removeChild(extra);
    for (const c of [...host.childNodes]) if (!(c.nodeType === 1 && c.namespaceURI === W && c.localName === 'pPr')) host.removeChild(c);
    host.appendChild(r);
    // The placeholder flag dropped; a check box's state and a date's w:fullDate as shown.
    for (const e of kids(pr, 'showingPlcHdr')) pr.removeChild(e);
    if (shown.checked != null) for (const e of pr.getElementsByTagNameNS(NS_W14, 'checked')) e.setAttributeNS(NS_W14, 'w14:val', shown.checked ? '1' : '0');
    if (shown.date) for (const e of kids(pr, 'date')) e.setAttributeNS(W, 'w:fullDate', fullDate(shown.date));
    changed++;
  }
  return changed;
}

/** A paragraph's or a control's text, its line breaks as '\n'. */
function textWithBreaks(el) {
  let out = '';
  const walk = (n) => {
    for (let c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      if (c.namespaceURI === NS.w && c.localName === 't') out += c.textContent;
      else if (c.namespaceURI === NS.w && (c.localName === 'br' || c.localName === 'cr')) out += '\n';
      else if (c.namespaceURI === NS.w && c.localName === 'tab') out += '\t';
      else if (!(c.namespaceURI === NS.w && (c.localName === 'rPr' || c.localName === 'pPr' || c.localName === 'delText'))) walk(c);
    }
  };
  walk(el);
  return out;
}
