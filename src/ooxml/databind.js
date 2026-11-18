// Content controls bound to custom XML (w:dataBinding): a control shows the value of the node its XPath picks in one of
// the package's custom XML parts (customXml/itemN.xml), as Word does, and an edit to the control goes back into that
// node, which every control bound to it then shows. The XPath and the part's id come from the file, so they are read by
// a small resolver of the paths Word writes (/root/a, /ns0:root[1]/ns0:a[1], .../@attr), never a general XPath engine.
import { NS, kids, parseXml, serializeXml, startTag, startTags, tagAttr, setTagAttr, unescapeXml, xmlChars } from './xml.js';

const REL_CUSTOM_XML = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';
const REL_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps';
const NS_DS = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml';
const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
// (Bounds on what a file can make this do: parts read, their size, an XPath's length and steps, the prefixes mapped.)
const MAX_PARTS = 64;
const MAX_PART_BYTES = 8 * 1024 * 1024;
const MAX_TOTAL_BYTES = 16 * 1024 * 1024;
const MAX_XPATH = 2048;
const MAX_STEPS = 64;
const MAX_MAPPINGS = 4096;
const MAX_CACHE = 2048;
// (Elements looked at over all lookups in one store: a crafted part of millions of siblings and thousands of controls
// bound past them would otherwise cost their product. Found nodes are remembered, so only new bindings spend it.)
const MAX_WORK = 2e7;
const MAX_ITEMS = 1000; // a repeating section's items
// (Elements copied into repeating sections' new items over one part: a crafted file of many sections, or sections
// whose items hold large content, would otherwise grow the document a thousandfold or more as it opens.)
const MAX_CLONED = 200000;
const MAX_PICTURE = 16 * 1024 * 1024; // a bound picture's base64 text
// (Bound pictures over a package: each distinct picture decoded and stored once, at most this many and this much.)
const MAX_PICTURES = 100;
const MAX_PICTURE_BYTES = 64 * 1024 * 1024;
const MAX_PH_RPR = 64; // a placeholder building block's run properties copied into controls (elements)
const REL_GLOSSARY = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument';
const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const NS_W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
// Word's placeholders when a control names none it has.
const DEFAULT_PH = { text: 'Click or tap here to enter text.', richText: 'Click or tap here to enter text.', date: 'Click or tap to enter a date.', dropdown: 'Choose an item.', combobox: 'Choose an item.' };

// One step: an element (prefix optional) with an optional position, or, last, an attribute or text().
const NAME = '[\\p{L}_][\\p{L}\\p{N}_.\\-]*';
const STEP = new RegExp(`^(@)?(?:(${NAME}):)?(${NAME})(?:\\[([1-9][0-9]{0,5})\\])?$`, 'u');
const MAPPING = new RegExp(`xmlns:(${NAME})\\s*=\\s*(?:'([^']*)'|"([^"]*)")`, 'gu');

/** A part's id as Word compares them: without braces, upper case. */
const normId = (id) => (id ? String(id).trim().replace(/^\{|\}$/g, '').toUpperCase() : '');

/** The binding in a w:sdtPr's XML: { xpath, store, prefixes } or null. */
export function bindingOf(prXml) {
  const t = typeof prXml === 'string' && prXml.length < 1e6 ? startTag(prXml, 'w:dataBinding') : null;
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

/**
 * The node `steps` pick in a parsed part: { el, attr? } (attr: the attribute node), or null. `budget.left` counts the
 * elements looked at down; at 0 nothing more is found.
 */
function resolve(doc, steps, budget) {
  let node = null;
  for (const s of steps) {
    if (s.text) break;
    if (s.attr) {
      const a = node && node.getAttributeNodeNS(s.ns, s.name);
      return a ? { el: node, attr: a } : null;
    }
    const root = !node;
    let hit = null;
    let k = 0;
    for (let c = root ? doc.documentElement : node.firstChild; c; c = root ? null : c.nextSibling) {
      if (--budget.left < 0) return null;
      if (c.nodeType === 1 && c.localName === s.name && (c.namespaceURI || null) === s.ns && ++k === s.n) { hit = c; break; }
    }
    node = hit;
    if (!node) return null;
  }
  return node ? { el: node } : null;
}

/** Start tag `name`'s attribute `attr` in a control's w:sdtPr XML (as written, still escaped), or null: one scan. */
export const prAttr = (xml, name, attr) => {
  const t = typeof xml === 'string' ? startTag(xml, name) : null;
  return t ? tagAttr(t.tag, attr) : null;
};

/** Whether a control's w:sdtPr XML says it shows its placeholder (a w:showingPlcHdr not set to 0 or false). */
export function showsPlc(xml) {
  if (typeof xml !== 'string') return false;
  for (const t of startTags(xml, 'w:showingPlcHdr')) {
    const v = tagAttr(t.tag, 'w:val');
    if (v !== '0' && v !== 'false') return true;
  }
  return false;
}

/** A control's w:sdtPr XML without its w:showingPlcHdr (scanned, not by a pattern that goes back over unclosed tags). */
export function dropPlc(xml) {
  if (typeof xml !== 'string') return xml;
  let out = '';
  let at = 0;
  for (const t of startTags(xml, 'w:showingPlcHdr')) {
    if (t.at < at) continue;
    out += xml.slice(at, t.at);
    at = t.end;
    if (!t.selfClosing) {
      const c = xml.indexOf('</w:showingPlcHdr>', at);
      if (c >= 0) at = c + '</w:showingPlcHdr>'.length;
    }
  }
  return at ? out + xml.slice(at) : xml;
}

/** A node's value as a bound control shows it. */
const valueOf = (hit) => (hit.attr ? hit.attr.value : hit.el.textContent);

/** The package's custom XML parts, and the values controls bound to them show. */
export class BindingStore {
  constructor(pkg) {
    this.pkg = pkg;
    this.parts = [];
    this._steps = new Map();
    this._hits = new Map();
    this._budget = { left: MAX_WORK };
    this._gen = 0;
    let total = 0;
    const rels = pkg.mainPath ? pkg.readRels(pkg.mainPath).filter((r) => r.type === REL_CUSTOM_XML && !r.external) : [];
    for (const r of rels.slice(0, MAX_PARTS)) {
      const bytes = pkg.files[r.target];
      if (!bytes || bytes.length > MAX_PART_BYTES || total + bytes.length > MAX_TOTAL_BYTES || this.parts.some((p) => p.path === r.target)) continue;
      total += bytes.length;
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
    this._glossary = null;
  }

  /**
   * The placeholder a control shows while its node is empty, as Word shows it: its w:placeholder's building block in
   * the glossary (its text, its first run's w:rPr element), else Word's default for its kind in the PlaceholderText
   * style. { text, rPr (an element, or null) }.
   */
  placeholder(prXml, kind) {
    const t = prXml ? startTag(prXml, 'w:docPart') : null;
    const name = t ? unescapeXml(tagAttr(t.tag, 'w:val') ?? '') : '';
    const part = name ? this._docParts().get(name) : null;
    if (part && part.text) return part;
    return { text: DEFAULT_PH[kind] || DEFAULT_PH.text, rPr: null };
  }

  _docParts() {
    if (this._glossary) return this._glossary;
    const map = (this._glossary = new Map());
    const rel = this.pkg.mainPath && this.pkg.readRels(this.pkg.mainPath).find((r) => r.type === REL_GLOSSARY && !r.external);
    if (!rel || !(this.pkg.files[rel.target]?.length <= MAX_PART_BYTES)) return map;
    let doc;
    try { doc = this.pkg.xml(rel.target); } catch { return map; }
    // (The glossary's own building blocks only, not ones nested in another's text; a block's run formatting kept when
    // it is a few properties, as Word's are: it's copied into every empty control that shows the block.)
    const parts = doc?.documentElement ? kids(kids(doc.documentElement, 'docParts')[0], 'docPart') : [];
    for (let i = 0; i < parts.length && map.size < 10000; i++) {
      const dp = parts[i];
      const nm = kids(kids(dp, 'docPartPr')[0], 'name')[0]?.getAttributeNS(NS.w, 'val');
      const body = kids(dp, 'docPartBody')[0];
      if (!nm || !body || map.has(nm)) continue;
      const text = kids(body, 'p').map(textWithBreaks).join('\n');
      const r = body.getElementsByTagNameNS(NS.w, 'r')[0];
      const rPr = r ? kids(r, 'rPr')[0] || null : null;
      map.set(nm, { text, rPr: rPr && rPr.getElementsByTagName('*').length <= MAX_PH_RPR ? rPr : null });
    }
    return map;
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
    // (Remembered: a set changes nodes' values, never which node a path picks.)
    const key = `${normId(b.store)}\u0000${b.prefixes}\u0000${b.xpath}`;
    if (this._hits.has(key)) return this._hits.get(key);
    const steps = this._compiled(b);
    let found = null;
    if (steps) {
      const id = normId(b.store);
      const named = id && this.parts.find((p) => p.id === id);
      for (const part of named ? [named] : this.parts) {
        const hit = resolve(part.doc, steps, this._budget);
        if (hit) { found = { part, ...hit }; break; }
      }
    }
    if (this._hits.size >= MAX_CACHE * 4) this._hits.clear();
    this._hits.set(key, found);
    return found;
  }

  /** The value a control's binding (its w:sdtPr XML or a binding) shows, or null when it finds nothing. */
  value(b) {
    const hit = this.find(typeof b === 'string' ? bindingOf(b) : b);
    if (!hit) return null;
    // (Read once until the next set: a node's text is all its descendants', which may be the whole part.)
    if (hit.gen !== this._gen) { hit.v = valueOf(hit); hit.gen = this._gen; }
    return hit.v;
  }

  /** The node's identity (for controls bound to the same one), or null. */
  target(b) {
    const hit = this.find(typeof b === 'string' ? bindingOf(b) : b);
    return hit ? (hit.attr || hit.el) : null;
  }

  /**
   * How many nodes a repeating section's binding (its XPath's last step without a position: /root/items/item) finds,
   * in the part chosen as `find` chooses it; null when its parent is found in none.
   */
  count(b) {
    const steps = b && this._compiled(b);
    if (!steps || !steps.length) return null;
    const last = steps.at(-1);
    const id = normId(b.store);
    const named = id && this.parts.find((p) => p.id === id);
    for (const part of named ? [named] : this.parts) {
      const root = part.doc.documentElement;
      if (steps.length === 1) return root.localName === last.name && (root.namespaceURI || null) === last.ns ? 1 : 0;
      const hit = resolve(part.doc, steps.slice(0, -1), this._budget);
      if (!hit || hit.attr) continue;
      // (Counted against the same budget as the lookups: a section bound past millions of siblings, thousands of times.)
      let n = 0;
      for (let c = hit.el.firstChild; c; c = c.nextSibling) {
        if (--this._budget.left < 0) return null;
        if (c.nodeType === 1 && c.localName === last.name && (c.namespaceURI || null) === last.ns) n++;
      }
      return n;
    }
    return null;
  }

  /** Set the value a binding finds; true when it changed. */
  set(b, value) {
    const hit = this.find(typeof b === 'string' ? bindingOf(b) : b);
    const v = xmlChars(value ?? '');
    if (!hit || valueOf(hit) === v) return false;
    this._gen++;
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
  const fmt = unescapeXml(prAttr(pr, 'w:dateFormat', 'w:val') ?? '').slice(0, 256) || 'M/d/yyyy';
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
    const code = (name) => { const v = prAttr(prXml, name, 'w14:val'); return v && /^[0-9A-Fa-f]{1,6}$/.test(v) ? v : null; };
    const hex = on ? code('w14:checkedState') || '2612' : code('w14:uncheckedState') || '2610';
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
  const multi = kind === 'richText' || /^(?:1|true|on)$/.test(prAttr(prXml, 'w:text', 'w:multiLine') ?? '');
  return { text: multi ? v : v.replace(/\n/g, ' ') };
}

function* startTagsIn(xml, name) {
  let n = 0;
  for (let t = startTag(xml, name); t && n < 10000; t = startTag(xml, name, t.end), n++) yield t.tag;
}

/** A control's w:sdtPr XML once it shows a bound value: no placeholder; a check box's state, a date's w:fullDate set. */
export function boundPr(prXml, shown) {
  // (Tags found by scanning, not by patterns that go back over a co-editor's unclosed tags.)
  let pr = dropPlc(prXml);
  const retag = (name, f) => {
    const t = startTag(pr, name);
    if (t) pr = pr.slice(0, t.at) + f(t) + pr.slice(t.end);
  };
  if (shown.checked != null) retag('w14:checked', (t) => setTagAttr(t.tag, 'w14:val', shown.checked ? '1' : '0'));
  if (shown.date) retag('w:date', (t) => `<w:date w:fullDate="${attrEsc(fullDate(shown.date))}"${t.selfClosing ? '/>' : '>'}`);
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
    const v = prAttr(prXml, 'w:storeMappedDataAs', 'w:val');
    const as = v && /^\w{1,32}$/.test(v) ? v : 'dateTime';
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

/** A w:sdtPr's XML with w:showingPlcHdr added (before its binding, where Word writes it). */
export function withPlc(prXml) {
  if (typeof prXml !== 'string' || /<w:showingPlcHdr\b/.test(prXml)) return prXml;
  const at = prXml.search(/<w(?:15)?:dataBinding\b/);
  if (at >= 0) return `${prXml.slice(0, at)}<w:showingPlcHdr/>${prXml.slice(at)}`;
  return prXml.replace(/<\/w:sdtPr>/, '<w:showingPlcHdr/></w:sdtPr>');
}

/**
 * Every bound control in a parsed part (document, header, footer) given the value it is bound to, as Word does when it
 * opens the file: a repeating section given an item for each node it is bound to (the last copied, or the extra ones
 * dropped); a control's cached content replaced by its node's value in the control's own run formatting (w:sdtPr's
 * w:rPr), its placeholder flag dropped; one whose node is empty showing its placeholder (the glossary's building block);
 * a picture control showing the picture its node holds (base64), fitted in the picture's box. A control that finds no
 * node keeps what it has, as in Word; so does one already showing its value. Inline controls, and ones around a
 * paragraph or a table cell, are filled; others (around several cells, a row) are left. `partPath`: the part's path,
 * for a picture's relationship. Returns the number of controls changed.
 */
export function applyBindings(xmlDoc, store, partPath = store?.pkg?.mainPath) {
  if (!store || store.empty) return 0;
  const W = NS.w;
  let changed = repeatSections(xmlDoc, store);
  const sdts = [...xmlDoc.getElementsByTagNameNS(W, 'sdt')];
  for (const sdt of sdts) {
    const pr = kids(sdt, 'sdtPr')[0];
    const content = kids(sdt, 'sdtContent')[0];
    const b = pr && content && bindingOfEl(pr);
    if (!b) continue;
    const value = store.value(b);
    if (value == null) continue;
    const kind = kindOf(pr);
    if (kind === 'picture') { if (value && bindPicture(content, value, store, partPath)) changed++; continue; }
    const plc = kids(pr, 'showingPlcHdr').some((e) => !/^(0|false|off)$/.test(e.getAttributeNS(W, 'val') || ''));
    const prXml = serializeXml(pr);
    let shown = boundDisplay(prXml, kind, value);
    let phRPr = null;
    if (!shown) {
      // An empty node: the placeholder (a check box shows its state, whatever).
      if (plc || kind === 'checkbox') continue;
      const ph = store.placeholder(prXml, kind);
      shown = { text: ph.text };
      phRPr = ph.rPr;
    }
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
    if (!phRPr && !plc && cached === shown.text && value !== '') continue;
    // Formatting: the control's run properties (a check box keeps its symbol font: its first run's); a placeholder's,
    // its building block's (else the PlaceholderText style).
    const doc = sdt.ownerDocument;
    const firstRun = host.getElementsByTagNameNS(W, 'r')[0];
    let rPr;
    if (value === '') {
      rPr = phRPr ? doc.importNode(phRPr, true) : doc.createElementNS(W, 'w:rPr');
      if (!phRPr) { const st = doc.createElementNS(W, 'w:rStyle'); st.setAttributeNS(W, 'w:val', 'PlaceholderText'); rPr.appendChild(st); }
    } else {
      const own = kind === 'checkbox' && firstRun ? kids(firstRun, 'rPr')[0] : kids(pr, 'rPr')[0];
      rPr = own ? own.cloneNode(true) : null;
    }
    const r = doc.createElementNS(W, 'w:r');
    if (rPr) r.appendChild(rPr);
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
    if (value === '') {
      // The placeholder flag, before the binding as Word writes it.
      const flag = doc.createElementNS(W, 'w:showingPlcHdr');
      const before = kids(pr).find((e) => e.localName === 'dataBinding') || null;
      pr.insertBefore(flag, before);
    } else {
      // The placeholder flag dropped; a check box's state and a date's w:fullDate as shown.
      for (const e of kids(pr, 'showingPlcHdr')) pr.removeChild(e);
      if (shown.checked != null) for (const e of pr.getElementsByTagNameNS(NS_W14, 'checked')) e.setAttributeNS(NS_W14, 'w14:val', shown.checked ? '1' : '0');
      if (shown.date) for (const e of kids(pr, 'date')) e.setAttributeNS(W, 'w:fullDate', fullDate(shown.date));
    }
    changed++;
  }
  return changed;
}

/**
 * Repeating sections (w15:repeatingSection) bound to a set of nodes: an item for each node, as Word makes them on
 * opening. Items short are copies of the last one, the bindings inside it moved on to their own node
 * (/root[1]/items[1]/item[2]/name[1] in the second); items over are dropped (one is always kept). Copies get new ids
 * and lose their paragraphs' ids and bookmarks, which may not repeat.
 */
function repeatSections(xmlDoc, store) {
  const W = NS.w;
  let changed = 0;
  let nextId = 0x2f000000;
  let cloned = 0;
  // (A section inside another's item is left as it is: in a crafted file, its marker after its content, it would be
  // filled first and the outer section then copy its copies, a thousand times a thousand.)
  const isSection = (s) => s?.localName === 'sdt' && kids(kids(s, 'sdtPr')[0], 'repeatingSection', NS_W15).length > 0;
  const nested = (s) => {
    for (let a = s.parentNode, d = 0; a && a.nodeType === 1 && d < 256; a = a.parentNode, d++) if (a.localName === 'sdtContent' && isSection(a.parentNode)) return true;
    return false;
  };
  for (const sec of [...xmlDoc.getElementsByTagNameNS(NS_W15, 'repeatingSection')]) {
    const pr = sec.parentNode;
    const sdt = pr?.parentNode;
    if (!pr || pr.localName !== 'sdtPr' || !sdt || sdt.localName !== 'sdt' || nested(sdt)) continue;
    const bind = kids(pr).find((e) => e.localName === 'dataBinding' && (e.namespaceURI === NS_W15 || e.namespaceURI === W));
    const xpath = bind?.getAttributeNS(W, 'xpath');
    const content = kids(sdt, 'sdtContent')[0];
    if (!xpath || !content) continue;
    const b = { xpath, store: bind.getAttributeNS(W, 'storeItemID') || '', prefixes: bind.getAttributeNS(W, 'prefixMappings') || '' };
    const steps = store._compiled(b);
    const last = steps?.at(-1);
    if (!last || last.attr || last.text || /\]\s*$/.test(xpath)) continue; // (the set: its last step has no position)
    const count = store.count(b);
    if (count == null) continue;
    const items = kids(content, 'sdt').filter((it) => kids(kids(it, 'sdtPr')[0], 'repeatingSectionItem', NS_W15).length);
    if (!items.length || items.length !== kids(content).length) continue;
    let want = Math.min(MAX_ITEMS, Math.max(1, count));
    if (want === items.length) continue;
    const model = items[Math.min(items.length, want) - 1];
    const from = Math.min(items.length, want);
    // (New items only while the part's copying budget lasts: a section that would pass it gets the items that fit.)
    if (want > from) {
      const size = model.getElementsByTagName('*').length + 1;
      want = Math.min(want, from + Math.max(0, Math.floor((MAX_CLONED - cloned) / size)));
      cloned += (want - from) * size;
    }
    for (const extra of items.slice(want)) content.removeChild(extra);
    const base = xpath.trim();
    for (let k = from + 1; k <= want; k++) {
      const copy = model.cloneNode(true);
      // Its bindings moved from the model's node to this item's: <section xpath>[from] to [k].
      for (const db of [...copy.getElementsByTagNameNS(W, 'dataBinding'), ...copy.getElementsByTagNameNS(NS_W15, 'dataBinding')]) {
        const x = db.getAttributeNS(W, 'xpath') || '';
        const head = `${base}[${from}]`;
        if (x.startsWith(head) && (x.length === head.length || x[head.length] === '/')) db.setAttributeNS(W, 'w:xpath', `${base}[${k}]${x.slice(head.length)}`);
      }
      for (const id of copy.getElementsByTagNameNS(W, 'id')) if (id.parentNode?.localName === 'sdtPr') id.setAttributeNS(W, 'w:val', String(nextId++));
      for (const p of copy.getElementsByTagNameNS(W, 'p')) { p.removeAttributeNS(NS_W14, 'paraId'); p.removeAttributeNS(NS_W14, 'textId'); }
      for (const name of ['bookmarkStart', 'bookmarkEnd']) for (const bm of [...copy.getElementsByTagNameNS(W, name)]) bm.parentNode.removeChild(bm);
      content.appendChild(copy);
    }
    changed++;
  }
  return changed;
}

// A picture's type and size in pixels, from its first bytes: { mime, ext, w, h } or null.
function pictureInfo(b) {
  const u16 = (i, le) => (le ? b[i] | (b[i + 1] << 8) : (b[i] << 8) | b[i + 1]);
  const u32 = (i, le) => (le ? (b[i] | (b[i + 1] << 8) | (b[i + 2] << 16)) + b[i + 3] * 0x1000000 : b[i] * 0x1000000 + ((b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3]));
  if (b.length > 24 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return { mime: 'image/png', ext: 'png', w: u32(16), h: u32(20) };
  if (b.length > 10 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return { mime: 'image/gif', ext: 'gif', w: u16(6, true), h: u16(8, true) };
  if (b.length > 26 && b[0] === 0x42 && b[1] === 0x4d) return { mime: 'image/bmp', ext: 'bmp', w: u32(18, true), h: Math.abs(u32(22, true) | 0) };
  if (b.length > 4 && b[0] === 0xff && b[1] === 0xd8) {
    // JPEG: the frame header (SOFn) gives the size; segments walked by their lengths, a bounded number of them.
    for (let i = 2, n = 0; i + 9 < b.length && n < 10000; n++) {
      if (b[i] !== 0xff) return null;
      const m = b[i + 1];
      if (m === 0xff) { i++; continue; }
      if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return { mime: 'image/jpeg', ext: 'jpeg', w: u16(i + 7), h: u16(i + 5) };
      i += 2 + u16(i + 2);
    }
  }
  return null;
}

/** A base64 text as bytes, or null (not base64, or too long). */
function fromBase64(s) {
  if (s.length > MAX_PICTURE) return null;
  const clean = s.replace(/\s+/g, '');
  if (!clean || !/^[A-Za-z0-9+/]+={0,2}$/.test(clean)) return null;
  try {
    const bin = atob(clean);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  } catch { return null; }
}

/**
 * A picture control's picture made the one its node holds (base64), as Word shows it: a new image part (the one it
 * showed is left, in case anything else shows it too), fitted in the picture's box keeping its proportions. True
 * when it changed.
 */
function bindPicture(content, value, store, partPath) {
  const pkg = store.pkg;
  const blip = content.getElementsByTagNameNS(NS_A, 'blip')[0];
  const extent = content.getElementsByTagNameNS(NS_WP, 'extent')[0];
  if (!blip || !extent || !partPath) return false;
  // (Each distinct picture decoded and stored once, however many controls show it, a bounded number and size of them
  // over the package: { byValue: value -> { info, path (its part), rels: part -> rId } | null, bytes, n }.)
  const pics = store._pics || (store._pics = { byValue: new Map(), bytes: 0, n: 0 });
  let pic = pics.byValue.get(value);
  if (pic === undefined) {
    const bytes = pics.byValue.size < MAX_PICTURES ? fromBase64(value) : null;
    const info = bytes && pictureInfo(bytes);
    pic = info && info.w > 0 && info.h > 0 && pics.bytes + bytes.length <= MAX_PICTURE_BYTES ? { info, bytes, path: null, rels: new Map() } : null;
    if (pic) pics.bytes += bytes.length;
    pics.byValue.set(value, pic);
  }
  if (!pic) return false;
  const { info } = pic;
  const rId = blip.getAttributeNS(NS_R, 'embed');
  const rel = rId && pkg.readRels(partPath).find((r) => r.id === rId);
  const dir = partPath.slice(0, partPath.lastIndexOf('/') + 1);
  if (rel && !rel.external && rel.target === pic.path) return false;
  if (!pic.path) {
    const had = rel && !rel.external && pkg.files[rel.target];
    if (had && had.length === pic.bytes.length && had.every((x, i) => x === pic.bytes[i])) return false;
    do pics.n++; while (pkg.files[`${dir}media/bound${pics.n}.${info.ext}`]);
    pic.path = `${dir}media/bound${pics.n}.${info.ext}`;
    pkg.files[pic.path] = pic.bytes;
    pkg.addOverride(pic.path, info.mime);
  }
  if (!pic.rels.has(partPath)) {
    // (Relative to the part's folder when it's there, as the part's other pictures are.)
    const target = pic.path.startsWith(dir) ? pic.path.slice(dir.length) : `/${pic.path}`;
    pic.rels.set(partPath, pkg.addRel(REL_IMAGE, target, false, partPath));
  }
  blip.setAttributeNS(NS_R, 'r:embed', pic.rels.get(partPath));
  // Fitted in the box (wp:extent), and the picture's own size (a:ext in pic:spPr) with it.
  const cx = Number(extent.getAttribute('cx')), cy = Number(extent.getAttribute('cy'));
  if (cx > 0 && cy > 0) {
    const scale = Math.min(cx / info.w, cy / info.h);
    const nx = String(Math.max(1, Math.round(info.w * scale))), ny = String(Math.max(1, Math.round(info.h * scale)));
    extent.setAttribute('cx', nx);
    extent.setAttribute('cy', ny);
    const ext = [...content.getElementsByTagNameNS(NS_A, 'ext')].find((e) => e.parentNode?.localName === 'xfrm' && e.hasAttribute('cx'));
    if (ext) { ext.setAttribute('cx', nx); ext.setAttribute('cy', ny); }
  }
  return true;
}

/** A paragraph's or a control's text, its line breaks as '\n'. */
function textWithBreaks(el) {
  // (Walked without recursion: a crafted control nested 100,000 deep would otherwise overflow the stack.)
  let out = '';
  let c = el.firstChild;
  while (c && c !== el) {
    let down = false;
    if (c.nodeType === 1) {
      const w = c.namespaceURI === NS.w;
      if (w && c.localName === 't') out += c.textContent;
      else if (w && (c.localName === 'br' || c.localName === 'cr')) out += '\n';
      else if (w && c.localName === 'tab') out += '\t';
      else down = !(w && (c.localName === 'rPr' || c.localName === 'pPr' || c.localName === 'delText'));
    }
    if (down && c.firstChild) { c = c.firstChild; continue; }
    while (c && c !== el && !c.nextSibling) c = c.parentNode;
    if (!c || c === el) break;
    c = c.nextSibling;
  }
  return out;
}
