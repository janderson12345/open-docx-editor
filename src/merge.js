// Mail merge, as Word's Mailings > Finish & Merge > Edit Individual Documents: a document's merge fields (MERGEFIELD,
// with MERGEREC and MERGESEQ) filled from each record of a data source, giving one document per record or all of them
// in one, each record starting on a new page (a next-page section break, as Word joins them). The fields become plain
// text in the result, in their results' formatting, as in Word's merged document. Works without an editor (in Node,
// call setXmlImpl(...) with a DOMParser/XMLSerializer first) and in the browser.
import { Fragment } from 'prosemirror-model';
import { DocxSession } from './ooxml/docx.js';
import { schema } from './editor/schema.js';
import { marksFromRPr } from './ooxml/reader.js';
import { withSectType } from './editor/commands.js';
import { unescapeXml, parseXml, serializeXml, kids, NS } from './ooxml/xml.js';

const fldType = (n) => (n.type === schema.nodes.opaque_inline && n.attrs.kind === 'fldChar' ? (/fldCharType="(\w+)"/.exec(n.attrs.xml) || [])[1] : null);
const instrOf = (n) => (n.type === schema.nodes.opaque_inline && n.attrs.kind === 'instrText' ? unescapeXml((/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/.exec(n.attrs.xml) || [])[1] || '') : null);

/** A field instruction's name and its merge field's name: ` MERGEFIELD "First Name" \* Upper ` -> MERGEFIELD, First Name. */
export function parseMergeField(instr) {
  const m = /^\s*(MERGEFIELD|MERGEREC|MERGESEQ)\b(?:\s+(?:"([^"]*)"|([^\s\\]+)))?/i.exec(instr || '');
  return m ? { kind: m[1].toUpperCase(), name: m[2] ?? m[3] ?? '' } : null;
}

/**
 * The document's merge fields in order: [{ kind, name, instr, from, to, resultFrom, resultTo, simple }]. A complex
 * field runs from its begin fldChar (`from`) to after its end (`to`); a simple one (w:fldSimple) is one node.
 */
export function mergeFieldsIn(doc) {
  const out = [];
  const stack = [];
  doc.descendants((n, pos) => {
    if (n.type === schema.nodes.opaque_inline && n.attrs.kind === 'fldSimple') {
      const instr = unescapeXml((/w:instr="([^"]*)"/.exec(n.attrs.xml) || [])[1] || '');
      const f = parseMergeField(instr);
      if (f) out.push({ ...f, instr, from: pos, to: pos + n.nodeSize, resultFrom: null, resultTo: null, simple: n });
      return false;
    }
    const t = fldType(n);
    if (t === 'begin') stack.push({ instr: '', from: pos, resultFrom: null });
    else if (t === 'separate' && stack.length) stack[stack.length - 1].resultFrom = pos + n.nodeSize;
    else if (t === 'end' && stack.length) {
      const f = stack.pop();
      const m = parseMergeField(f.instr);
      if (m) out.push({ ...m, instr: f.instr, from: f.from, to: pos + n.nodeSize, resultFrom: f.resultFrom, resultTo: f.resultFrom == null ? null : pos, simple: null });
    } else if (stack.length && stack[stack.length - 1].resultFrom == null) {
      const i = instrOf(n);
      if (i != null) stack[stack.length - 1].instr += i;
    }
    return true;
  });
  return out.sort((a, b) => a.from - b.from);
}

/** The rIds of the document's headers and footers. */
const storyIds = (session) => session.pkg.rels.filter((r) => /\/(?:header|footer)$/.test(r.type)).map((r) => r.id);

/** The names of a .docx's merge fields (MERGEFIELD), each once, in order: the body's, then its headers' and footers'. */
export function mergeFieldNames(bytes) {
  const session = bytes instanceof DocxSession ? bytes : new DocxSession(bytes);
  const docs = [session.doc, ...storyIds(session).map((id) => session.story(id)).filter(Boolean)];
  return [...new Set(docs.flatMap((d) => mergeFieldsIn(d)).filter((f) => f.kind === 'MERGEFIELD').map((f) => f.name))];
}

// --- Fields as Word nests them ----------------------------------------------------------------------------------------

/**
 * The document's fields as Word nests them: [{ parts, from, to, resultFrom, resultTo, simple, kids }], `parts` the
 * field code's text and the fields nested in it ({ IF { MERGEFIELD Amount } > 100 ... }), `kids` those in its result.
 */
export function fieldTree(doc) {
  const top = [];
  const stack = [];
  const add = (f) => {
    const p = stack[stack.length - 1];
    if (!p) top.push(f);
    else if (p.resultFrom == null) p.parts.push(f);
    else p.kids.push(f);
  };
  doc.descendants((n, pos) => {
    if (n.type === schema.nodes.opaque_inline && n.attrs.kind === 'fldSimple') {
      const instr = unescapeXml((/w:instr="([^"]*)"/.exec(n.attrs.xml) || [])[1] || '');
      add({ parts: [instr], from: pos, to: pos + n.nodeSize, resultFrom: null, resultTo: null, simple: n, kids: [] });
      return false;
    }
    const t = fldType(n);
    const f = stack[stack.length - 1];
    if (t === 'begin') stack.push({ parts: [], from: pos, resultFrom: null, simple: null, kids: [] });
    else if (t === 'separate' && f) f.resultFrom = pos + n.nodeSize;
    else if (t === 'end' && f) {
      stack.pop();
      Object.assign(f, { to: pos + n.nodeSize, resultTo: f.resultFrom == null ? null : pos });
      add(f);
    } else if (f && f.resultFrom == null) {
      const i = instrOf(n);
      if (i != null) { if (typeof f.parts[f.parts.length - 1] === 'string') f.parts[f.parts.length - 1] += i; else f.parts.push(i); }
    }
    return true;
  });
  return top;
}

const codeOf = (f) => f.parts.filter((p) => typeof p === 'string').join('');
const nested = (f) => f.parts.filter((p) => typeof p !== 'string');
const isIf = (f) => /^\s*IF\b/i.test(codeOf(f));
const usesMerge = (f) => !!parseMergeField(codeOf(f)) || nested(f).some(usesMerge);
/** NEXT, NEXTIF or SKIPIF: the fields that move a merge on to the next record or leave the record out. */
const stepOf = (f) => /^\s*(NEXTIF|SKIPIF|NEXT)\b/i.exec(codeOf(f))?.[1].toUpperCase() ?? null;

/**
 * Where a merge is in its records, as one merged document is filled: `records`, the current one (`i`, moved on by
 * NEXT and NEXTIF; past the last, every field is empty, as in Word), `seq` (MERGESEQ: the merged documents so far,
 * this one included) and `skip` (a SKIPIF was true: the document is left out). `preview`: Preview Results, where Word
 * leaves MERGESEQ empty and SKIPIF has no effect.
 */
export function mergeCursor(records, i = 0, { seq = i + 1, preview = false } = {}) {
  return { records, i, seq, preview, skip: false };
}
const current = (ctx) => ctx.records[ctx.i];

/** A field's value for the record: a merge field's, an IF's text, any other field's shown result. */
function fieldValue(doc, f, ctx) {
  const code = codeOf(f);
  const m = parseMergeField(code);
  if (m) {
    if (m.kind === 'MERGEFIELD') return formatMergeValue(recordValue(current(ctx), m.name), code);
    if (!current(ctx)) return '';
    return m.kind === 'MERGEREC' ? String(ctx.i + 1) : ctx.preview ? '' : String(ctx.seq);
  }
  if (isIf(f)) return ifValue(doc, f, ctx);
  return f.resultFrom != null && f.resultTo > f.resultFrom ? doc.textBetween(f.resultFrom, f.resultTo, '', '') : '';
}

/**
 * An IF field's code in words, as Word reads it: split at spaces outside quotes, \" a quote in quotes, a nested
 * field's value one word however many spaces it has ({ MERGEFIELD City } = "New York" is true for New York in Word).
 */
function ifWords(doc, f, ctx) {
  const words = [];
  let cur = null;
  let q = false;
  const end = () => { if (cur) words.push(cur); cur = null; };
  for (const p of f.parts) {
    if (typeof p !== 'string') { (cur ||= { t: '', quoted: q }).t += fieldValue(doc, p, ctx); continue; }
    for (let i = 0; i < p.length; i++) {
      const c = p[i];
      if (q) {
        if (c === '\\' && (p[i + 1] === '"' || p[i + 1] === '\\')) cur.t += p[++i];
        else if (c === '"') { q = false; end(); } else cur.t += c;
      } else if (c === '"') { end(); q = true; cur = { t: '', quoted: true }; } else if (/\s/.test(c)) end();
      else (cur ||= { t: '', quoted: false }).t += c;
    }
  }
  end();
  return words;
}

const asNumber = (s) => { const t = String(s).replace(/[$,\s]/g, ''); return t !== '' && Number.isFinite(Number(t)) ? Number(t) : null; };

/**
 * Whether `text` is all of `pattern`, its `*` any run of characters and its `?` any one, in the same case. Without a
 * regular expression: one built from a document's pattern backtracks without end on a crafted one ("*a*a*a*a*a*a*b"
 * took seconds on 50 characters), where this goes back only to the last `*`, at most the two lengths multiplied.
 */
function wildcard(text, pattern) {
  // (Operands far longer than any field's: no match, rather than the seconds two of 40,000 characters took.)
  if (text.length * pattern.length > 1e7) return false;
  let t = 0, p = 0, star = -1, mark = 0;
  while (t < text.length) {
    if (pattern[p] === '*') { star = p++; mark = t; }
    else if (p < pattern.length && (pattern[p] === '?' || pattern[p] === text[t])) { t++; p++; }
    else if (star >= 0) { p = star + 1; t = ++mark; }
    else return false;
  }
  while (pattern[p] === '*') p++;
  return p === pattern.length;
}

/**
 * An IF comparison as Word's merge makes it (checked against Word's own merge of test/fixtures/mergeif.docx): numbers
 * as numbers, quoted or not ("100.0" = 100); text equal only in the same case, and a `*` or `?` in the second text a
 * pattern, also in the same case ("a*" isn't true for Ada); text ordered ignoring case (alan < "B").
 */
export function ifCompare(a, op, b) {
  if ((op === '=' || op === '<>') && /[*?]/.test(b)) return wildcard(String(a), b) === (op === '=');
  const x = asNumber(a);
  const y = asNumber(b);
  const c = x != null && y != null ? Math.sign(x - y) : op === '=' || op === '<>' ? (a === b ? 0 : 1) : Math.sign(a.localeCompare(b, 'en'));
  return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[op] ?? false;
}

/** An IF field's text for the record: { IF a op b "true text" "false text" }, with its own \* switches. */
function ifValue(doc, f, ctx) {
  const all = ifWords(doc, f, ctx).slice(1);
  const words = [];
  let sw = '';
  for (let i = 0; i < all.length; i++) {
    const w = all[i];
    if (!w.quoted && /^\\[*#@]$/.test(w.t)) { const arg = all[++i]?.t ?? ''; sw += w.t === '\\*' ? ` \\* ${arg}` : ` ${w.t} "${arg}"`; } else if (!w.quoted && w.t.startsWith('\\')) sw += ` ${w.t}`;
    else words.push(w.t);
  }
  const [a, op, b, yes = '', no = ''] = words;
  if (a == null || op == null || b == null) return '';
  const text = ifCompare(a, op, b) ? yes : no;
  return sw ? formatMergeValue(text, sw) : text;
}

/**
 * NEXT, NEXTIF or SKIPIF at the cursor, as Word's merge takes them (checked against Word's own merge of
 * test/fixtures/mergenext.docx and mergeskip*.docx): NEXT moves on to the next record, NEXTIF when its comparison
 * (IF's) is true, the fields after it filled from that record; SKIPIF, when true, leaves the whole document out.
 */
function step(doc, f, kind, ctx) {
  if (kind === 'NEXT') { ctx.i++; return; }
  if (kind === 'SKIPIF' && ctx.preview) return;
  const [a, op, b] = ifWords(doc, f, ctx).slice(1).map((w) => w.t);
  if (a == null || op == null || b == null || !ifCompare(a, op, b)) return;
  if (kind === 'NEXTIF') ctx.i++; else ctx.skip = true;
}

/**
 * The fields merging a record fills, outermost first: merge fields, IF fields testing or showing one, and NEXT, NEXTIF
 * and SKIPIF (their text nothing), each with its value: [{ field, value }]. `ctx` a mergeCursor, moved on as the
 * fields in order say. (Fields nested in other fields are filled within them.)
 */
export function mergeTargets(doc, ctx) {
  const out = [];
  const walk = (list) => {
    for (const f of list) {
      // (Once a SKIPIF leaves the document out, nothing after it moves the merge on: the next starts after this record.)
      if (ctx.skip) return;
      const kind = stepOf(f);
      if (kind) { step(doc, f, kind, ctx); out.push({ field: f, value: '' }); } else if (parseMergeField(codeOf(f)) || (isIf(f) && usesMerge(f))) out.push({ field: f, value: fieldValue(doc, f, ctx) });
      else { walk(nested(f)); walk(f.kids); }
    }
  };
  walk(fieldTree(doc));
  return out.sort((a, b) => a.field.from - b.field.from);
}

/**
 * The formatting a field's text takes: its result's first character's (\* MERGEFORMAT), else its code's. `inside`, the
 * textblock holding the field and its content's start ({ node, start }): looked through alone, not from the document's
 * start (a preview of many fields, mailmerge.js). With `index`, the child of it the field starts at, looked through
 * from there (a merge of many fields in one paragraph, mergeRecord).
 */
export function fieldMarks(doc, f, inside = null) {
  if (f.simple) {
    const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(f.simple.attrs.xml)?.[0];
    return rPr ? marksFromRPr(rPr).map((m) => schema.markFromJSON(m)) : f.simple.marks;
  }
  const range = f.resultFrom != null && f.resultTo > f.resultFrom ? [f.resultFrom, f.resultTo] : [f.from, f.to];
  let found = null;
  const take = (n) => { if (found == null && (n.isText || n.type === schema.nodes.opaque_inline)) found = n.marks; return found == null; };
  if (inside?.index != null) {
    const { node } = inside;
    for (let i = inside.index, pos = f.from; found == null && i < node.childCount && pos < range[1]; i++) {
      const c = node.child(i);
      if (pos + c.nodeSize > range[0]) take(c);
      pos += c.nodeSize;
    }
  } else if (inside) inside.node.nodesBetween(range[0] - inside.start, range[1] - inside.start, take);
  else doc.nodesBetween(range[0], range[1], take);
  return found || [];
}

// --- Field switches ---------------------------------------------------------------------------------------------------

const quoted = (instr, sw) => { const m = new RegExp(`\\\\${sw}\\s+(?:"([^"]*)"|(\\S+))`, 'i').exec(instr); return m ? (m[1] ?? m[2]) : null; };
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** A date picture (\@ "MMMM d, yyyy") applied to a Date. */
function formatDate(d, pic) {
  const pad = (v, n = 2) => String(v).padStart(n, '0');
  const h12 = d.getHours() % 12 || 12;
  const tokens = {
    yyyy: () => d.getFullYear(), yy: () => pad(d.getFullYear() % 100), MMMM: () => MONTHS[d.getMonth()], MMM: () => MONTHS[d.getMonth()].slice(0, 3),
    MM: () => pad(d.getMonth() + 1), M: () => d.getMonth() + 1, dddd: () => DAYS[d.getDay()], ddd: () => DAYS[d.getDay()].slice(0, 3), dd: () => pad(d.getDate()),
    d: () => d.getDate(), HH: () => pad(d.getHours()), H: () => d.getHours(), hh: () => pad(h12), h: () => h12, mm: () => pad(d.getMinutes()), m: () => d.getMinutes(),
    ss: () => pad(d.getSeconds()), s: () => d.getSeconds(), 'AM/PM': () => (d.getHours() < 12 ? 'AM' : 'PM'), 'am/pm': () => (d.getHours() < 12 ? 'am' : 'pm'),
  };
  return pic.replace(/'([^']*)'|yyyy|yy|MMMM|MMM|MM|M|dddd|ddd|dd|d|HH|H|hh|h|mm|m|ss|s|AM\/PM|am\/pm/g, (t, lit) => (lit != null ? lit : String(tokens[t]())));
}

/** A number picture (\# "$#,##0.00") applied to a number: digits (0 shown, # if any), a thousands comma, decimals. */
function formatNumber(v, pic) {
  if (pic.length > 64) return String(v);
  const m = /^([^#0,.]*)([#0,]*)(?:\.([#0]+))?(.*)$/.exec(pic);
  if (!m) return String(v);
  const [, pre, int, dec = '', post] = m;
  const places = dec.length;
  const fixed = Math.abs(v).toFixed(places);
  let [ip, dp = ''] = fixed.split('.');
  if (ip === '0' && !int.includes('0')) ip = '';
  if (int.includes(',')) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const minInt = (int.replace(/[#,]/g, '').length);
  ip = ip.padStart(minInt, '0');
  dp = dp.replace(/0+$/, (z) => z.slice(0, Math.max(0, dec.replace(/#/g, '').length - (dp.length - z.length))));
  return `${v < 0 ? '-' : ''}${pre}${ip}${dp ? `.${dp}` : ''}${post}`;
}

/** A field's value with its switches (\* case, \# number, \@ date, \b before, \f after), as Word formats it. */
export function formatMergeValue(value, instr) {
  let s = value == null ? '' : value instanceof Date ? value : String(value);
  const date = quoted(instr, '@');
  if (date) {
    // (A date alone, 2026-12-01, is that day where the document is merged: JavaScript would read it as UTC midnight.)
    const ymd = typeof s === 'string' && /^(\d{4})-(\d{2})-(\d{2})$/.exec(s.trim());
    const d = s instanceof Date ? s : ymd ? new Date(+ymd[1], +ymd[2] - 1, +ymd[3]) : new Date(s);
    if (s !== '' && !Number.isNaN(d.getTime())) s = formatDate(d, date);
  }
  if (s instanceof Date) s = s.toLocaleDateString('en-US');
  const num = quoted(instr, '#');
  if (num && s !== '' && !Number.isNaN(Number(String(s).replace(/[$,\s]/g, '')))) s = formatNumber(Number(String(s).replace(/[$,\s]/g, '')), num);
  for (const m of instr.matchAll(/\\\*\s*(\w+)/g)) {
    const k = m[1].toLowerCase();
    if (k === 'upper') s = s.toUpperCase();
    else if (k === 'lower') s = s.toLowerCase();
    else if (k === 'firstcap') s = s.charAt(0).toUpperCase() + s.slice(1);
    else if (k === 'caps') s = s.replace(/(^|[^\p{L}'])(\p{L})/gu, (_, a, b) => a + b.toUpperCase());
  }
  if (s !== '') s = (quoted(instr, 'b') ?? '') + s + (quoted(instr, 'f') ?? '');
  return s;
}

/** A record's value for a field name: its own key, or one that matches ignoring case and spaces (Word's matching). */
export function recordValue(record, name) {
  if (!record) return '';
  if (Object.prototype.hasOwnProperty.call(record, name)) return record[name];
  const norm = (k) => String(k).toLowerCase().replace(/[\s_]/g, '');
  const k = Object.keys(record).find((x) => norm(x) === norm(name));
  return k == null ? '' : record[k];
}

// --- Merging ----------------------------------------------------------------------------------------------------------

/**
 * The document with its merge fields replaced by the cursor's records' values (as plain text in the results'
 * formatting); the cursor is left where the document's NEXT and NEXTIF fields moved it.
 */
export function mergeRecord(doc, ctx) {
  // (IF fields too: Word's merged document has their text, the field gone; NEXT, NEXTIF and SKIPIF leave nothing.)
  const targets = mergeTargets(doc, ctx);
  if (!targets.length) return doc;
  // One pass, the fields in order: each textblock holding fields made again once, its fields' nodes swapped for their
  // text, and the blocks above it copied once (a replace and two resolves a field went through the whole document for
  // each, as many times as it had fields).
  let next = 0;
  const block = (node, start) => {
    const end = start + node.content.size;
    const kids = [];
    let pos = start;
    let skipTo = -1;
    node.forEach((child, _, index) => {
      const at = pos;
      pos += child.nodeSize;
      if (at < skipTo) return;
      while (next < targets.length && targets[next].field.from < at) next++;
      while (next < targets.length && targets[next].field.from === at) {
        const { field: f, value } = targets[next++];
        // (A field spanning paragraphs: left as it is; such fields aren't merge fields in Word.)
        if (f.to > end) continue;
        if (value) kids.push(schema.text(value, fieldMarks(doc, f, { node, start, index })));
        skipTo = f.to;
        return;
      }
      kids.push(child);
    });
    // (Text either side of a field's place joined when its formatting is the same, as a replace joins it.)
    return node.copy(Fragment.fromArray(kids));
  };
  // `node` at content start `start`: the same node when no field starts inside it.
  const walk = (node, start) => {
    if (next >= targets.length || targets[next].field.from >= start + node.content.size) return node;
    if (node.inlineContent) return block(node, start);
    const kids = [];
    let changed = false;
    let pos = start;
    node.forEach((child) => {
      // (A field before this child that wasn't in a textblock: none in this schema, passed over.)
      while (next < targets.length && targets[next].field.from < pos) next++;
      const k = child.isLeaf ? child : walk(child, pos + 1);
      if (k !== child) changed = true;
      kids.push(k);
      pos += child.nodeSize;
    });
    return changed ? node.copy(Fragment.fromArray(kids)) : node;
  };
  return walk(doc, 0);
}

/** w14:paraId and the like dropped (each must be unique; Word gives new ones). */
function withoutIds(node) {
  if (node.isText) return node;
  const attrs = node.attrs.pattrs ? { ...node.attrs, pattrs: null } : node.attrs;
  const kids = [];
  node.forEach((c) => kids.push(withoutIds(c)));
  return node.type.create(attrs, kids, node.marks);
}

/** One document of several records' merged documents, each record after the first on a new page. */
function joinRecords(docs) {
  const first = docs[0];
  const blocks = [];
  docs.forEach((d, i) => {
    const kids = [];
    withoutIds(d).forEach((c) => kids.push(c));
    if (i < docs.length - 1) {
      // The record ends with a next-page section break in the document's own last section's setup, on its last
      // paragraph's mark as Word puts it (an empty paragraph after a table).
      const sect = withSectType(d.attrs.sectPr || '<w:sectPr/>', 'nextPage');
      const last = kids[kids.length - 1];
      const own = last && last.type === schema.nodes.paragraph && !(last.attrs.extra || []).some((e) => e[0] === 'sectPr');
      const para = own ? last : schema.nodes.paragraph.create();
      const marked = para.type.create({ ...para.attrs, extra: [...(para.attrs.extra || []), ['sectPr', sect]] }, para.content, para.marks);
      if (own) kids[kids.length - 1] = marked; else kids.push(marked);
    }
    blocks.push(...kids);
  });
  // (The last record's own last section: its headers and footers are that record's.)
  return first.type.create(docs[docs.length - 1].attrs, blocks);
}

// --- Headers and footers ----------------------------------------------------------------------------------------------

const REF = /<w:(header|footer)Reference\b[^>]*>(?:<\/w:\1Reference>)?/g;

/** The document with each section's w:sectPr XML put through `fn` (sections in order). */
function mapSections(doc, fn) {
  const kids = [];
  doc.forEach((n) => {
    const s = n.type === schema.nodes.paragraph && (n.attrs.extra || []).find((e) => e[0] === 'sectPr');
    kids.push(s ? n.type.create({ ...n.attrs, extra: n.attrs.extra.map((e) => (e === s ? ['sectPr', fn(e[1])] : e)) }, n.content, n.marks) : n);
  });
  return doc.type.create({ ...doc.attrs, sectPr: fn(doc.attrs.sectPr || '<w:sectPr/>') }, kids);
}

/**
 * Copies of header and footer parts (their own relationships too, for their pictures): `copy(rId)` makes one and
 * returns its rId, the names and rIds those pkg.addOverride and pkg.addRel give, and `done()` writes the content types
 * and the document's relationships once, after the last (adding each copy to them read and wrote them whole, for
 * every record: time that grew with the square of the records). Nothing reads the new relationships before `done`.
 */
function storyCopier(session) {
  const pkg = session.pkg;
  const relsOf = (p) => `${p.slice(0, p.lastIndexOf('/') + 1)}_rels/${p.slice(p.lastIndexOf('/') + 1)}.rels`;
  const relsPath = relsOf(pkg.mainPath);
  const relDoc = pkg.xml(relsPath) || parseXml(`<Relationships xmlns="${NS.rel}"/>`);
  const used = new Set(kids(relDoc.documentElement, 'Relationship', NS.rel).map((r) => r.getAttribute('Id')));
  const ctDoc = pkg.xml('[Content_Types].xml');
  const parts = new Set(kids(ctDoc.documentElement, 'Override', NS.ct).map((o) => o.getAttribute('PartName')));
  const prefix = pkg.idPrefix || 'rId';
  // (The first free number of each part name and, with an editor's own prefix, of the rIds: they only go up.)
  const free = { header: 1, footer: 1, id: 1 };
  let rels = false;
  let types = false;
  return {
    copy(rId) {
      const rel = pkg.relById(rId);
      const kind = /footer$/.test(rel.type) ? 'footer' : 'header';
      let n = free[kind];
      while (pkg.files[`${pkg.mainDir}${kind}${n}.xml`]) n++;
      free[kind] = n + 1;
      const path = `${pkg.mainDir}${kind}${n}.xml`;
      pkg.files[path] = pkg.files[rel.target];
      if (pkg.files[relsOf(rel.target)]) pkg.files[relsOf(path)] = pkg.files[relsOf(rel.target)];
      if (!parts.has(`/${path}`)) {
        const o = ctDoc.createElementNS(NS.ct, 'Override');
        o.setAttribute('PartName', `/${path}`);
        o.setAttribute('ContentType', `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`);
        ctDoc.documentElement.appendChild(o);
        parts.add(`/${path}`);
        types = true;
      }
      let k = pkg.idPrefix ? free.id : used.size + 1;
      while (used.has(prefix + k)) k++;
      if (pkg.idPrefix) free.id = k;
      const id = prefix + k;
      used.add(id);
      const r = relDoc.createElementNS(NS.rel, 'Relationship');
      r.setAttribute('Id', id);
      r.setAttribute('Type', rel.type);
      r.setAttribute('Target', path.slice(pkg.mainDir.length));
      relDoc.documentElement.appendChild(r);
      rels = true;
      return id;
    },
    done() {
      if (types) pkg.setText('[Content_Types].xml', serializeXml(ctDoc));
      if (rels) { pkg.setText(relsPath, serializeXml(relDoc)); pkg.rels = pkg.readRels(pkg.mainPath); }
    },
  };
}

/**
 * The headers and footers with merge fields, each merged document's own: rId -> [merged story per document], empty
 * when none has any. Each is filled from the record its document ended on (after its NEXT fields), as in Word.
 */
function mergedStories(session, records, units) {
  const stories = new Map();
  for (const id of storyIds(session)) {
    const doc = session.story(id);
    if (doc && mergeTargets(doc, mergeCursor(records)).length) stories.set(id, units.map((u) => mergeRecord(doc, mergeCursor(records, u.last, { seq: u.seq }))));
  }
  return stories;
}

/**
 * The merged documents, as Word's merge makes them: each starting at the record after the one the document before
 * ended on (NEXT and NEXTIF moving on within one), those a SKIPIF left out dropped: [{ doc, first, last, seq }].
 */
export function mergeUnits(doc, records) {
  const units = [];
  for (let i = 0; i < records.length;) {
    const ctx = mergeCursor(records, i, { seq: units.length + 1 });
    const merged = mergeRecord(doc, ctx);
    if (!ctx.skip) units.push({ doc: merged, first: i, last: ctx.i, seq: ctx.seq });
    // (ctx.i is never before i: the loop ends.)
    i = ctx.i + 1;
  }
  return units;
}

/**
 * Record `i`'s document with its own headers and footers, as Word's merged document has them: record 0 keeps the
 * parts (filled with its values), each later record a copy of each, every section naming its headers and footers
 * (a section that inherited them would get the record before's).
 */
function withRecordStories(session, doc, stories, i, copier) {
  if (!stories.size) return doc;
  const copies = new Map();
  const idFor = (id) => {
    if (!stories.has(id)) return id;
    if (!copies.has(id)) copies.set(id, i === 0 ? id : copier.copy(id));
    return copies.get(id);
  };
  let refs = {};
  const out = mapSections(doc, (xml) => {
    refs = { ...refs };
    for (const m of xml.matchAll(REF)) {
      const type = /w:type="(\w+)"/.exec(m[0])?.[1] || 'default';
      const id = /r:id="([^"]+)"/.exec(m[0])?.[1];
      if (id) refs[`${m[1]}.${type}`] = id;
    }
    const own = Object.entries(refs).map(([k, id]) => { const [kind, type] = k.split('.'); return `<w:${kind}Reference w:type="${type}" r:id="${idFor(id)}"/>`; }).join('');
    const rest = xml.replace(REF, '');
    return /^<w:sectPr\b[^>]*\/>$/.test(rest) ? rest.replace(/\/>$/, `>${own}</w:sectPr>`) : rest.replace(/^(<w:sectPr\b[^>]*>)/, `$1${own}`);
  });
  // (A copy's relationship is written after the last record, storyCopier: its merged story is set now, saved after.)
  for (const [id, copy] of copies) {
    session.story(copy);
    session.setStory(copy, stories.get(id)[i]);
  }
  return out;
}

/**
 * Mail merge: `records` (objects, field name -> value) into the .docx's merge fields. With `combine` (the default) one
 * .docx of all of them, each record starting on a new page, as Word's Edit Individual Documents; without, one .docx per
 * record. (A document with NEXT or NEXTIF fields takes several records each; one a SKIPIF leaves out is dropped.)
 * Returns { bytes } or { files: [bytes...] }, and `fields` (the document's merge field names) and `missing` (those no
 * record has a value for).
 */
export function mailMerge(bytes, records, { combine = true } = {}) {
  if (!Array.isArray(records) || !records.length) throw new Error('mailMerge: no records');
  const session = new DocxSession(bytes);
  const fields = mergeFieldNames(session);
  const missing = fields.filter((name) => !records.some((r) => recordValue(r, name) !== ''));
  const units = mergeUnits(session.doc, records);
  // (Every record left out by SKIPIF: one empty document, its headers' and footers' merge fields empty.)
  if (!units.length) units.push({ doc: session.doc.type.create(session.doc.attrs, schema.nodes.paragraph.create()), first: records.length, last: records.length, seq: 1 });
  const docs = units.map((u) => u.doc);
  // Merge fields in headers and footers are filled for each document too, as in Word.
  const stories = mergedStories(session, records, units);
  if (combine) {
    const copier = storyCopier(session);
    const joined = joinRecords(docs.map((d, i) => withRecordStories(session, d, stories, i, copier)));
    copier.done();
    return { bytes: session.save(joined), fields, missing };
  }
  return {
    files: docs.map((d, i) => {
      const s = new DocxSession(bytes);
      for (const [id, merged] of stories) { s.story(id); s.setStory(id, merged[i]); }
      return s.save(d);
    }),
    fields,
    missing,
  };
}
