// Mail merge, as Word's Mailings > Finish & Merge > Edit Individual Documents: a document's merge fields (MERGEFIELD,
// with MERGEREC and MERGESEQ) filled from each record of a data source, giving one document per record or all of them
// in one, each record starting on a new page (a next-page section break, as Word joins them). The fields become plain
// text in the result, in their results' formatting, as in Word's merged document. Works without an editor (in Node,
// call setXmlImpl(...) with a DOMParser/XMLSerializer first) and in the browser.
import { Fragment, Slice } from 'prosemirror-model';
import { DocxSession } from './ooxml/docx.js';
import { schema } from './editor/schema.js';
import { marksFromRPr } from './ooxml/reader.js';
import { withSectType } from './editor/commands.js';

const fldType = (n) => (n.type === schema.nodes.opaque_inline && n.attrs.kind === 'fldChar' ? (/fldCharType="(\w+)"/.exec(n.attrs.xml) || [])[1] : null);
const unescapeXml = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
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

/** The names of a .docx's merge fields (MERGEFIELD), each once, in order. */
export function mergeFieldNames(bytes) {
  const session = bytes instanceof DocxSession ? bytes : new DocxSession(bytes);
  return [...new Set(mergeFieldsIn(session.doc).filter((f) => f.kind === 'MERGEFIELD').map((f) => f.name))];
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

/** The document with its merge fields replaced by `record`'s values (as plain text in the results' formatting). */
export function mergeRecord(doc, record, index = 0) {
  let out = doc;
  for (const f of mergeFieldsIn(doc).reverse()) {
    const value = f.kind === 'MERGEREC' || f.kind === 'MERGESEQ' ? String(index + 1) : formatMergeValue(recordValue(record, f.name), f.instr);
    let marks = [];
    if (f.simple) {
      const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(f.simple.attrs.xml)?.[0];
      marks = rPr ? marksFromRPr(rPr).map((m) => schema.markFromJSON(m)) : f.simple.marks;
    } else {
      // The result's first character's formatting (\* MERGEFORMAT), else the field code's.
      const range = f.resultFrom != null && f.resultTo > f.resultFrom ? [f.resultFrom, f.resultTo] : [f.from, f.to];
      let found = null;
      out.nodesBetween(range[0], range[1], (n) => { if (found == null && (n.isText || n.type === schema.nodes.opaque_inline)) found = n.marks; return found == null; });
      marks = found || [];
    }
    const $from = out.resolve(f.from);
    const $to = out.resolve(f.to);
    // (A field spanning paragraphs: only its first paragraph's part is replaced; such fields aren't merge fields in Word.)
    if ($from.parent !== $to.parent) continue;
    const text = value ? Fragment.from(schema.text(value, marks)) : Fragment.empty;
    out = out.replace(f.from, f.to, new Slice(text, 0, 0));
  }
  return out;
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
  return first.type.create(first.attrs, blocks);
}

/**
 * Mail merge: `records` (objects, field name -> value) into the .docx's merge fields. With `combine` (the default) one
 * .docx of all of them, each record starting on a new page, as Word's Edit Individual Documents; without, one .docx per
 * record. Returns { bytes } or { files: [bytes...] }, and `fields` (the document's merge field names) and `missing`
 * (those no record has a value for).
 */
export function mailMerge(bytes, records, { combine = true } = {}) {
  if (!Array.isArray(records) || !records.length) throw new Error('mailMerge: no records');
  const session = new DocxSession(bytes);
  const fields = mergeFieldNames(session);
  const missing = fields.filter((name) => !records.some((r) => recordValue(r, name) !== ''));
  const docs = records.map((r, i) => mergeRecord(session.doc, r, i));
  if (combine) return { bytes: session.save(joinRecords(docs)), fields, missing };
  return { files: docs.map((d) => new DocxSession(bytes).save(d)), fields, missing };
}
