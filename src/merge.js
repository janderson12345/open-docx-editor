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

/** A field's value for the record: a merge field's, an IF's text, any other field's shown result. */
function fieldValue(doc, f, record, index) {
  const code = codeOf(f);
  const m = parseMergeField(code);
  if (m) return m.kind === 'MERGEFIELD' ? formatMergeValue(recordValue(record, m.name), code) : String(index + 1);
  if (isIf(f)) return ifValue(doc, f, record, index);
  return f.resultFrom != null && f.resultTo > f.resultFrom ? doc.textBetween(f.resultFrom, f.resultTo, '', '') : '';
}

/**
 * An IF field's code in words, as Word reads it: split at spaces outside quotes, \" a quote in quotes, a nested
 * field's value one word however many spaces it has ({ MERGEFIELD City } = "New York" is true for New York in Word).
 */
function ifWords(doc, f, record, index) {
  const words = [];
  let cur = null;
  let q = false;
  const end = () => { if (cur) words.push(cur); cur = null; };
  for (const p of f.parts) {
    if (typeof p !== 'string') { (cur ||= { t: '', quoted: q }).t += fieldValue(doc, p, record, index); continue; }
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
 * An IF comparison as Word's merge makes it (checked against Word's own merge of test/fixtures/mergeif.docx): numbers
 * as numbers, quoted or not ("100.0" = 100); text equal only in the same case, and a `*` or `?` in the second text a
 * pattern, also in the same case ("a*" isn't true for Ada); text ordered ignoring case (alan < "B").
 */
export function ifCompare(a, op, b) {
  if ((op === '=' || op === '<>') && /[*?]/.test(b)) {
    const pattern = b.replace(/[.+^$()|[\]{}\\]/g, '\\$&').replace(/\*/g, '[\\s\\S]*').replace(/\?/g, '[\\s\\S]');
    const re = new RegExp('^' + pattern + '$');
    return re.test(a) === (op === '=');
  }
  const x = asNumber(a);
  const y = asNumber(b);
  const c = x != null && y != null ? Math.sign(x - y) : op === '=' || op === '<>' ? (a === b ? 0 : 1) : Math.sign(a.localeCompare(b, 'en'));
  return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[op] ?? false;
}

/** An IF field's text for the record: { IF a op b "true text" "false text" }, with its own \* switches. */
function ifValue(doc, f, record, index) {
  const all = ifWords(doc, f, record, index).slice(1);
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
 * The fields merging a record fills, outermost first: merge fields, and IF fields testing or showing one, each with its
 * value for the record: [{ field, value }]. (Fields nested in other fields are filled within them.)
 */
export function mergeTargets(doc, record, index = 0) {
  const out = [];
  const walk = (list) => {
    for (const f of list) {
      if (parseMergeField(codeOf(f)) || (isIf(f) && usesMerge(f))) out.push({ field: f, value: fieldValue(doc, f, record, index) });
      else { walk(nested(f)); walk(f.kids); }
    }
  };
  walk(fieldTree(doc));
  return out.sort((a, b) => a.field.from - b.field.from);
}

/** The formatting a field's text takes: its result's first character's (\* MERGEFORMAT), else its code's. */
export function fieldMarks(doc, f) {
  if (f.simple) {
    const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(f.simple.attrs.xml)?.[0];
    return rPr ? marksFromRPr(rPr).map((m) => schema.markFromJSON(m)) : f.simple.marks;
  }
  const range = f.resultFrom != null && f.resultTo > f.resultFrom ? [f.resultFrom, f.resultTo] : [f.from, f.to];
  let found = null;
  doc.nodesBetween(range[0], range[1], (n) => { if (found == null && (n.isText || n.type === schema.nodes.opaque_inline)) found = n.marks; return found == null; });
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

/** The document with its merge fields replaced by `record`'s values (as plain text in the results' formatting). */
export function mergeRecord(doc, record, index = 0) {
  let out = doc;
  // (IF fields too: Word's merged document has their text, the field gone.)
  for (const { field: f, value } of mergeTargets(doc, record, index).reverse()) {
    const marks = fieldMarks(out, f);
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

/** A copy of a header or footer part (its own relationships too, for its pictures); returns the copy's rId. */
function copyStory(session, rId) {
  const pkg = session.pkg;
  const rel = pkg.relById(rId);
  const kind = /footer$/.test(rel.type) ? 'footer' : 'header';
  const relsOf = (p) => `${p.slice(0, p.lastIndexOf('/') + 1)}_rels/${p.slice(p.lastIndexOf('/') + 1)}.rels`;
  let n = 1;
  while (pkg.files[`${pkg.mainDir}${kind}${n}.xml`]) n++;
  const path = `${pkg.mainDir}${kind}${n}.xml`;
  pkg.files[path] = pkg.files[rel.target];
  if (pkg.files[relsOf(rel.target)]) pkg.files[relsOf(path)] = pkg.files[relsOf(rel.target)];
  pkg.addOverride(path, `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`);
  return pkg.addRel(rel.type, path.slice(pkg.mainDir.length));
}

/**
 * The headers and footers with merge fields, each record's own: { stories: rId -> [merged story per record] }, empty
 * when none has any.
 */
function mergedStories(session, records) {
  const stories = new Map();
  for (const id of storyIds(session)) {
    const doc = session.story(id);
    if (doc && mergeTargets(doc, records[0]).length) stories.set(id, records.map((r, i) => mergeRecord(doc, r, i)));
  }
  return stories;
}

/**
 * Record `i`'s document with its own headers and footers, as Word's merged document has them: record 0 keeps the
 * parts (filled with its values), each later record a copy of each, every section naming its headers and footers
 * (a section that inherited them would get the record before's).
 */
function withRecordStories(session, doc, stories, i) {
  if (!stories.size) return doc;
  const copies = new Map();
  const idFor = (id) => {
    if (!stories.has(id)) return id;
    if (!copies.has(id)) copies.set(id, i === 0 ? id : copyStory(session, id));
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
  for (const [id, copy] of copies) {
    session.story(copy);
    session.setStory(copy, stories.get(id)[i]);
  }
  return out;
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
  // Merge fields in headers and footers are filled for each record too, as in Word.
  const stories = mergedStories(session, records);
  if (combine) return { bytes: session.save(joinRecords(docs.map((d, i) => withRecordStories(session, d, stories, i)))), fields, missing };
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
