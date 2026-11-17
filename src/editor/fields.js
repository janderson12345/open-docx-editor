// Form fields: Word content controls (w:sdt), inline (an `sdt` mark over its
// text) or block-level (paragraphs carrying attrs.sdt). Read and write their
// values by tag or title, and a forms mode where only the fields can be edited.
import { announce as say, liveRegion } from './live.js';
import { Plugin, PluginKey, TextSelection } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { Fragment, Slice } from 'prosemirror-model';
import { schema } from './schema.js';
import { fieldInfo } from '../ooxml/reader.js';
import { legacyInfo } from '../ooxml/ffdata.js';
import { unescapeXml, startTag, startTags, tagAttr, setTagAttr } from '../ooxml/xml.js';
import { openSectionAt } from './protect.js';
import { formatDate, prAttr, showsPlc, dropPlc } from '../ooxml/databind.js';


/** Every field in document order: [{ key, tag, alias, kind, from, to, block, value, checked?, items? }] */
export function listFields(doc) {
  const byKey = new Map();
  const order = [];
  const add = (key, info, from, to, text, extra = {}) => {
    let f = byKey.get(key);
    if (!f) {
      f = { key, tag: unescapeXml(info.tag), alias: unescapeXml(info.alias), kind: info.kind, from, to, text: '', ...extra };
      byKey.set(key, f);
      order.push(f);
    }
    f.from = Math.min(f.from, from);
    f.to = Math.max(f.to, to);
    f.text += text;
  };
  doc.descendants((node, pos) => {
    if (node.isBlock && node.attrs.sdt && node.type === schema.nodes.paragraph) {
      const open = node.attrs.sdt.open || '';
      const info = fieldInfo(open);
      const key = `b${node.attrs.sdt.g}`;
      const had = byKey.has(key);
      add(key, info, pos, pos + node.nodeSize, (had ? '\n' : '') + node.textContent, { block: true, pr: open });
      return false;
    }
    if (!node.isInline) return true;
    const m = node.marks.find((x) => x.type === schema.marks.sdt);
    if (m) add(`i${m.attrs.key}`, m.attrs, pos, pos + node.nodeSize, node.textContent, { block: false, mark: m, pr: m.attrs.pr });
    return false;
  });
  order.push(...legacyFields(doc));
  order.sort((a, b) => a.from - b.from);
  for (const f of order) {
    if (f.legacy) continue;
    const ph = f.mark?.attrs.phText;
    f.value = (ph != null && f.text === ph) || (f.block && PLC.test(f.pr)) ? '' : f.text;
    // (Tags found by scanning: a co-editor's w:sdtPr of unclosed tags must not send a pattern back over them.)
    if (f.kind === 'checkbox') f.value = f.checked = /^(1|true)$/.test(prAttr(f.pr, 'w14:checked', 'w14:val') ?? '');
    if (f.kind === 'dropdown' || f.kind === 'combobox') {
      f.items = [];
      for (const t of startTags(f.pr || '', 'w:listItem')) {
        const v = tagAttr(t.tag, 'w:value') ?? undefined;
        f.items.push({ text: unescapeXml(tagAttr(t.tag, 'w:displayText') ?? v), value: unescapeXml(v) });
      }
    }
    // A list's value is the chosen item's value (what Word stores), not its display text.
    if (f.items && f.value) f.value = f.items.find((i) => i.text === f.value)?.value ?? f.value;
  }
  return order;
}

// Legacy form fields (Word 97-2003: FORMTEXT, FORMCHECKBOX, FORMDROPDOWN): a complex field whose
// begin fldChar holds w:ffData. A text field's value is its result text; a check box and a drop-down
// keep their state in ffData and show nothing between separate and end. A text field's result may run
// over several paragraphs (Word writes each line of a multi-line value as a paragraph): its value has a
// line break for each paragraph mark.
const fldOf = (n) => (n.type === schema.nodes.opaque_inline && n.attrs.kind === 'fldChar' ? (/fldCharType="(\w+)"/.exec(n.attrs.xml) || [])[1] : null);
function legacyFields(doc) {
  const out = [];
  let open = null; // (left open at a paragraph's end when the field goes on into the next)
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    node.forEach((c, off) => {
      const at = pos + 1 + off;
      const t = fldOf(c);
      if (t === 'begin') open = { info: legacyInfo(c.attrs.xml), begin: at, sep: null };
      else if (t === 'separate' && open) open.sep = at + c.nodeSize;
      else if (t === 'end' && open) {
        const { info, begin, sep } = open;
        open = null;
        if (!info) return;
        const from = info.kind === 'text' && sep != null ? sep : begin;
        const to = info.kind === 'text' && sep != null ? at : begin + 1;
        const text = info.kind === 'text' ? doc.textBetween(from, to, '\n') : info.kind === 'dropdown' ? info.text : '';
        const f = { key: `l${out.length}`, tag: info.name || undefined, alias: undefined, kind: info.kind, from, to, text, legacy: true, begin, block: false, pr: '' };
        if (info.kind === 'checkbox') f.value = f.checked = info.checked;
        else if (info.kind === 'dropdown') { f.items = info.items.map((x) => ({ text: x, value: x })); f.value = info.text; }
        else { f.value = text.trim() ? text : ''; if (info.maxLength) f.maxLength = info.maxLength; }
        out.push(f);
      }
    });
    return false;
  });
  return out;
}

function writeLegacy(tr, f, value) {
  const beginNode = tr.doc.nodeAt(f.begin);
  let xml = beginNode.attrs.xml;
  if (f.kind === 'checkbox') {
    const on = value === true || value === 'true' || value === 1 || value === '1';
    const tag = `<w:checked w:val="${on ? 1 : 0}"/>`;
    const c = startTag(xml, 'w:checked');
    if (c?.selfClosing) xml = xml.slice(0, c.at) + tag + xml.slice(c.end);
    else {
      const box = xml.indexOf('<w:checkBox>');
      const d = box >= 0 ? startTag(xml, 'w:default', box) : null;
      if (d?.selfClosing) xml = xml.slice(0, d.end) + tag + xml.slice(d.end);
    }
    tr.setNodeMarkup(f.begin, null, { ...beginNode.attrs, xml });
    return;
  }
  if (f.kind === 'dropdown') {
    const i = Math.max(0, f.items.findIndex((x) => x.value === String(value)));
    const tag = `<w:result w:val="${i}"/>`;
    const r = startTag(xml, 'w:result');
    xml = r?.selfClosing ? xml.slice(0, r.at) + tag + xml.slice(r.end) : xml.replace('<w:ddList>', `<w:ddList>${tag}`);
    tr.setNodeMarkup(f.begin, null, { ...beginNode.attrs, xml });
    return;
  }
  let text = value == null || value === '' ? '\u2002\u2002\u2002\u2002\u2002' : String(value).replace(/\r\n?/g, '\n'); // Word's empty text field: five en spaces
  // A value longer than the field's maximum length is cut to it, as Word does (a line break counts as one character).
  if (f.maxLength && value != null && value !== '') text = [...text].slice(0, f.maxLength).join('');
  replaceLines(tr, f.from, f.to, text, legacyMarks(tr.doc, f));
}

/** The run formatting of a legacy text field's result: its (last) text's, else its begin's. */
function legacyMarks(doc, f) {
  let marks = doc.nodeAt(f.begin).marks;
  doc.nodesBetween(f.from, f.to, (n) => { if (n.isText) marks = n.marks; });
  return marks;
}

/** Replace from..to (within a legacy text field) with `text`, its lines as paragraphs. */
function replaceLines(tr, from, to, text, marks) {
  const lines = text.split('\n');
  const $from = tr.doc.resolve(from);
  const $to = tr.doc.resolve(to);
  if (lines.length === 1 && $from.sameParent($to)) {
    if (!text) tr.delete(from, to);
    else if (from === to) tr.insert(from, schema.text(text, marks));
    else tr.replaceWith(from, to, schema.text(text, marks));
    return;
  }
  // Each line a paragraph, as Word writes them: the first goes on in the field's own paragraph, the last ends in the
  // paragraph the field ends in, and the ones between are split from the first (as Enter would split it).
  const split = { ...$from.parent.attrs, deco: null };
  const paras = lines.map((l, i) => {
    const end = i === lines.length - 1 && !$from.sameParent($to);
    return (end ? $to : $from).parent.type.create(end ? $to.parent.attrs : split, l ? schema.text(l, marks) : null);
  });
  tr.replace(from, to, new Slice(Fragment.from(paras), 1, 1));
}

/** Field values as { name: value } (name = tag, else title); repeated names give the first. */
export function fieldValues(doc) {
  const out = {};
  for (const f of listFields(doc)) {
    const name = f.tag || f.alias;
    if (name && !(name in out)) out[name] = f.value;
  }
  return out;
}

const PLC = { test: showsPlc };
const noPlc = (sdt) => ({ ...sdt, open: dropPlc(sdt.open) });
/** `xml` with its first start tag `name` made `f(tag)` (as it was without one). */
const retag = (xml, name, f) => {
  const t = startTag(xml, name);
  return t ? xml.slice(0, t.at) + f(t) + xml.slice(t.end) : xml;
};

const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
const PH_DEFAULT = 'Click or tap here to enter text.';

/** Replace one field's value inside `tr` (positions from tr.doc). */
export function writeField(tr, f, value) {
  if (f.legacy) return writeLegacy(tr, f, value);
  if (f.block) {
    const first = tr.doc.nodeAt(f.from);
    const lines = String(value ?? '').split('\n');
    const runMarks = (first.firstChild?.marks || []).filter((x) => !(x.type === schema.marks.cstyle && x.attrs.val === 'PlaceholderText'));
    const attrs = { ...first.attrs, sdt: noPlc(first.attrs.sdt) };
    const paras = lines.map((l) => schema.nodes.paragraph.create(attrs, l ? schema.text(l, runMarks) : null));
    tr.replaceWith(f.from, f.to, paras);
    return;
  }
  const m = f.mark;
  let pr = m.attrs.pr;
  let text;
  let phText = null;
  if (f.kind === 'checkbox') {
    const on = value === true || value === 'true' || value === 1 || value === '1';
    pr = retag(pr, 'w14:checked', (t) => setTagAttr(t.tag, 'w14:val', on ? '1' : '0'));
    const code = (name) => { const v = prAttr(pr, name, 'w14:val'); return v && /^[0-9A-Fa-f]{1,6}$/.test(v) && parseInt(v, 16) <= 0x10ffff ? v : null; };
    text = String.fromCodePoint(parseInt(on ? code('w14:checkedState') || '2612' : code('w14:uncheckedState') || '2610', 16));
  } else if (value == null || value === '') {
    // An empty control shows its placeholder, as in Word.
    text = m.attrs.phText || PH_DEFAULT;
    phText = text;
    if (!/<w:showingPlcHdr/.test(pr)) pr = pr.replace(/<\/w:sdtPr>$/, '<w:showingPlcHdr/></w:sdtPr>');
  } else {
    text = String(value);
    if (f.kind === 'dropdown' || f.kind === 'combobox') text = f.items?.find((i) => i.value === text)?.text ?? text;
    pr = dropPlc(pr);
    if (f.kind === 'date') {
      // Word's date picker reads w:fullDate; keep it in step with the text.
      const d = new Date(value instanceof Date ? value : text);
      const full = isNaN(d) ? null : `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}T00:00:00Z`;
      pr = retag(pr, 'w:date', (t) => `<w:date${full ? ` w:fullDate="${full}"` : ''}${t.selfClosing ? '/>' : '>'}`);
      if (value instanceof Date) text = formatDate(d, pr);
    }
  }
  const mark = schema.marks.sdt.create({ ...m.attrs, pr, phText });
  const firstText = (() => { let t = null; tr.doc.nodesBetween(f.from, f.to, (n) => { if (!t && n.isText) t = n; }); return t; })();
  const isPh = (x) => x.type === schema.marks.cstyle && x.attrs.val === 'PlaceholderText';
  let marks = (firstText ? firstText.marks : [m]).filter((x) => !(x.type === schema.marks.sdt && x.attrs.key === m.attrs.key) && !isPh(x));
  if (phText != null) marks = schema.marks.cstyle.create({ val: 'PlaceholderText' }).addToSet(marks);
  marks = mark.addToSet(marks);
  tr.replaceWith(f.from, f.to, schema.text(text, marks));
}

/** Set fields by tag or title: { Name: 'value', Agree: true }. Returns the names that matched nothing. */
export function setFieldValues(view, values) {
  const tr = view.state.tr.setMeta('odx-untracked', true).setMeta(formsKey, { allow: true });
  const missing = new Set(Object.keys(values));
  // Last field first, so earlier positions stay valid.
  const fields = listFields(tr.doc).reverse();
  for (const f of fields) {
    const name = [f.tag, f.alias].find((n) => n != null && n in values);
    if (name == null) continue;
    missing.delete(name);
    writeField(tr, f, values[name]);
  }
  if (tr.docChanged) view.dispatch(tr);
  return [...missing];
}

// Screen readers read a field's text but not its name or kind, so forms mode announces them (4.1.2, 4.1.3),
// e.g. "Customer name, text field, empty" or "Agree, check box, checked".
const KIND = { text: 'text field', richText: 'text field', checkbox: 'check box', dropdown: 'drop-down list', combobox: 'combo box', date: 'date picker' };
export function describeField(f) {
  const name = f.alias || f.tag || 'Field';
  const kind = KIND[f.kind] || 'field';
  const state = f.kind === 'checkbox' ? (f.checked ? 'checked' : 'not checked')
    : f.value === '' ? 'empty' : (f.items?.find((i) => i.value === f.value)?.text ?? f.text);
  return `${name}, ${kind}, ${state}`;
}
const announce = (view, msg) => say(msg);

// --- Forms mode: only the fields can be changed --------------------------------------------

export const formsKey = new PluginKey('odx-forms');

const inField = (fields, from, to) => fields.some((f) => from >= f.from && to <= f.to);

// A legacy text field's length as Word counts it: its characters, a paragraph mark one, its blank (en spaces) none.
const ffLength = (s) => [...s.replace(/\u2002/g, '')].length;
/** Whether an edit takes a legacy text field past its maximum length (typing there is refused, as in Word). */
function overLength(before, after) {
  const now = legacyFields(after);
  if (!now.some((f) => f.maxLength)) return false;
  const was = legacyFields(before);
  if (was.length !== now.length) return false;
  return now.some((f, i) => f.maxLength && ffLength(f.text) > f.maxLength && ffLength(f.text) > ffLength(was[i].text));
}

/**
 * A paste into a legacy text field that would take it past its maximum length: its text, cut to the room left, goes in
 * instead (true), in the field's formatting. A paste that fits is left to the editor (false).
 */
function pasteToFit(view, slice) {
  const { state } = view;
  const { from, to } = state.selection;
  const f = legacyFields(state.doc).find((x) => x.kind === 'text' && x.maxLength && from >= x.from && to <= x.to);
  if (!f) return false;
  const text = slice.content.textBetween(0, slice.content.size, '\n', (leaf) => (leaf.type === schema.nodes.hard_break ? '\n' : '')).replace(/\r\n?/g, '\n');
  const room = f.maxLength - (ffLength(f.text) - ffLength(state.doc.textBetween(from, to, '\n')));
  if (ffLength(text) <= room) return false;
  // (Word's blank, the en spaces, goes when the field is pasted into.)
  const cut = [...text.replace(/\u2002/g, '')].slice(0, Math.max(0, room)).join('');
  const tr = state.tr;
  const blank = !f.value;
  const end = blank ? f.to : to;
  replaceLines(tr, blank ? f.from : from, end, cut, legacyMarks(state.doc, f));
  // (The caret after what was pasted, as a paste leaves it.)
  if (tr.docChanged) view.dispatch(tr.setSelection(TextSelection.create(tr.doc, tr.mapping.map(end))).setMeta('uiEvent', 'paste').scrollIntoView());
  return true;
}

const shadeKey = new PluginKey('odx-ff-shade');
function shade(doc) {
  const decos = legacyFields(doc).filter((f) => f.kind === 'text' && f.to > f.from).map((f) => Decoration.inline(f.from, f.to, { class: 'odx-ff' }));
  return decos.length ? DecorationSet.create(doc, decos) : null;
}

export function formsPlugin() {
  return [new Plugin({
    key: shadeKey,
    state: { init: (_, state) => shade(state.doc), apply: (tr, set) => (tr.docChanged ? shade(tr.doc) : set) },
  }), new Plugin({
    key: formsKey,
    // Screen readers only watch live regions that are already on the page, so it's made up front.
    view: () => { liveRegion(); return {}; },
    state: {
      init: () => false,
      apply: (tr, on) => { const m = tr.getMeta(formsKey); return m && 'on' in m ? m.on : on; },
    },
    filterTransaction(tr, state) {
      if (!formsKey.getState(state) || !tr.docChanged || tr.getMeta(formsKey)?.allow || tr.getMeta('history$') || tr.getMeta('odx-fields')) return true;
      for (let i = 0; i < tr.steps.length; i++) {
        const step = tr.steps[i];
        const doc = tr.docs[i];
        const fields = listFields(doc);
        const from = step.from ?? step.pos;
        const to = step.to ?? step.pos;
        // (A section the form leaves unprotected is edited freely.)
        if (from != null && openSectionAt(doc, from) && openSectionAt(doc, to)) continue;
        if (from == null || !inField(fields, from, to)) return false;
        // Typing must not swallow the whole control: keep at least the field itself.
        const f = fields.find((x) => from >= x.from && to <= x.to);
        if (f && f.kind === 'checkbox') return false;
      }
      return !overLength(state.doc, tr.doc);
    },
    // Text typed at the very start of an inline field joins the field.
    appendTransaction(trs, oldState, state) {
      if (!formsKey.getState(state)) return null;
      // The caret stays in the fields, as in a protected Word form: put outside one (Ctrl+Home, a click on the form's
      // own text), it goes to the next field, else the first (typing there would otherwise be refused, the browser's
      // caret and ours apart).
      if (trs.some((t) => t.selectionSet) && !trs.some((t) => t.docChanged)) {
        const fields = listFields(state.doc);
        const { from, to } = state.selection;
        if (fields.length && !fields.some((x) => from >= x.from && to <= x.to) && !openSectionAt(state.doc, from)) {
          const to = fields.find((x) => x.from >= from) || fields[0];
          say(describeField(to)); // (as Tab does: NVDA heard only the placeholder after Ctrl+Home)
          return state.tr.setSelection(fieldSelection(state.doc, to));
        }
        return null;
      }
      if (!trs.some((t) => t.docChanged)) return null;
      const tr = state.tr;
      for (const t of trs) {
        t.steps.forEach((step, i) => {
          if (step.slice?.size) {
            const from = t.mapping.slice(i).map(step.from);
            const to = from + step.slice.size;
            const $after = state.doc.resolve(Math.min(to, state.doc.content.size));
            const near = $after.nodeAfter?.marks.find((m) => m.type === schema.marks.sdt) || state.doc.resolve(from).nodeBefore?.marks.find((m) => m.type === schema.marks.sdt);
            if (near) tr.addMark(from, to, near);
          }
        });
      }
      // A field whose text is no longer its placeholder holds a real value: drop the grey placeholder look.
      const ph = schema.marks.cstyle;
      for (const f of listFields(tr.doc)) {
        if (f.block || f.legacy || f.mark.attrs.phText == null || f.text === f.mark.attrs.phText) continue;
        // (And no w:showingPlcHdr: with it Word shows the value as the placeholder and reads the field as empty.)
        tr.removeMark(f.from, f.to, f.mark).addMark(f.from, f.to, schema.marks.sdt.create({ ...f.mark.attrs, phText: null, pr: dropPlc(f.mark.attrs.pr || '') }));
        tr.doc.nodesBetween(f.from, f.to, (n, pos) => {
          const m = n.isText && n.marks.find((x) => x.type === ph && x.attrs.val === 'PlaceholderText');
          if (m) tr.removeMark(Math.max(pos, f.from), Math.min(pos + n.nodeSize, f.to), m);
        });
      }
      // Typing into a block field that shows its placeholder: it now holds a real value.
      const sel = tr.selection.$from;
      for (let d = sel.depth; d > 0; d--) {
        const n = sel.node(d);
        if (n.type === schema.nodes.paragraph && n.attrs.sdt && PLC.test(n.attrs.sdt.open)) {
          const g = n.attrs.sdt.g;
          tr.doc.descendants((m, pos) => {
            if (m.type !== schema.nodes.paragraph || m.attrs.sdt?.g !== g) return !m.isTextblock;
            tr.setNodeMarkup(pos, null, { ...m.attrs, sdt: noPlc(m.attrs.sdt) });
            tr.removeMark(pos + 1, pos + m.nodeSize - 1, schema.marks.cstyle.create({ val: 'PlaceholderText' }));
            return false;
          });
        }
      }
      return tr.docChanged ? tr.setMeta(formsKey, { allow: true }) : null;
    },
    props: {
      // A paste too long for a legacy text field is cut to fit (typing past its maximum length is refused).
      handlePaste: (view, event, slice) => formsKey.getState(view.state) && pasteToFit(view, slice),
      // Tab and Shift+Tab move between fields, as in a protected Word form.
      handleKeyDown(view, e) {
        if (!formsKey.getState(view.state)) return false;
        const fields = listFields(view.state.doc);
        if (e.key === ' ') {
          // Space ticks a check box the cursor is on.
          const at = view.state.selection.from;
          const box = fields.find((x) => x.kind === 'checkbox' && at >= x.from && at <= x.to);
          if (!box) return false;
          const tr = view.state.tr.setMeta(formsKey, { allow: true }).setMeta('odx-untracked', true);
          writeField(tr, box, !box.checked);
          view.dispatch(tr.setSelection(TextSelection.create(tr.doc, box.from, box.from + 1)));
          announce(view, describeField({ ...box, checked: !box.checked }));
          return true;
        }
        // Alt+Down (or Enter) opens a drop-down's list or a date's calendar, as in Word.
        if ((e.key === 'ArrowDown' && e.altKey) || e.key === 'Enter') {
          const at = view.state.selection.from;
          const f = fields.find((x) => !x.block && at >= x.from && at <= x.to && (x.kind === 'date' || ((x.kind === 'dropdown' || x.kind === 'combobox') && x.items?.length)));
          if (!f) return false;
          const c = view.coordsAtPos(f.from);
          const ev = { clientX: c.left, clientY: c.bottom };
          const put = (v) => {
            const tr = view.state.tr.setMeta(formsKey, { allow: true }).setMeta('odx-untracked', true);
            const fresh = listFields(view.state.doc).find((x) => x.key === f.key);
            if (!fresh) return;
            writeField(tr, fresh, v);
            view.dispatch(tr);
            announce(view, describeField(listFields(view.state.doc).find((x) => x.key === f.key) || fresh));
          };
          if (f.kind === 'date') pickDate(view, f, ev, (d) => put(formatDate(d, f.pr)));
          else pickFromList(view, f, ev, put);
          return true;
        }
        if (e.key !== 'Tab') return false;
        if (!fields.length) return false;
        const here = view.state.selection.from;
        const i = fields.findIndex((x) => here >= x.from && here <= x.to);
        const next = e.shiftKey
          ? (i >= 0 ? fields[(i - 1 + fields.length) % fields.length] : [...fields].reverse().find((x) => x.to < here) || fields.at(-1))
          : (i >= 0 ? fields[(i + 1) % fields.length] : fields.find((x) => x.from > here) || fields[0]);
        view.dispatch(view.state.tr.setSelection(fieldSelection(view.state.doc, next)).scrollIntoView());
        announce(view, describeField(next));
        return true;
      },
      // Legacy text fields have no mark of their own: shade their result text, as Word does on screen.
      decorations(state) {
        const set = shadeKey.getState(state);
        return set || null;
      },
      attributes: (state) => (formsKey.getState(state) ? { class: 'odx-forms' } : {}),
      handleClick(view, pos, event) {
        if (!formsKey.getState(view.state)) return false;
        const fields = listFields(view.state.doc);
        const f = fields.find((x) => pos >= x.from && pos <= x.to);
        if (!f && openSectionAt(view.state.doc, pos)) return false;
        if (!f) {
          // A click on the form's own text: the caret goes to the field after it (see appendTransaction), here once the
          // click is over, so the browser's caret moves there too and typing isn't lost.
          if (!fields.length) return false;
          view.dispatch(view.state.tr.setSelection(fieldSelection(view.state.doc, fields.find((x) => x.from >= pos) || fields[0])));
          // (When the caret was in that field already, the selection hasn't changed and the browser's caret, put on the
          // clicked text, would stay there: set it to ours.)
          const toDom = () => { const sel = view.state.selection; const a = view.domAtPos(sel.anchor), h = view.domAtPos(sel.head); getSelection().setBaseAndExtent(a.node, a.offset, h.node, h.offset); };
          toDom();
          setTimeout(toDom); // (and again once the browser has finished with the click)
          return true;
        }
        if (f.block) {
          // Clicking a block field's placeholder selects it, so typing replaces it.
          if (f.value !== '' || !f.text) return false;
          view.dispatch(view.state.tr.setSelection(view.state.selection.constructor.create(view.state.doc, f.from + 1, f.to - 1)));
          return true;
        }
        const name = f.tag || f.alias;
        if (f.kind === 'checkbox') {
          const tr = view.state.tr.setMeta(formsKey, { allow: true }).setMeta('odx-untracked', true);
          writeField(tr, f, !f.checked);
          view.dispatch(tr);
          announce(view, describeField({ ...f, checked: !f.checked }));
          return true;
        }
        if (f.kind === 'date') {
          pickDate(view, f, event, (d) => {
            const tr = view.state.tr.setMeta(formsKey, { allow: true }).setMeta('odx-untracked', true);
            const fresh = listFields(view.state.doc).find((x) => x.key === f.key);
            if (fresh) { writeField(tr, fresh, formatDate(d, f.pr)); view.dispatch(tr); }
          });
          return false; // the cursor still goes into the field, so the date can be typed too
        }
        if ((f.kind === 'dropdown' || f.kind === 'combobox') && f.items?.length) {
          pickFromList(view, f, event, (v) => {
            const tr = view.state.tr.setMeta(formsKey, { allow: true }).setMeta('odx-untracked', true);
            const fresh = listFields(view.state.doc).find((x) => x.key === f.key);
            if (fresh) { writeField(tr, fresh, v); view.dispatch(tr); }
          });
          return true;
        }
        // Clicking a placeholder selects it, so typing replaces it.
        if (f.value === '' && f.mark?.attrs.phText != null && name !== undefined) {
          view.dispatch(view.state.tr.setSelection(view.state.selection.constructor.create(view.state.doc, f.from, f.to)));
          return true;
        }
        return false;
      },
    },
  })];
}

/** Where the caret goes in a field reached by Tab: its placeholder selected (typing replaces it), else its end. */
function fieldSelection(doc, f) {
  const a = f.block ? f.from + 1 : f.from;
  const b = f.block ? f.to - 1 : f.to;
  const ph = f.value === '' && f.kind !== 'checkbox';
  return TextSelection.create(doc, ph ? a : b, b);
}

function pickFromList(view, f, event, done) {
  document.querySelector('.odx-fieldpick')?.remove();
  const sel = document.createElement('select');
  sel.className = 'odx-fieldpick';
  sel.size = Math.min(8, f.items.length);
  sel.setAttribute('aria-label', f.alias || f.tag || 'Choose an item');
  for (const it of f.items) {
    const o = document.createElement('option');
    o.value = it.value ?? it.text;
    o.textContent = it.text;
    if (it.text === f.text) o.selected = true;
    sel.append(o);
  }
  sel.style.left = `${event.clientX}px`;
  sel.style.top = `${event.clientY + 8}px`;
  const close = () => { document.removeEventListener('pointerdown', outside, true); if (sel.isConnected) sel.remove(); };
  // Closed by a click elsewhere, not by losing focus: a layout pass refocusing the editor had closed it under the
  // reader's hand (and the forms test's, two runs in three).
  const outside = (e) => { if (e.target !== sel && !sel.contains(e.target)) close(); };
  document.addEventListener('pointerdown', outside, true);
  // A click picks; the arrow keys only move through the list, Enter picks (as in Word): a change on Down had picked and
  // closed it, and the Enter that followed, back in the document, opened it again (NVDA: "Color, list" after "Red").
  let keyed = false;
  const pick = () => { done(sel.value); close(); view.focus(); };
  sel.onchange = () => { if (!keyed) pick(); };
  sel.onkeydown = (e) => {
    keyed = /^(Arrow|Page|Home|End)/.test(e.key);
    if (e.key === 'Escape') { e.preventDefault(); close(); view.focus(); }
    if (e.key === 'Enter') { e.preventDefault(); pick(); }
  };
  sel.onpointerdown = () => { keyed = false; };
  document.body.append(sel);
  sel.focus();
}

export { formatDate };

function pickDate(view, f, event, done) {
  document.querySelector('.odx-fieldpick')?.remove();
  const inp = document.createElement('input');
  inp.type = 'date';
  inp.className = 'odx-fieldpick';
  inp.setAttribute('aria-label', `Choose a date for ${f.alias || f.tag || 'this field'}`);
  const cur = (/w:fullDate="(\d{4}-\d\d-\d\d)/.exec(f.pr) || [])[1];
  if (cur && f.value) inp.value = cur;
  inp.style.left = `${event.clientX}px`;
  inp.style.top = `${event.clientY + 8}px`;
  const close = () => { document.removeEventListener('pointerdown', outside, true); if (inp.isConnected) inp.remove(); };
  // (Closed by a click elsewhere, not by losing focus, as the drop-down's list.)
  const outside = (e) => { if (e.target !== inp) close(); };
  document.addEventListener('pointerdown', outside, true);
  inp.onchange = () => { if (!inp.value) return; const [y, m, d] = inp.value.split('-').map(Number); done(new Date(y, m - 1, d)); close(); view.focus(); };
  inp.onkeydown = (e) => { if (e.key === 'Escape') { close(); view.focus(); } };
  document.body.append(inp);
  try { inp.showPicker?.(); } catch { /* not allowed without a gesture */ }
}

/** Turn forms mode on or off. */
export function setForms(view, on) {
  view.dispatch(view.state.tr.setMeta(formsKey, { on: !!on }));
}

export { xmlEsc };
