// Fields in drawn headers and footers: PAGE, NUMPAGES and the like get their live values, and the document's
// properties (TITLE, AUTHOR, DOCPROPERTY ...), which Word evaluates again there.
import { formatNumber } from '../ooxml/styles.js';
import { unescapeXml } from '../ooxml/xml.js';

// A field's value styled like the text before it (for a field saved with no result).
function styledLike(el, text) {
  let prev = el.previousSibling;
  while (prev && !(prev.nodeType === 1 && !prev.classList.contains('odx-opaque-inline') && prev.textContent)) prev = prev.previousSibling;
  if (!prev) { const t = document.createElement('span'); t.className = 'odx-t'; t.textContent = text; return t; } // (its own text span: textspans.js)
  const copy = prev.cloneNode(true);
  let inner = copy;
  while (inner.firstElementChild) inner = inner.firstElementChild;
  inner.textContent = text;
  return copy;
}

const PROP_FIELDS = { TITLE: 'Title', SUBJECT: 'Subject', AUTHOR: 'Author', KEYWORDS: 'Keywords', COMMENTS: 'Comments' };
// Word's own property names (DOCPROPERTY): ones we don't read keep their saved results; any other name is an error.
const BUILTIN = new Set(['title', 'subject', 'author', 'keywords', 'comments', 'template', 'lastsavedby', 'revisionnumber',
  'application', 'lastprinted', 'createtime', 'lastsavedtime', 'editingtime', 'pages', 'words', 'characters', 'security',
  'category', 'format', 'manager', 'company', 'bytes', 'lines', 'paragraphs', 'slides', 'notes', 'hiddenslides',
  'multimediaclips', 'hyperlinkbase', 'characterswithspaces', 'charactersspaces', 'nameofapplication']);

/** The file's properties by lowercase name: the core ones (Title, Subject, Author...), the app's, and custom ones. */
export function docProps(pkg) {
  const out = new Map();
  const core = (pkg && pkg.text('docProps/core.xml')) || '';
  const app = (pkg && pkg.text('docProps/app.xml')) || '';
  const tag = (xml, t) => { const m = new RegExp(`<${t}(?:\\s[^>]*)?>([^<]*)</${t}>`).exec(xml); return m ? unescapeXml(m[1]) : null; };
  for (const [k, t] of [['title', 'dc:title'], ['subject', 'dc:subject'], ['author', 'dc:creator'], ['keywords', 'cp:keywords'], ['comments', 'dc:description'], ['category', 'cp:category'], ['lastsavedby', 'cp:lastModifiedBy']]) {
    const v = tag(core, t);
    out.set(k, v ?? '');
  }
  for (const [k, t] of [['company', 'Company'], ['manager', 'Manager']]) { const v = tag(app, t); if (v != null) out.set(k, v); }
  const custom = (pkg && pkg.text('docProps/custom.xml')) || '';
  for (const m of custom.matchAll(/<property\b[^>]*\bname="([^"]+)"[^>]*>\s*<vt:\w+>([^<]*)<\/vt:\w+>/g)) out.set(unescapeXml(m[1]).toLowerCase(), unescapeXml(m[2]));
  return out;
}

/**
 * Replace the cached results of PAGE-style fields with live values. Handles
 * simple fields (w:fldSimple) and complex ones (fldChar begin / instrText /
 * separate / result runs / end), using the data-* hints the schema renders.
 */
export function fillFields(root, values) {
  const nameOf = (instr) => (/^\s*([A-Z]+)/.exec(instr || '') || [])[1];
  // A page number in its field's format switch (\* roman, \* ALPHABETIC...), else the section's format.
  const FMT = { roman: 'lowerRoman', Roman: 'upperRoman', ROMAN: 'upperRoman', alphabetic: 'lowerLetter', ALPHABETIC: 'upperLetter', Arabic: null, arabic: null };
  const valueOf = (instr) => {
    const name = nameOf(instr);
    // Property fields Word evaluates again in headers and footers (the body keeps their saved results).
    const props = values.PROPS;
    if (props && PROP_FIELDS[name]) return props.get(PROP_FIELDS[name].toLowerCase()) ?? '';
    if (props && name === 'DOCPROPERTY') {
      const m = /DOCPROPERTY\s+(?:"([^"]+)"|(\S+))/.exec(instr);
      const key = m && (m[1] || m[2]).toLowerCase();
      if (!key) return null;
      if (props.has(key)) return props.get(key);
      return BUILTIN.has(key) ? null : 'Error! Unknown document property name.';
    }
    const v = values[name];
    if (name !== 'PAGE' || v == null) return v;
    const sw = /\\\*\s*(\w+)/.exec(instr || '');
    return formatNumber(Number(v), sw && sw[1] in FMT ? FMT[sw[1]] : values.PAGEFMT);
  };
  root.querySelectorAll('[data-kind="fldSimple"][data-instr]').forEach((f) => {
    const v = valueOf(f.dataset.instr);
    if (v == null) return;
    const t = document.createElement('span');
    t.className = 'odx-t';
    t.textContent = String(v);
    f.replaceChildren(t);
  });
  let instr = '';
  let mode = null; // 'instr' | 'result'
  let value = null;
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n.nodeType === 1) {
      const fld = n.dataset && n.dataset.fld;
      if (fld === 'begin') { mode = 'instr'; instr = ''; value = null; }
      else if (fld === 'separate') { value = valueOf(instr); mode = 'result'; }
      else if (fld === 'end') {
        // A field saved with no result text: show the value, styled like the text before it.
        if (mode === 'result' && value != null && value !== '') n.before(styledLike(n, String(value)));
        mode = null;
      }
      else if (mode === 'instr' && n.dataset && n.dataset.instr != null) instr += n.dataset.instr;
    } else if (mode === 'result' && value != null && n.parentElement.closest('.odx-opaque-inline') == null) {
      n.textContent = String(value);
      value = '';
    }
  }
}
