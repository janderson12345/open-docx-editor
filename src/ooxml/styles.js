// styles.xml, theme fonts and numbering.xml: read once per document, turned
// into a CSS stylesheet for the editor and a list-label resolver.
import { kids, kid, attr, onOff, NS, REL_TYPES } from './xml.js';
import { readRPr, readPPr, rPrToCss, pPrToCss, fontStack, fontCss, lineFactor, sizeCss, mergeTabs, readBorders, shdFill, borderCss } from './props.js';
import { twipsToPx } from './units.js';

export function readTheme(pkg) {
  const path = pkg.partByType(REL_TYPES.theme);
  const doc = path && pkg.xml(path);
  const theme = { major: 'Calibri Light', minor: 'Calibri' };
  if (!doc) return theme;
  for (const kind of ['major', 'minor']) {
    const f = doc.getElementsByTagNameNS(NS.a, kind + 'Font')[0];
    const latin = f && kid(f, 'latin', NS.a);
    if (latin) theme[kind] = latin.getAttribute('typeface');
    const cs = f && kid(f, 'cs', NS.a);
    if (cs && cs.getAttribute('typeface')) theme[kind + 'Cs'] = cs.getAttribute('typeface');
  }
  theme.colors = {};
  const scheme = doc.getElementsByTagNameNS(NS.a, 'clrScheme')[0];
  for (const c of scheme ? kids(scheme, null, NS.a) : []) {
    const v = kids(c, null, NS.a)[0];
    const hex = v && (v.getAttribute('val') && v.localName === 'srgbClr' ? v.getAttribute('val') : v.getAttribute('lastClr'));
    if (hex) theme.colors[c.localName] = hex;
  }
  return theme;
}

export function readStyles(pkg) {
  const path = pkg.partByType(REL_TYPES.styles);
  const doc = path && pkg.xml(path);
  const styles = new Map();
  const defaults = { rPr: {}, pPr: {} };
  // With no styles part at all, Word's own defaults: Aptos 12pt, 8pt after, 1.15 lines (Word's PDF of such a file).
  if (!doc) return { styles, defaults: { rPr: { font: 'Aptos', size: 24 }, pPr: { after: 160, line: 276, lineRule: 'auto' } }, defaultParagraph: null, defaultTable: null };
  const dd = kid(doc.documentElement, 'docDefaults');
  if (dd) {
    defaults.rPr = readRPr(kid(kid(dd, 'rPrDefault'), 'rPr'));
    defaults.pPr = readPPr(kid(kid(dd, 'pPrDefault'), 'pPr'));
  }
  let defaultParagraph = null;
  let defaultTable = null; // the table style a table that names none gets (Normal Table: its cell margins)
  for (const s of kids(doc.documentElement, 'style')) {
    const id = attr(s, 'styleId');
    const st = {
      id,
      type: attr(s, 'type'),
      name: attr(kid(s, 'name'), 'val') || id,
      basedOn: attr(kid(s, 'basedOn'), 'val'),
      next: attr(kid(s, 'next'), 'val'),
      isDefault: attr(s, 'default') === '1' || attr(s, 'default') === 'true',
      hidden: !!kid(s, 'semiHidden') || !!kid(s, 'hidden'),
      unhide: !!kid(s, 'unhideWhenUsed'), // hidden until used (Heading 2 to 9 in Word's template)
      uiPriority: Number(attr(kid(s, 'uiPriority'), 'val') ?? 99),
      qFormat: !!kid(s, 'qFormat'),
      rPr: readRPr(kid(s, 'rPr')),
      pPr: readPPr(kid(s, 'pPr')),
    };
    if (st.type === 'table') st.table = readTableStyle(s);
    styles.set(id, st);
    if (st.type === 'paragraph' && st.isDefault) defaultParagraph = id;
    if (st.type === 'table' && st.isDefault) defaultTable = id;
  }
  return { styles, defaults, defaultParagraph, defaultTable };
}

/** Merge a style with everything it is based on. */
export function resolveStyle(styles, id, seen = new Set()) {
  const st = styles.get(id);
  if (!st || seen.has(id)) return { rPr: {}, pPr: {} };
  seen.add(id);
  const base = st.basedOn ? resolveStyle(styles, st.basedOn, seen) : { rPr: {}, pPr: {} };
  const pPr = { ...base.pPr, ...st.pPr };
  if (base.pPr.tabs || st.pPr.tabs) pPr.tabs = mergeTabs(base.pPr.tabs, st.pPr.tabs);
  return { rPr: { ...base.rPr, ...st.rPr }, pPr };
}

/** CSS for the page: document defaults plus one rule per paragraph/character style. */
export function stylesToCss(scope, info, theme) {
  const { styles, defaults, defaultParagraph } = info;
  const rules = [];
  // With no font in the document defaults Word sets text in Times New Roman, not the theme's (rich: an empty
  // w:rPrDefault, a Calibri theme, Word's PDF in Times); the theme's fonts only where something names them.
  const baseFont = defaults.rPr.font || (defaults.rPr.themeFont ? theme[defaults.rPr.themeFont.startsWith('major') ? 'major' : 'minor'] : 'Times New Roman');
  // With no w:sz in the document defaults Word uses 10pt.
  rules.push(
    // The theme's fonts, for runs that name them (w:asciiTheme): as variables, with their line heights.
    `${scope}{--odx-font-minor:${fontStack(theme.minor || 'Calibri')};--odx-lh-minor:${lineFactor(theme.minor || 'Calibri')};--odx-font-major:${fontStack(theme.major || theme.minor || 'Calibri')};--odx-lh-major:${lineFactor(theme.major || theme.minor || 'Calibri')}}`,
    `${scope}{${fontCss(baseFont || 'Calibri')};${sizeCss((defaults.rPr.size || 20) / 2)};${rPrToCss({ ...defaults.rPr, font: null, themeFont: null, size: null }, theme)}}`,
  );
  rules.push(`${scope} p{margin:0;margin-inline-start:calc(0pt - var(--odx-bdl, 0pt));margin-inline-end:calc(0pt - var(--odx-bdr, 0pt));direction:ltr;${pPrToCss(defaults.pPr)}}`); // (left-to-right unless it says otherwise, also in a right-to-left table)
  for (const st of styles.values()) {
    if (st.type !== 'paragraph' && st.type !== 'character') continue;
    const r = resolveStyle(styles, st.id);
    const sel = st.type === 'paragraph'
      ? `${scope} p[data-style="${cssEsc(st.id)}"]` + (st.id === defaultParagraph ? `,${scope} p:not([data-style])` : '')
      : `${scope} span[data-cstyle="${cssEsc(st.id)}"]`;
    const css = [st.type === 'paragraph' ? pPrToCss(r.pPr) : '', rPrToCss(r.rPr, theme)].filter(Boolean).join(';');
    if (css) rules.push(`${sel}{${css}}`);
  }
  // Built-in headings a document uses without defining them (python-docx's add_heading in a bare template): drawn as
  // Word's own Heading 1 to 9 and Title, which Word applies to them.
  const major = `font-family:var(--odx-font-major);--odx-lh:var(--odx-lh-major)`;
  const BUILT_IN = {
    Heading1: 'font-size:16pt;color:#2F5496;margin-top:12pt', Heading2: 'font-size:13pt;color:#2F5496;margin-top:2pt',
    Heading3: 'font-size:12pt;color:#1F3763;margin-top:2pt', Heading4: 'font-style:italic;color:#2F5496;margin-top:2pt',
    Heading5: 'color:#2F5496;margin-top:2pt', Heading6: 'color:#1F3763;margin-top:2pt', Heading7: 'font-style:italic;color:#1F3763;margin-top:2pt',
    Heading8: 'font-size:10.5pt;color:#272727;margin-top:2pt', Heading9: 'font-size:10.5pt;font-style:italic;color:#272727;margin-top:2pt',
    Title: 'font-size:28pt;letter-spacing:-0.5pt;margin-bottom:0',
  };
  for (const [id, css] of Object.entries(BUILT_IN)) {
    if (!styles.has(id)) rules.push(`${scope} p[data-style="${id}"]{${major};${css};${id === 'Title' ? '' : 'margin-bottom:0'}}`);
  }
  return rules.join('\n');
}

const cssEsc = (s) => String(s).replace(/["\\]/g, '\\$&');

/** Paragraph styles worth offering in a style picker, Word-gallery order. */
export function styleMenu(info, used = null) {
  // A style hidden until used (w:semiHidden with w:unhideWhenUsed) shows once the document uses it, as in Word's
  // gallery, and a heading level once the level above it is used (Heading 2 after Heading 1).
  const level = (s) => Number((/^heading (\d)$/i.exec(s.name) || [])[1]) || 0;
  const usedLevels = new Set([...info.styles.values()].filter((s) => used?.has(s.id)).map(level));
  const inUse = (s) => !used || used.has(s.id) || (level(s) > 1 && usedLevels.has(level(s) - 1));
  return [...info.styles.values()]
    .filter((s) => s.type === 'paragraph' && (!s.hidden || s.qFormat) && !(s.hidden && s.unhide && !inUse(s)))
    .sort((a, b) => a.uiPriority - b.uiPriority || a.name.localeCompare(b.name))
    .map((s) => ({ id: s.id, name: s.name.replace(/^heading (\d)$/, 'Heading $1').replace(/^title$/, 'Title') }));
}

// --- Numbering ------------------------------------------------------------

const BULLET_MAP = { '\uF0B7': '•', '\uF0A7': '▪', '\uF0D8': '➢', '\uF0FC': '✓', '\uF076': '❖', '\uF06E': '■' };

export function readNumbering(pkg) {
  const path = pkg.partByType(REL_TYPES.numbering);
  const doc = path && pkg.xml(path);
  const abstract = new Map();
  const nums = new Map();
  if (!doc) return { abstract, nums };
  const readLvl = (l) => ({
    start: Number(attr(kid(l, 'start'), 'val') ?? 1),
    fmt: attr(kid(l, 'numFmt'), 'val') || 'decimal',
    text: attr(kid(l, 'lvlText'), 'val') ?? '',
    pPr: readPPr(kid(l, 'pPr')),
    font: attr(kid(kid(l, 'rPr'), 'rFonts'), 'ascii') || null,
    size: Number(attr(kid(kid(l, 'rPr'), 'sz'), 'val')) || null, // half-points
    // The label's own bold, italic, colour and underline (null where the level doesn't set them).
    bold: onOff(kid(kid(l, 'rPr'), 'b')),
    italic: onOff(kid(kid(l, 'rPr'), 'i')),
    color: attr(kid(kid(l, 'rPr'), 'color'), 'val') || null,
    underline: attr(kid(kid(l, 'rPr'), 'u'), 'val') || null,
    jc: attr(kid(l, 'lvlJc'), 'val') || 'left', // the label's alignment at its position
    suff: attr(kid(l, 'suff'), 'val') || 'tab', // what follows the label: a tab, a space or nothing
    legal: onOff(kid(l, 'isLgl')) || false, // legal numbering: every level's number in Arabic numerals ("1.1.1", not "I.1.a")
    picBullet: attr(kid(l, 'lvlPicBulletId'), 'val') ?? null, // a picture bullet (w:numPicBullet)
    hidden: onOff(kid(kid(l, 'rPr'), 'vanish')) || false, // the number is hidden text: not shown

  });
  for (const an of kids(doc.documentElement, 'abstractNum')) {
    const levels = [];
    for (const l of kids(an, 'lvl')) levels[Number(attr(l, 'ilvl'))] = readLvl(l);
    abstract.set(attr(an, 'abstractNumId'), levels);
  }
  for (const n of kids(doc.documentElement, 'num')) {
    const overrides = {};
    const own = []; // levels the num defines itself (w:lvlOverride/w:lvl), in place of its abstract numbering's
    for (const o of kids(n, 'lvlOverride')) {
      const so = kid(o, 'startOverride');
      if (so) overrides[Number(attr(o, 'ilvl'))] = Number(attr(so, 'val'));
      const lv = kid(o, 'lvl');
      if (lv) own[Number(attr(o, 'ilvl'))] = readLvl(lv);
    }
    const abstractId = attr(kid(n, 'abstractNumId'), 'val');
    // (sample S071's bullets: its abstract level is a "*", the num's override a Symbol bullet, as Word shows.)
    const base = abstract.get(abstractId);
    const levels = own.length && base ? base.map((l, i) => own[i] || l) : undefined;
    if (levels) own.forEach((l, i) => { if (l && !levels[i]) levels[i] = l; });
    nums.set(attr(n, 'numId'), { abstractId, overrides, levels });
  }
  return { abstract, nums };
}

export function listLevel(numbering, numId, ilvl) {
  const n = numbering.nums.get(String(numId));
  const levels = n && (n.levels || numbering.abstract.get(n.abstractId));
  return levels ? levels[ilvl] || null : null;
}

export function formatNumber(n, fmt) {
  switch (fmt) {
    case 'lowerLetter': return letters(n).toLowerCase();
    case 'upperLetter': return letters(n);
    case 'lowerRoman': return roman(n).toLowerCase();
    case 'upperRoman': return roman(n);
    case 'decimalZero': return n < 10 ? '0' + n : String(n);
    case 'chineseCounting': case 'chineseCountingThousand': case 'ideographDigital': return chinese(n, fmt === 'ideographDigital');
    case 'hebrew1': return hebrewNumber(n);
    case 'hebrew2': return n > 0 ? HEBREW[(n - 1) % 22].repeat(Math.floor((n - 1) / 22) + 1) : '';
    case 'ordinal': return ordinal(n);
    case 'cardinalText': return capital(words(n));
    case 'ordinalText': return capital(ordinalWords(n));
    case 'decimalEnclosedCircle': case 'decimalEnclosedCircleChinese': return n >= 1 && n <= 20 ? String.fromCharCode(0x2460 + n - 1) : String(n);
    case 'decimalEnclosedParen': return n >= 1 && n <= 20 ? String.fromCharCode(0x2474 + n - 1) : String(n);
    case 'decimalEnclosedFullstop': return n >= 1 && n <= 20 ? String.fromCharCode(0x2488 + n - 1) : String(n);
    case 'decimalFullWidth': case 'decimalFullWidth2': return String(n).replace(/\d/g, (d) => String.fromCharCode(0xFF10 + Number(d)));
    case 'arabicAlpha': return cycle(n, ARABIC_ALPHA);
    case 'arabicAbjad': return cycle(n, ARABIC_ABJAD);
    case 'russianLower': return cycle(n, RUSSIAN);
    case 'russianUpper': return cycle(n, RUSSIAN).toUpperCase();
    // (Ten stems and twelve branches, then digits, as Word numbers past them.)
    case 'ideographTraditional': return n >= 1 && n <= 10 ? '甲乙丙丁戊己庚辛壬癸'[n - 1] : String(n);
    case 'ideographZodiac': return n >= 1 && n <= 12 ? '子丑寅卯辰巳午未申酉戍亥'[n - 1] : String(n);
    case 'none': return '';
    default: return String(n);
  }
}
// An alphabet's letters in turn, then each doubled, tripled... as Word's lettered lists go on.
const cycle = (n, abc) => (n > 0 ? abc[(n - 1) % abc.length].repeat(Math.floor((n - 1) / abc.length) + 1) : String(n));
const ARABIC_ALPHA = [...'أبتثجحخدذرزسشصضطظعغفقكلمنهوي'];
const ARABIC_ABJAD = [...'أبجدهوزحطيكلمنسعفصقرشتثخذضظغ'];
const RUSSIAN = [...'абвгдежзиклмнопрстуфхцчшщыэюя'];
// English ordinals and numbers in words (Word's ordinal, cardinalText and ordinalText).
const ordinal = (n) => `${n}${n % 100 >= 11 && n % 100 <= 13 ? 'th' : ['th', 'st', 'nd', 'rd'][n % 10] || 'th'}`;
const capital = (s) => s.charAt(0).toUpperCase() + s.slice(1);
const ONES = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen'];
const TENS = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];
function words(n) {
  if (n < 0 || n > 999999) return String(n);
  if (n < 20) return ONES[n];
  if (n < 100) return TENS[Math.floor(n / 10)] + (n % 10 ? `-${ONES[n % 10]}` : '');
  if (n < 1000) return `${ONES[Math.floor(n / 100)]} hundred${n % 100 ? ` ${words(n % 100)}` : ''}`;
  return `${words(Math.floor(n / 1000))} thousand${n % 1000 ? ` ${words(n % 1000)}` : ''}`;
}
const ORD = { one: 'first', two: 'second', three: 'third', five: 'fifth', eight: 'eighth', nine: 'ninth', twelve: 'twelfth' };
function ordinalWords(n) {
  const w = words(n);
  return w.replace(/(\w+)$/, (last) => ORD[last] || (last.endsWith('y') ? `${last.slice(0, -1)}ieth` : `${last}th`));
}
// Chinese numerals (一、二、三... 十, 十一, 二十, 一百零一), Word's chineseCounting and chineseCountingThousand; digit by
// digit (一〇一) for ideographDigital.
const CN = '〇一二三四五六七八九';
function chinese(n, digital) {
  if (digital) return String(n).split('').map((d) => CN[d]).join('');
  if (n <= 0 || n > 9999) return String(n);
  const units = ['', '十', '百', '千'];
  const ds = String(n).split('').map(Number);
  let s = '', zero = false;
  ds.forEach((d, i) => {
    const u = units[ds.length - 1 - i];
    if (!d) { zero = s !== ''; return; }
    if (zero) { s += '零'; zero = false; }
    s += (d === 1 && u === '十' && s === '' ? '' : CN[d]) + u;
  });
  return s;
}
// Hebrew letters as numbers (hebrew1: א=1... י=10... ק=100), 15 and 16 written טו and טז as Word does; hebrew2 is the
// alphabet in order.
const HEBREW = 'אבגדהוזחטיכלמנסעפצקרשת';
function hebrewNumber(n) {
  if (n <= 0) return String(n);
  const vals = [[400, 'ת'], [300, 'ש'], [200, 'ר'], [100, 'ק'], [90, 'צ'], [80, 'פ'], [70, 'ע'], [60, 'ס'], [50, 'נ'], [40, 'מ'], [30, 'ל'], [20, 'כ'], [10, 'י'], [9, 'ט'], [8, 'ח'], [7, 'ז'], [6, 'ו'], [5, 'ה'], [4, 'ד'], [3, 'ג'], [2, 'ב'], [1, 'א']];
  let s = '';
  let r = n;
  while (r >= 400) { s += 'ת'; r -= 400; }
  if (r % 100 === 15) { r -= 15; s += hebrewNumber(r).replace(/^\d+$/, '') + 'טו'; return s; }
  if (r % 100 === 16) { r -= 16; s += hebrewNumber(r).replace(/^\d+$/, '') + 'טז'; return s; }
  for (const [v, ch] of vals) while (r >= v) { s += ch; r -= v; }
  return s;
}
// A, B... Z, then AA, BB... ZZ, AAA: Word's letters for list items and page numbers alike.
const letters = (n) => (n > 0 ? String.fromCharCode(65 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1) : '');
const roman = (n) => {
  const t = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let s = '';
  for (const [v, r] of t) while (n >= v) { s += r; n -= v; }
  return s;
};

/**
 * Walk paragraphs in order and compute each list label ("1.", "a)", "•").
 * Returns a function: (numId, ilvl) -> label, to be called in document order.
 */
export function listCounter(numbering) {
  const counters = new Map(); // abstractId -> number[]
  return (numId, ilvl) => {
    const n = numbering.nums.get(String(numId));
    if (!n) return null;
    const levels = n.levels || numbering.abstract.get(n.abstractId);
    const lvl = levels && levels[ilvl];
    if (!lvl) return null;
    if (lvl.fmt === 'bullet') return BULLET_MAP[lvl.text] || lvl.text || '•';
    let c = counters.get(n.abstractId);
    if (!c) counters.set(n.abstractId, (c = []));
    const startOf = (i) => n.overrides[i] ?? levels[i]?.start ?? 1;
    c[ilvl] = c[ilvl] == null ? startOf(ilvl) : c[ilvl] + 1;
    c.length = ilvl + 1; // deeper levels restart
    return lvl.text.replace(/%(\d)/g, (_, d) => {
      const i = Number(d) - 1;
      const v = c[i] ?? startOf(i);
      const fmt = levels[i]?.fmt || 'decimal';
      // (Legal numbering makes each level Arabic; one with leading zeros keeps them: SECTION 1.07.)
      return formatNumber(v, lvl.legal && fmt !== 'none' && fmt !== 'decimalZero' ? 'decimal' : fmt);
    });
  };
}

// --- Table styles -----------------------------------------------------------

function readTableStyle(s) {
  const part = (el) => ({
    borders: readBorders(kid(kid(el, 'tblPr'), 'tblBorders')),
    cellBorders: readBorders(kid(kid(el, 'tcPr'), 'tcBorders')),
    shading: shdFill(kid(kid(el, 'tcPr'), 'shd')),
    cellMar: readCellMar(kid(kid(el, 'tblPr'), 'tblCellMar')),
    rPr: readRPr(kid(el, 'rPr')),
    pPr: readPPr(kid(el, 'pPr')),
  });
  const t = part(s);
  t.cond = {};
  for (const c of kids(s, 'tblStylePr')) t.cond[attr(c, 'type')] = part(c);
  return t;
}

export function readCellMar(el) {
  const out = {};
  for (const c of kids(el)) out[{ start: 'left', end: 'right' }[c.localName] || c.localName] = twipsToPx(attr(c, 'w') || 0);
  return out;
}

/** Merge a table style with the ones it is based on. */
export function resolveTableStyle(styles, id, seen = new Set()) {
  const st = styles.get(id);
  if (!st || !st.table || seen.has(id)) return null;
  seen.add(id);
  const base = st.basedOn ? resolveTableStyle(styles, st.basedOn, seen) : null;
  if (!base) return st.table;
  const merge = (a, b) => ({
    borders: { ...a.borders, ...b.borders },
    cellBorders: { ...a.cellBorders, ...b.cellBorders },
    shading: b.shading ?? a.shading,
    cellMar: { ...a.cellMar, ...b.cellMar },
    rPr: { ...a.rPr, ...b.rPr },
    pPr: { ...a.pPr, ...b.pPr },
  });
  const out = merge(base, st.table);
  out.cond = { ...base.cond };
  for (const [k, v] of Object.entries(st.table.cond)) out.cond[k] = base.cond[k] ? merge(base.cond[k], v) : v;
  return out;
}

/**
 * CSS rules for table styles' whole-table text and paragraph formatting.
 * Like Word, a table style's paragraph settings beat the Normal style but not other paragraph styles.
 */
export function tableStylesCss(scope, info, theme) {
  const rules = [];
  const def = info.defaultParagraph;
  for (const st of info.styles.values()) {
    if (st.type !== 'table') continue;
    const t = resolveTableStyle(info.styles, st.id);
    if (!t) continue;
    const sel = (x) => `${scope} table[data-tstyle="${st.id.replace(/["\\]/g, '\\$&')}"] ${x}`;
    // Like its paragraph settings, a table style's text formatting beats the Normal style's.
    // Word counts a font size set in the document defaults as part of the table style here.
    // (A size the table style sets itself loses to one the Normal style sets: Consultant Request Form, Normal's 12pt over
    // Table Grid's 11. The document defaults' size still beats Normal's, sample S125.)
    const normalSize = def ? resolveStyle(info.styles, def).rPr.size : null;
    const ownSize = t.rPr.size != null && normalSize != null && normalSize !== info.defaults.rPr.size ? null : t.rPr.size;
    const r = rPrToCss({ ...t.rPr, size: t.rPr.size != null ? ownSize : info.defaults.rPr.size }, theme);
    if (r) rules.push(`${sel('td')},${sel('td > p:not([data-style])')},${sel(`td > p[data-style="${def}"]`)}{${r}}`);
    // The table style's spacing beats the Normal style's only where Normal leaves it to the document defaults (unset, or
    // set to the same value): sample S072's Normal sets 12pt after, which its Table Grid's 0 doesn't undo.
    const normal = def ? resolveStyle(info.styles, def).pPr : {};
    const beats = (k) => t.pPr[k] != null && (normal[k] == null || normal[k] === info.defaults.pPr[k]);
    const tp = { align: t.pPr.align };
    for (const k of ['before', 'after']) if (beats(k)) tp[k] = t.pPr[k];
    if (beats('line')) { tp.line = t.pPr.line; tp.lineRule = t.pPr.lineRule; }
    const p = pPrToCss(tp);
    if (p) rules.push(`${sel('td > p:not([data-style]),')}${sel(`td > p[data-style="${def}"]`)}{${p}}`);
    // Other paragraph styles take the table style's spacing too, for each setting their own chain leaves to the defaults.
    const keys = ['before', 'after', 'line', 'lineRule'];
    for (const ps of info.styles.values()) {
      if (ps.type !== 'paragraph' || ps.id === def) continue;
      const own = resolveStyle(info.styles, ps.id).pPr;
      const take = Object.fromEntries(keys.filter((k) => t.pPr[k] != null && own[k] == null && !(k === 'lineRule' && own.line != null)).map((k) => [k, t.pPr[k]]));
      if (take.line == null && own.line == null) delete take.lineRule;
      if (take.line != null && take.lineRule == null) take.lineRule = t.pPr.lineRule;
      const css = Object.keys(take).length && pPrToCss(take);
      if (css) rules.push(`${sel(`td > p[data-style="${ps.id.replace(/["\\]/g, '\\$&')}"]`)}{${css}}`);
    }
  }
  const size = info.defaults.rPr.size;
  if (size) {
    const sel = (x) => `${scope} table:not([data-tstyle]) ${x}`;
    rules.push(`${sel('td > p:not([data-style])')},${sel(`td > p[data-style="${def}"]`)}{${sizeCss(size / 2)}}`);
  }
  return rules.join('\n');
}

/**
 * Where a table sits: indent, alignment and floating (w:tblpPr), as CSS for its wrapper.
 * Before Word 2013 (compatibility mode < 15) the indent measures to the first cell's text,
 * so the table's edge sits one left cell margin further out.
 */
// The same table properties give the same placement: remembered per document's styles (the editor asks again for every
// table on every keystroke, and parsing their XML was most of a long document's typing time).
const placements = new WeakMap(); // info -> Map(key -> css)
export function tablePlacementCss(info, tblPrXml, compatMode, parseXml) {
  if (!tblPrXml) return compatMode < 15 ? 'margin-left:-7.2px' : '';
  if (!placements.has(info)) placements.set(info, new Map());
  const memo = placements.get(info);
  const key = `${compatMode}|${tblPrXml}`;
  if (!memo.has(key)) memo.set(key, placementCss(info, tblPrXml, compatMode, parseXml));
  return memo.get(key);
}
function placementCss(info, tblPrXml, compatMode, parseXml) {
  const el = parseXml(tblPrXml.replace(/^<w:tblPr/, '<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"')).documentElement;
  const styleId = attr(kid(el, 'tblStyle'), 'val') || info.defaultTable;
  const ts = (styleId && resolveTableStyle(info.styles, styleId)) || { cellMar: {} };
  const mar = { ...ts.cellMar, ...readCellMar(kid(el, 'tblCellMar')) };
  const ind = kid(el, 'tblInd');
  const indPx = ind && attr(ind, 'type') !== 'pct' ? twipsToPx(attr(ind, 'w') || 0) : 0;
  const left = indPx - (compatMode < 15 ? mar.left ?? 7.2 : 0);
  const jc = attr(kid(el, 'jc'), 'val');
  const css = [];
  const fp = kid(el, 'tblpPr');
  if (fp) {
    const tw = (k) => twipsToPx(attr(fp, k) || 0);
    // Set down at a place on its page (vertAnchor page or margin): out of the floats, in its own band; pagination
    // (pages.js, floatBands) puts it there and lets the text after it fill the page above and below it, as Word does
    // (the contingency plan's validation table at 9877tw down page 4, the captions after it in the file above it).
    const v = attr(fp, 'vertAnchor');
    if ((v === 'page' || v === 'margin') && !attr(fp, 'tblpYSpec')) {
      const x = attr(fp, 'horzAnchor') === 'page' ? `calc(${tw('tblpX')}px - var(--odx-page-left, 96px))` : `${left + tw('tblpX')}px`;
      return [`margin-left:${x}`, `--odx-fy:${tw('tblpY')}px`, `--odx-fyrel:${v}`, `--odx-fyb:${tw('bottomFromText')}px`, `--odx-fyt:${tw('topFromText')}px`].join(';');
    }
    const right = attr(fp, 'tblpXSpec') === 'right' || jc === 'right';
    css.push('width:fit-content', `float:${right ? 'right' : 'left'}`, `position:relative;z-index:1`);
    css.push(`margin-left:${right ? tw('leftFromText') : left + tw('tblpX')}px`, `margin-right:${tw('rightFromText')}px`, `margin-bottom:${tw('bottomFromText')}px`);
    if (tw('tblpY') > 0 && attr(fp, 'vertAnchor') !== 'page') css.push(`top:${tw('tblpY')}px`);
    return css.join(';');
  }
  // A centred table of a set width wider than the column overhangs both margins equally, as in Word (auto margins
  // would leave it against the left one).
  const tw = kid(el, 'tblW');
  const wPx = tw && attr(tw, 'type') === 'dxa' ? twipsToPx(attr(tw, 'w') || 0) : 0;
  if (jc === 'center' && wPx) css.push('width:fit-content', `margin-left:calc((100% - ${+wPx.toFixed(2)}px) / 2)`, 'margin-right:0');
  else if (jc === 'center') css.push('width:fit-content', 'margin-left:auto', 'margin-right:auto');
  else if (jc === 'right' || jc === 'end') css.push('width:fit-content', 'margin-left:auto');
  else if (left) css.push(`margin-left:${+left.toFixed(2)}px`);
  // A right-to-left table (w:bidiVisual): its first column at the right, its indent from the right margin. (Its cells'
  // paragraphs keep their own direction: editor.css and styles.js give every paragraph one.)
  if (onOff(kid(el, 'bidiVisual'))) return css.join(';').replace(/margin-left:/g, 'margin-inline-start:').replace(/margin-right:/g, 'margin-inline-end:') + ';direction:rtl';
  return css.join(';');
}

const LOOK_BITS = { firstRow: 0x20, lastRow: 0x40, firstColumn: 0x80, lastColumn: 0x100, noHBand: 0x200, noVBand: 0x400 };

/**
 * Per-cell CSS for one table: borders (table, style, conditional regions, cell),
 * shading and conditional text formatting. `cells` is [{ row, col, rowspan, colspan, tcPr }].
 */
/**
 * The table properties a table made in the editor (Insert Table without a Table Grid style, a pasted table) is saved
 * with, and drawn with: single lines all round and between its cells.
 */
export const DEFAULT_TBLPR = '<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:tblLook w:val="04A0"/></w:tblPr>';

export function tableCellCss(info, tblPrXml, nRows, nCols, cells, parseXml, theme) {
  tblPrXml ||= DEFAULT_TBLPR; // (as it is saved)
  let styleId = null;
  let direct = {};
  let look = { firstRow: true, firstColumn: true, noVBand: true };
  let cellMar = {};
  if (tblPrXml) {
    const el = parseXml(tblPrXml.replace(/^<w:tblPr/, '<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"')).documentElement;
    styleId = attr(kid(el, 'tblStyle'), 'val');
    direct = readBorders(kid(el, 'tblBorders'));
    cellMar = readCellMar(kid(el, 'tblCellMar'));
    const lk = kid(el, 'tblLook');
    if (lk) {
      const v = parseInt(attr(lk, 'val') || '0', 16);
      look = {};
      for (const [k, bit] of Object.entries(LOOK_BITS)) {
        const a = attr(lk, k);
        look[k] = a != null ? a === '1' || a === 'true' : !!(v & bit);
      }
    }
  }
  styleId ||= info.defaultTable;
  const ts = (styleId && resolveTableStyle(info.styles, styleId)) || { borders: {}, cellBorders: {}, cond: {}, rPr: {}, cellMar: {} };
  const tb = { ...ts.borders, ...direct };
  const mar = { ...ts.cellMar, ...cellMar };
  // First each cell's own borders, shading and text formatting.
  const resolved = cells.map((c) => {
    const lastRow = c.row + c.rowspan === nRows;
    const lastCol = c.col + c.colspan === nCols;
    const edge = {
      top: c.row === 0 ? tb.top : tb.insideH,
      bottom: lastRow ? tb.bottom : tb.insideH,
      left: c.col === 0 ? tb.left : tb.insideV,
      right: lastCol ? tb.right : tb.insideV,
    };
    let shading = ts.shading;
    let rPr = {};
    const cellPad = { ...mar };
    // Conditional regions, lowest priority first.
    const hBand = look.firstRow ? c.row - 1 : c.row;
    const vBand = look.firstColumn ? c.col - 1 : c.col;
    const regions = [];
    if (!look.noVBand && vBand >= 0) regions.push([vBand % 2 ? 'band2Vert' : 'band1Vert', 'col']);
    if (!look.noHBand && hBand >= 0 && !(look.lastRow && lastRow)) regions.push([hBand % 2 ? 'band2Horz' : 'band1Horz', 'row']);
    if (look.firstColumn && c.col === 0) regions.push(['firstCol', 'col']);
    if (look.lastColumn && lastCol) regions.push(['lastCol', 'col']);
    if (look.firstRow && c.row === 0) regions.push(['firstRow', 'row']);
    if (look.lastRow && lastRow) regions.push(['lastRow', 'row']);
    // The corner cells, over both their row's and their column's formatting (Word's order: ne, nw, se, sw).
    const top = c.row === 0, left = c.col === 0;
    if (look.firstRow && look.lastColumn && top && lastCol) regions.push(['neCell', 'cell']);
    if (look.firstRow && look.firstColumn && top && left) regions.push(['nwCell', 'cell']);
    if (look.lastRow && look.lastColumn && lastRow && lastCol) regions.push(['seCell', 'cell']);
    if (look.lastRow && look.firstColumn && lastRow && left) regions.push(['swCell', 'cell']);
    for (const [name, kind] of regions) {
      const cond = ts.cond[name];
      if (!cond) continue;
      const b = { ...cond.borders, ...cond.cellBorders };
      if (kind === 'cell') {
        for (const side of ['top', 'bottom', 'left', 'right']) if (b[side]) edge[side] = b[side];
      } else if (kind === 'row') {
        if (b.top) edge.top = b.top;
        if (b.bottom) edge.bottom = b.bottom;
        if (c.col === 0 ? b.left : b.insideV) edge.left = c.col === 0 ? b.left : b.insideV;
        if (lastCol ? b.right : b.insideV) edge.right = lastCol ? b.right : b.insideV;
      } else {
        if (b.left) edge.left = b.left;
        if (b.right) edge.right = b.right;
        if (c.row === 0 ? b.top : b.insideH) edge.top = c.row === 0 ? b.top : b.insideH;
        if (lastRow ? b.bottom : b.insideH) edge.bottom = lastRow ? b.bottom : b.insideH;
      }
      if (cond.shading) shading = cond.shading;
      rPr = { ...rPr, ...cond.rPr };
    }
    // Direct cell formatting wins.
    let vAlign = null;
    let dir = null;
    for (const [name, xml] of c.tcPr) {
      if (name === 'vAlign') vAlign = { center: 'middle', bottom: 'bottom' }[(/w:val="(\w+)"/.exec(xml) || [])[1]] || null;
      // Text running up (btLr) or down (tbRl) the cell: the cell is as tall as its text is long, not one letter a line.
      if (name === 'textDirection') dir = { btLr: 'sideways-lr', tbRl: 'vertical-rl', tbRlV: 'vertical-rl' }[(/w:val="(\w+)"/.exec(xml) || [])[1]] || null;
      if (name !== 'tcBorders' && name !== 'shd' && name !== 'tcMar') continue;
      const el = parseXml(xml.replace(/^<w:(\w+)/, '<w:$1 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"')).documentElement;
      if (name === 'shd') shading = shdFill(el) ?? shading;
      else if (name === 'tcMar') Object.assign(cellPad, readCellMar(el));
      else Object.assign(edge, readBorders(el));
    }
    return { c, lastRow, lastCol, edge, shading, rPr, cellPad, vAlign, dir };
  });
  // A line between two cells that only one of them sets is still there for both (Word draws the border between them
  // from either side): each draws half of it, and each row grows by its share, so rows measure as in Word.
  const drawn = (b) => b && b.val && !['nil', 'none'].includes(b.val);
  const overlaps = (a0, a1, b0, b1) => a0 < b1 && b0 < a1;
  const byRow = new Map(); // row -> cells starting or ending next to it
  const add = (k, r) => { if (!byRow.has(k)) byRow.set(k, []); byRow.get(k).push(r); };
  for (const r of resolved) for (let k = r.c.row - 1; k <= r.c.row + r.c.rowspan; k++) add(k, r);
  for (const r of resolved) {
    const { c } = r;
    const near = new Set();
    for (let k = c.row; k <= c.row + c.rowspan; k++) for (const n of byRow.get(k) || []) near.add(n);
    for (const n of near) {
      if (n === r) continue;
      const m = n.c;
      if (m.row + m.rowspan === c.row && overlaps(m.col, m.col + m.colspan, c.col, c.col + c.colspan) && !drawn(r.edge.top) && drawn(n.edge.bottom)) r.edge.top = n.edge.bottom;
      if (c.row + c.rowspan === m.row && overlaps(m.col, m.col + m.colspan, c.col, c.col + c.colspan) && !drawn(r.edge.bottom) && drawn(n.edge.top)) r.edge.bottom = n.edge.top;
      if (m.col + m.colspan === c.col && overlaps(m.row, m.row + m.rowspan, c.row, c.row + c.rowspan) && !drawn(r.edge.left) && drawn(n.edge.right)) r.edge.left = n.edge.right;
      if (c.col + c.colspan === m.col && overlaps(m.row, m.row + m.rowspan, c.row, c.row + c.rowspan) && !drawn(r.edge.right) && drawn(n.edge.left)) r.edge.right = n.edge.left;
    }
  }
  // A row's cells share its top and bottom margins in Word: the largest any of them sets (sample S091's middle column,
  // with none of its own, sits 80tw down like the cell beside it).
  const rowTop = new Map();
  const rowBottom = new Map();
  for (const { c, cellPad } of resolved) {
    if (cellPad.top != null) rowTop.set(c.row, Math.max(rowTop.get(c.row) ?? 0, cellPad.top));
    const last = c.row + c.rowspan - 1;
    if (cellPad.bottom != null) rowBottom.set(last, Math.max(rowBottom.get(last) ?? 0, cellPad.bottom));
  }
  for (const { c, cellPad } of resolved) {
    if (rowTop.has(c.row)) cellPad.top = Math.max(cellPad.top ?? 0, rowTop.get(c.row));
    if (rowBottom.has(c.row + c.rowspan - 1)) cellPad.bottom = Math.max(cellPad.bottom ?? 0, rowBottom.get(c.row + c.rowspan - 1));
  }
  const out = [];
  for (const { c, lastRow, lastCol, edge, shading, rPr, cellPad, vAlign, dir } of resolved) {
    // Plain lines are drawn as inset shadows rather than CSS borders: browsers
    // round border widths up to a whole pixel, which made every row of a
    // hairline table taller than in Word. Shared edges are drawn half by each
    // cell; the padding grows by the line's true width so rows measure as in Word.
    const css = [];
    const shadows = [];
    const grow = { top: 0, bottom: 0, left: 0, right: 0 };
    const outer = { top: c.row === 0, bottom: lastRow, left: c.col === 0, right: lastCol };
    // A line between two rows is all the lower row's in Word: its top border pushes its text down, and the upper row ends
    // where the line starts (sample S109's 1.5pt rule under the funding table's header row). Side lines are shared half and half.
    const below = !lastRow && resolved.some((n) => n.c.row === c.row + c.rowspan && overlaps(n.c.col, n.c.col + n.c.colspan, c.col, c.col + c.colspan) && drawn(n.edge.top));
    const DIR = { top: [0, 1], bottom: [0, -1], left: [1, 0], right: [-1, 0] };
    const r3 = (v) => +v.toFixed(3);
    for (const side of ['top', 'bottom', 'left', 'right']) {
      const b = edge[side];
      const solid = b && b.val && !['nil', 'none'].includes(b.val) && borderCss(b).includes(' solid ');
      if (!solid) {
        css.push(`border-${side}:${b ? borderCss(b) : 'none'}`);
        if ((side === 'top' || side === 'bottom') && drawn(b) && b.space > 0) grow[side] += (b.space * 4) / 3;
        continue;
      }
      css.push(`border-${side}:none`);
      const pt = Math.max(0.25, b.sz / 8);
      if (side === 'bottom' && below) { css.push(`--odx-bdb:${+(Math.max(0.25, b.sz / 8) * 4 / 3).toFixed(3)}px`); continue; } // a row at a page's foot still needs room for its line
      const share = outer[side] || side === 'top' ? 1 : 0.5;
      const px = (Math.max(pt * 4 / 3, 1) * share).toFixed(3); // drawn at least 1px wide in total
      const [x, y] = DIR[side];
      shadows.push(`inset ${x * px}px ${y * px}px 0 0 ${b.color && b.color !== 'auto' ? '#' + b.color : '#000'}`);
      grow[side] = (pt * 4) / 3 * share;
      // A border's space (w:space, points) lies between it and the cell's text in Word (doc 01's 2pt double frame: the
      // nested table 2.7px lower than ours).
      if ((side === 'top' || side === 'bottom') && b.space > 0) grow[side] += (b.space * 4) / 3;
    }
    if (shadows.length) css.push(`box-shadow:${shadows.join(',')}`);
    if (shading) css.push(`background-color:${shading}`);
    if (vAlign) css.push(`vertical-align:${vAlign}`);
    if (dir) css.push(`writing-mode:${dir}`);
    const pad = (side) => (cellPad[side] ?? (side === 'left' || side === 'right' ? 7 : 0)) + (side === 'top' || side === 'bottom' ? grow[side] : 0);
    if (cellPad.left != null) css.push(`padding-left:${r3(pad('left'))}px`);
    if (cellPad.right != null) css.push(`padding-right:${r3(pad('right'))}px`);
    css.push(`--odx-pt:${r3(pad('top'))}px;padding-top:${r3(pad('top'))}px`); // pagination adds to --odx-pt
    css.push(`--odx-pb:${r3(pad('bottom'))}px;padding-bottom:${r3(pad('bottom'))}px`);
    const r = rPrToCss(rPr, theme);
    if (r) css.push(r);
    out.push(css.join(';'));
  }
  return { styleId, css: out };
}
