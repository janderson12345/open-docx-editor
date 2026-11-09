// Word equations (Office Math, m:oMath / m:oMathPara) drawn as MathML, which browsers lay out natively: fractions,
// scripts, roots, n-ary operators (sums, integrals), delimiters, accents, bars, limits, functions, matrices and
// equation arrays. Display only: the equation's own XML is what the document keeps and saves.
import { parseXml, esc, elementKids as kidsOf, childNamed as child, childrenNamed } from './xml.js';

const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

/** A property's m:val (or a flag's presence: <m:degHide/> is on). */
const prop = (el, prName, name) => {
  const p = child(child(el, prName), name);
  if (!p) return null;
  if (p.hasAttributeNS && p.hasAttributeNS(M, 'val')) return p.getAttributeNS(M, 'val');
  return p.getAttribute('m:val') ?? '';
};
const on = (v) => v != null && v !== '0' && v !== 'off' && v !== 'false';

const mo = (c, extra = '') => `<mo${extra}>${esc(c)}</mo>`;
const row = (xs) => (xs.length === 1 ? xs[0] : `<mrow>${xs.join('')}</mrow>`);
// A run gives its tokens (an array), every other element one MathML element (a string): an argument's runs and
// elements go into one row together, so an operator between them is an infix one, spaced as such (b² − 4ac).
const flat = (els) => els.flatMap((n) => { const v = node(n); return Array.isArray(v) ? v : v ? [v] : []; });
/** An argument (m:e, m:num, m:sub...) as one MathML element. */
const arg = (el) => row(el ? flat(kidsOf(el)) : []) || '<mrow></mrow>';

const OPS = /[+\-=<>±∓×÷·∙∘≠≤≥≈≡∼∝∞∂∇∈∉⊂⊃⊆⊇∪∩∧∨¬→←↔⇒⇐⇔∀∃,;:!|/()[\]{}⟨⟩∑∏∫∮−]/;

/** A run's text as tokens: numbers (mn), letters (mi, one each, italic), the rest operators (mo). */
function run(r) {
  const text = childrenNamed(r, 't').map((t) => t.textContent).join('');
  if (!text) return [];
  const sty = prop(r, 'rPr', 'sty');
  const plain = on(prop(r, 'rPr', 'nor')) || sty === 'p' || sty === 'b';
  const bold = sty === 'b' || sty === 'bi';
  const variant = plain ? (bold ? ' mathvariant="bold"' : ' mathvariant="normal"') : bold ? ' mathvariant="bold-italic"' : '';
  // Plain text (a function name, a word): one identifier, upright.
  if (plain && /^[\p{L}\s]+$/u.test(text)) return [`<mi${variant}>${esc(text)}</mi>`];
  const out = [];
  for (const m of text.matchAll(/(\d+(?:[.,]\d+)*)|(\p{L})|(\s+)|(.)/gu)) {
    if (m[1]) out.push(`<mn>${m[1]}</mn>`);
    else if (m[2]) out.push(`<mi${variant}>${esc(m[2])}</mi>`);
    else if (m[3]) out.push('<mspace width="0.25em"></mspace>');
    // (A hyphen is a minus sign in an equation, as Word draws it.)
    else out.push(m[4] === '-' ? mo('−') : OPS.test(m[4]) ? mo(m[4]) : `<mi${variant}>${esc(m[4])}</mi>`);
  }
  return out;
}

function node(el) {
  switch (el.localName) {
    case 'r': return run(el);
    case 'oMath': return arg(el);
    case 'f': {
      const type = prop(el, 'fPr', 'type');
      const num = arg(child(el, 'num'));
      const den = arg(child(el, 'den'));
      if (type === 'lin') return `<mrow>${num}${mo('/')}${den}</mrow>`;
      return `<mfrac${type === 'noBar' ? ' linethickness="0"' : type === 'skw' ? ' bevelled="true"' : ''}>${num}${den}</mfrac>`;
    }
    case 'sSup': return `<msup>${arg(child(el, 'e'))}${arg(child(el, 'sup'))}</msup>`;
    case 'sSub': return `<msub>${arg(child(el, 'e'))}${arg(child(el, 'sub'))}</msub>`;
    case 'sSubSup': return `<msubsup>${arg(child(el, 'e'))}${arg(child(el, 'sub'))}${arg(child(el, 'sup'))}</msubsup>`;
    case 'sPre': return `<mmultiscripts>${arg(child(el, 'e'))}<mprescripts></mprescripts>${arg(child(el, 'sub'))}${arg(child(el, 'sup'))}</mmultiscripts>`;
    case 'rad': {
      const deg = child(el, 'deg');
      const hide = on(prop(el, 'radPr', 'degHide')) || !deg || !kidsOf(deg).length;
      return hide ? `<msqrt>${arg(child(el, 'e'))}</msqrt>` : `<mroot>${arg(child(el, 'e'))}${arg(deg)}</mroot>`;
    }
    case 'nary': {
      const chr = prop(el, 'naryPr', 'chr') || '∫';
      const op = mo(chr, ' largeop="true" movablelimits="false"');
      const sub = on(prop(el, 'naryPr', 'subHide')) ? null : child(el, 'sub');
      const sup = on(prop(el, 'naryPr', 'supHide')) ? null : child(el, 'sup');
      // Limits over and under the sign (undOvr), else as scripts beside it (subSup); Word's default puts an integral's
      // beside it, a sum's or product's over and under.
      const loc = prop(el, 'naryPr', 'limLoc') || (/[∫∬∭∮∯∰]/.test(chr) ? 'subSup' : 'undOvr');
      const [both, under, over] = loc === 'undOvr' ? ['munderover', 'munder', 'mover'] : ['msubsup', 'msub', 'msup'];
      const base = sub && sup ? `<${both}>${op}${arg(sub)}${arg(sup)}</${both}>` : sub ? `<${under}>${op}${arg(sub)}</${under}>` : sup ? `<${over}>${op}${arg(sup)}</${over}>` : op;
      return `<mrow>${base}${arg(child(el, 'e'))}</mrow>`;
    }
    case 'd': {
      const beg = prop(el, 'dPr', 'begChr') ?? '(';
      const end = prop(el, 'dPr', 'endChr') ?? ')';
      const sep = prop(el, 'dPr', 'sepChr') ?? '|';
      const parts = childrenNamed(el, 'e').map(arg);
      const inner = parts.join(sep ? mo(sep, ' separator="true"') : '');
      return `<mrow>${beg ? mo(beg, ' fence="true" stretchy="true"') : ''}${inner}${end ? mo(end, ' fence="true" stretchy="true"') : ''}</mrow>`;
    }
    case 'acc': return `<mover accent="true">${arg(child(el, 'e'))}${mo(prop(el, 'accPr', 'chr') || '̂', ' stretchy="true"')}</mover>`;
    case 'bar': return prop(el, 'barPr', 'pos') === 'top'
      ? `<mover accent="true">${arg(child(el, 'e'))}${mo('¯', ' stretchy="true"')}</mover>`
      : `<munder accentunder="true">${arg(child(el, 'e'))}${mo('_', ' stretchy="true"')}</munder>`;
    case 'groupChr': {
      const chr = prop(el, 'groupChrPr', 'chr') || '⏟';
      const top = prop(el, 'groupChrPr', 'pos') === 'top';
      return top ? `<mover>${arg(child(el, 'e'))}${mo(chr, ' stretchy="true"')}</mover>` : `<munder>${arg(child(el, 'e'))}${mo(chr, ' stretchy="true"')}</munder>`;
    }
    case 'limLow': return `<munder>${arg(child(el, 'e'))}${arg(child(el, 'lim'))}</munder>`;
    case 'limUpp': return `<mover>${arg(child(el, 'e'))}${arg(child(el, 'lim'))}</mover>`;
    // (A thin space after the name, as Word sets "sin θ": the invisible function application takes none in browsers.)
    case 'func': return `<mrow>${arg(child(el, 'fName'))}<mo>⁡</mo><mspace width="0.17em"></mspace>${arg(child(el, 'e'))}</mrow>`;
    case 'm': return `<mtable>${childrenNamed(el, 'mr').map((r) => `<mtr>${childrenNamed(r, 'e').map((e) => `<mtd>${arg(e)}</mtd>`).join('')}</mtr>`).join('')}</mtable>`;
    case 'eqArr': return `<mtable columnalign="left">${childrenNamed(el, 'e').map((e) => `<mtr><mtd>${arg(e)}</mtd></mtr>`).join('')}</mtable>`;
    case 'box': return arg(child(el, 'e'));
    case 'borderBox': return `<menclose notation="box">${arg(child(el, 'e'))}</menclose>`;
    case 'phant': return on(prop(el, 'phantPr', 'show') ?? '1') ? arg(child(el, 'e')) : `<mphantom>${arg(child(el, 'e'))}</mphantom>`;
    // Properties, control characters and Word's own run properties draw nothing.
    case 'rPr': case 'ctrlPr': case 'oMathParaPr': case 'argPr':
      return '';
    default:
      if (/Pr$/.test(el.localName)) return '';
      return flat(kidsOf(el)).join('');
  }
}

/** MathML for an m:oMath or m:oMathPara element's XML, or null when it can't be read. */
export function ommlToMathML(xml) {
  // (Kept as it stood in the document, its prefixes declared on the document's root: declared again around it.)
  let el;
  try {
    const root = parseXml(`<x xmlns:m="${M}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml}</x>`).documentElement;
    el = kidsOf(root)[0];
  } catch { return null; }
  if (!el) return null;
  if (el.localName === 'oMathPara') {
    return childrenNamed(el, 'oMath').map((m) => `<math display="block">${arg(m)}</math>`).join('');
  }
  if (el.localName === 'oMath') return `<math>${arg(el)}</math>`;
  return null;
}
