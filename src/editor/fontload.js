// The stand-in fonts (src/fonts: fonts.css and its files/), loaded from jsDelivr (this release's files on npm) or from
// where the host keeps its own copy, and the host's own substitutes for fonts a computer lacks. The browser downloads a
// font only when a document uses it, and uses an installed one first.
import { VERSION } from '../version.js';
import { setFontSubstitutes, setOmf } from '../ooxml/props.js';
import { OMF, OMF_RANGES, OMF_VERSION } from '../fonts/omf.js';

/** Where the stand-in fonts load from when the host doesn't say: this release's own files, on jsDelivr. */
export const FONTS_CDN = `https://cdn.jsdelivr.net/npm/open-docx-editor@${VERSION}/src/fonts/`;
/** Where the open-metric-fonts packages load from when the host doesn't say: npm's @open-metric-fonts, on jsDelivr. */
export const OMF_CDN = 'https://cdn.jsdelivr.net/npm/@open-metric-fonts/';

const quote = (name) => `"${String(name).replace(/["\\\r\n\f{}<>;]/g, '')}"`;

/**
 * Loads the stand-ins from `base` (a folder holding fonts.css and files/; false: the host loads fonts itself, e.g.
 * imports open-docx-editor/fonts.css) and sets `substitutes`: { 'Missing Font': 'Family' | { family, url } }, each
 * drawn where that font is missing (a url: a font file the editor loads for it).
 */
export function useFonts({ base = FONTS_CDN, substitutes = null, omf = OMF_CDN } = {}) {
  if (typeof document === 'undefined') return;
  setOmf(omf ? omfLoader(omf) : null);
  if (base) {
    const href = new URL('fonts.css', new URL(base.endsWith('/') ? base : `${base}/`, document.baseURI)).href;
    if (![...document.querySelectorAll('link[data-odx-fonts]')].some((l) => l.href === href)) {
      const link = document.createElement('link');
      link.rel = 'stylesheet';
      link.href = href;
      link.dataset.odxFonts = '';
      // (crossOrigin: a CDN's fonts load as CORS requests; this keeps the stylesheet's the same.)
      link.crossOrigin = 'anonymous';
      document.head.appendChild(link);
    }
  }
  if (!substitutes) return;
  const map = new Map();
  const faces = [];
  for (const [name, sub] of Object.entries(substitutes)) {
    const family = typeof sub === 'string' ? sub : sub?.family;
    if (!family) continue;
    map.set(name, family);
    if (typeof sub === 'object' && sub.url) faces.push(`@font-face { font-family: ${quote(family)}; src: url(${JSON.stringify(String(sub.url))}); font-display: swap; }`);
  }
  setFontSubstitutes(map);
  if (faces.length) {
    let style = document.querySelector('style[data-odx-font-substitutes]');
    if (!style) { style = document.createElement('style'); style.dataset.odxFontSubstitutes = ''; document.head.appendChild(style); }
    style.textContent = faces.join('\n');
  }
}

/**
 * The open-metric-fonts stand-ins (src/fonts/omf.js), for a font a document names: the family the editor draws it in
 * where the computer lacks it ("odx-omf <name>"), its faces registered the first time a document asks (one @font-face
 * per face and alphabet: the browser downloads a file only for the characters it draws). From `where`: OMF_CDN (the
 * pinned release on jsDelivr), or a folder of the host's own copy holding each package as npm has it,
 * <where>/<package>/files/<face>-<alphabet>.woff2 (node_modules/@open-metric-fonts, or a release zip's web/ folders).
 */
function omfLoader(where) {
  const cdn = where === OMF_CDN;
  const root = cdn ? null : new URL(where.endsWith('/') ? where : `${where}/`, document.baseURI).href;
  const folder = (pkg) => (cdn ? `${OMF_CDN}${pkg}@${OMF_VERSION}/files/` : `${root}${pkg}/files/`);
  const done = new Set();
  return (name) => {
    const entry = OMF[name];
    if (!entry) return null;
    const family = `odx-omf ${name}`;
    if (!done.has(name)) {
      done.add(name);
      const [pkg, , faces] = entry;
      const rules = [];
      for (const [weight, italic, stem, alphabets] of faces) {
        for (const a of alphabets.split(',')) {
          rules.push(`@font-face { font-family: ${quote(family)}; font-style: ${italic ? 'italic' : 'normal'}; font-weight: ${weight}; font-display: swap; `
            + `src: url(${JSON.stringify(`${folder(pkg)}${stem}-${a}.woff2`)}) format('woff2'); unicode-range: ${OMF_RANGES[`${pkg}:${a}`] || OMF_RANGES[a]}; }`);
        }
      }
      let style = document.querySelector('style[data-odx-omf]');
      if (!style) { style = document.createElement('style'); style.dataset.odxOmf = ''; document.head.appendChild(style); }
      style.appendChild(document.createTextNode(`${rules.join('\n')}\n`));
    }
    return family;
  };
}
