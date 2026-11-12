// The stand-in fonts (src/fonts: fonts.css and its files/), loaded from jsDelivr (this release's files on npm) or from
// where the host keeps its own copy, and the host's own substitutes for fonts a computer lacks. The browser downloads a
// font only when a document uses it, and uses an installed one first.
import { VERSION } from '../version.js';
import { setFontSubstitutes } from '../ooxml/props.js';

/** Where the stand-in fonts load from when the host doesn't say: this release's own files, on jsDelivr. */
export const FONTS_CDN = `https://cdn.jsdelivr.net/npm/open-docx-editor@${VERSION}/src/fonts/`;

const quote = (name) => `"${String(name).replace(/["\\\r\n\f{}<>;]/g, '')}"`;

/**
 * Loads the stand-ins from `base` (a folder holding fonts.css and files/; false: the host loads fonts itself, e.g.
 * imports open-docx-editor/fonts.css) and sets `substitutes`: { 'Missing Font': 'Family' | { family, url } }, each
 * drawn where that font is missing (a url: a font file the editor loads for it).
 */
export function useFonts({ base = FONTS_CDN, substitutes = null } = {}) {
  if (typeof document === 'undefined') return;
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
