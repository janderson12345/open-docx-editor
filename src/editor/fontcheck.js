// Fonts a document asks for that the browser can't draw: the host may want to tell the user that lines may break
// differently from Word's. (document.fonts.check() says yes for a font that isn't there at all, so availability is
// found the usual way: text set in the font is measured against the same text in two generic families.)

const SAMPLE = 'mmmmmmmmmmlli1WQ@#';
let ctx = null;
const known = new Map();

/**
 * The page's fonts loaded, or `ms` gone by. (WebKit can leave document.fonts.ready waiting for good after a face fails
 * to load; nothing here should wait on it for ever.)
 */
export const fontsSettled = (ms = 3000) => Promise.race([document.fonts?.ready, new Promise((r) => setTimeout(r, ms))]);

/** Whether the browser has the font `family` (installed, or by an @font-face rule, loaded first). */
export async function hasFont(family) {
  if (known.has(family)) return known.get(family);
  try { await Promise.race([document.fonts.load(`40px "${family.replace(/"/g, '')}"`, SAMPLE), new Promise((r) => setTimeout(r, 3000))]); } catch { /* no such face: measured below */ }
  if (!ctx) ctx = document.createElement('canvas').getContext('2d');
  const width = (stack) => { ctx.font = `40px ${stack}`; return ctx.measureText(SAMPLE).width; };
  const quoted = `"${family.replace(/"/g, '')}"`;
  const ok = ['monospace', 'serif'].some((g) => width(`${quoted}, ${g}`) !== width(g));
  known.set(family, ok);
  return ok;
}

const GENERIC = /^(serif|sans-serif|monospace|cursive|fantasy|system-ui|ui-\w+|emoji|math|fangsong)$/i;
const unquote = (s) => s.trim().replace(/^["']|["']$/g, '');

/**
 * The fonts the document's text asks for that aren't available, each with the font drawn instead (the next one in
 * its stack that is, or null for the browser's generic family): [{ font, substitute, standIn }]. `standIn`: drawn in
 * one of the editor's own stand-ins, another font sized to this one's widths (substitute then null). `root` is the
 * document's content element.
 */
export async function missingFonts(root) {
  const stacks = new Set();
  for (const el of root.querySelectorAll('.odx-t')) stacks.add(getComputedStyle(el).fontFamily);
  const out = new Map();
  for (const stack of stacks) {
    const names = stack.split(',').map(unquote).filter((n) => n && !GENERIC.test(n));
    const wanted = names.find((n) => !n.startsWith('odx-'));
    if (!wanted || out.has(wanted) || await hasFont(wanted)) continue;
    let found = { substitute: null, standIn: false };
    for (const n of names.slice(names.indexOf(wanted) + 1)) {
      if (!(await hasFont(n))) continue;
      // (The editor's own stand-ins, odx-aptos and the like: another font sized to this one's widths.)
      found = n.startsWith('odx-') ? { substitute: null, standIn: true } : { substitute: n, standIn: false };
      break;
    }
    out.set(wanted, found);
  }
  return [...out].map(([font, found]) => ({ font, ...found }));
}
