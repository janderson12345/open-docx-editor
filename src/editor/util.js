// Small helpers the editor's modules share.

/** Now, as Word writes a revision's or comment's date (w:date): ISO 8601, to the second, in UTC. */
export const now = () => new Date().toISOString().replace(/\.\d+Z$/, 'Z');

/** The w:id in an atom node's raw XML (a comment or note reference), or undefined. */
export const xmlId = (node) => (/w:id="([^"]*)"/.exec(node.attrs.xml) || [])[1];

/**
 * A paragraph's line height as getComputedStyle's lineHeight gives it ('15.33px' or 'normal'): document paragraphs keep
 * theirs in --odx-lhv, their own line-height being 0 (textspans.js). `cs` is the paragraph's computed style.
 */
export function lineHeightOf(cs) {
  const v = cs.getPropertyValue('--odx-lhv').trim();
  if (!v) return cs.lineHeight;
  if (v === 'normal' || v.endsWith('px')) return v;
  const n = Number(v);
  return Number.isFinite(n) ? `${n * parseFloat(cs.fontSize)}px` : cs.lineHeight;
}
