// Paragraphs in frames (w:framePr, other than drop caps), placed as Word places them: a frame floats at its place across
// the page (x or xAlign from the page, the margins or the text), its width w, and the text goes beside it (wrap around)
// or not (notBeside: above and below; none: under it). One anchored to the text sits y below where it comes; one
// anchored to the page or the margins is set down y from their top on the page where it comes (pages.js works out how
// far that is: --odx-fdy). Paragraphs in a row with the same w:framePr make one frame, one below the other.
import { twipsToPx } from '../ooxml/units.js';

const num = (xml, a) => { const m = new RegExp(`\\sw:${a}="(-?\\d+)"`).exec(xml); return m ? Number(m[1]) : null; };
const str = (xml, a) => new RegExp(`\\sw:${a}="([^"]*)"`).exec(xml)?.[1] ?? null;

/** A paragraph's frame (its w:framePr XML), or null; drop caps aren't frames here (editor.js draws them). */
export function frameOf(node) {
  const f = (node?.attrs.extra || []).find((e) => e[0] === 'framePr')?.[1];
  return f && !/w:dropCap="(?:drop|margin)"/.test(f) ? f : null;
}

/**
 * CSS and attributes for a framed paragraph on `page` (px: width, left, right): { style: [...], attrs: {...} }.
 * `same`: the paragraph before it is in the same frame (it goes below that one).
 */
export function frameStyle(xml, same, page) {
  const tw = (a) => twipsToPx(num(xml, a) || 0);
  const w = num(xml, 'w') ? twipsToPx(num(xml, 'w')) : null;
  const h = num(xml, 'h') ? twipsToPx(num(xml, 'h')) : null;
  const hRule = str(xml, 'hRule');
  const wrap = str(xml, 'wrap') || 'around';
  const hA = str(xml, 'hAnchor') || 'page';
  const vA = str(xml, 'vAnchor') || 'page';
  const hs = tw('hSpace'), vs = tw('vSpace');
  const cw = page.width - page.left - page.right;
  // Across: from the text's left edge. The area the frame is placed in: the page, or the margins (and the text).
  const [from, to] = hA === 'page' ? [-page.left, cw + page.right] : [0, cw];
  const xAlign = str(xml, 'xAlign');
  let side = 'left', left = 0, right = 0;
  if (xAlign === 'right' || xAlign === 'outside') { side = 'right'; right = cw - to; }
  else if (xAlign === 'center' && w != null) left = from + (to - from - w) / 2;
  else if (xAlign === 'left' || xAlign === 'inside' || xAlign === 'center') left = from;
  else {
    left = from + tw('x');
    // (The text goes on the side with more room: a frame right of the middle floats right.)
    if (w != null && left + w / 2 > cw / 2) { side = 'right'; right = cw - (left + w); }
  }
  const style = [];
  const attrs = {};
  if (w != null) style.push(`width:${w}px`, 'box-sizing:content-box');
  if (h != null && hRule === 'exact') style.push(`height:${h}px`, 'overflow:hidden');
  else if (h != null) style.push(`min-height:${h}px`);
  if (wrap === 'notBeside' || wrap === 'topAndBottom') {
    // Above and below it only: in the flow, at its place across.
    if (side === 'right') style.push('margin-left:auto', `margin-right:${right}px`);
    else style.push(`margin-left:${left}px`);
    return { style, attrs };
  }
  style.push(`float:${side}`, 'position:relative', 'z-index:1');
  // (x is where its text starts: its border and the border's space go outside that, as Word draws them.)
  style.push(side === 'left' ? `margin-left:calc(${left}px - var(--odx-bdl, 0pt));margin-right:${hs}px` : `margin-right:calc(${right}px - var(--odx-bdr, 0pt));margin-left:${hs}px`);
  if (vs) style.push(`margin-bottom:${vs}px`); // (else its own space after)
  if (same) {
    style.push(`clear:${side}`); // (its own space before stays: floats' margins don't collapse, nor do Word's)
    attrs['data-frame'] = 'cont';
  } else {
    // Down: below where it comes (anchored to the text), or where pages.js sets it down on its page.
    const dy = vA === 'text' ? `${tw('y')}px` : 'var(--odx-fdy, 0px)';
    style.push(`margin-top:${dy}`);
    // Beside the room above it the text runs on (the margin moving it down isn't part of it); under it none (wrap none).
    style.push(wrap === 'none' || wrap === 'through' ? 'shape-outside:polygon(0 0, 0 0, 0 0)' : `shape-outside:inset(max(0px, calc(${dy} - ${vs}px)) 0 0 0)`);
    attrs['data-frame'] = vA === 'text' ? 'text' : vA;
    if (vA !== 'text') {
      attrs['data-fy'] = String(tw('y'));
      const yAlign = str(xml, 'yAlign');
      if (yAlign && yAlign !== 'inline') attrs['data-fyalign'] = yAlign;
    }
  }
  if (wrap === 'none' || wrap === 'through') style.push('shape-outside:polygon(0 0, 0 0, 0 0)');
  return { style, attrs };
}
