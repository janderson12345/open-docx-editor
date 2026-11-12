// Print relayer (engine stage 3, first cut): in print, each line of each top-level paragraph is drawn as its own element
// at the exact (fractional) position the layout gives it, plus the paragraph's print baseline shift. Edge rounds a
// block's printed position to a whole pixel, and the text of every line inside it too (measured on printed test grids), but
// not a transform: a line of its own, moved by a transform, prints where it's laid out. Line breaks are still the
// browser's; only where each line prints changes. The original paragraphs stay (hidden) so pages keep their layout.
// On by default for printing; off with window.odxEngine = false (or ?noengine in the page's URL).

const LAYER = 'odx-printlines';

/** Text nodes of `el` in order, skipping pagination spacers and non-editable widgets' labels. */
function texts(el) {
  const out = [];
  const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, {
    acceptNode: (n) => (!n.data.length || n.parentElement.closest('.odx-push') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  for (let n = w.nextNode(); n; n = w.nextNode()) out.push(n);
  return out;
}

/** Whether a text node sits in an inline box of its own (a tab, a widget): its rect isn't the line's text top. */
const boxed = (n, stop) => {
  for (let e = n.parentElement; e && e !== stop; e = e.parentElement) if (!/^(inline|contents)$/.test(getComputedStyle(e).display)) return true;
  return false;
};
/** Whether a text node isn't drawn at all (hidden deleted text, collapsed markup). */
const unseen = (n, stop) => {
  for (let e = n.parentElement; e && e !== stop; e = e.parentElement) if (getComputedStyle(e).display === 'none') return true;
  return false;
};

/**
 * The paragraph's lines: [{ top, start: [node, offset], end: [node, offset] }] from the laid-out text. A line's top is
 * that of its plain text (a tab's box sits elsewhere); text in inline boxes joins the line it's on. null when a line
 * holds no plain text (the paragraph is then left to the normal print).
 */
function lines(el) {
  const r = document.createRange();
  const out = [];
  let cur = null;
  let pending = null; // boxed text before the first plain text of a line
  for (const n of texts(el)) {
    if (unseen(n, el)) continue;
    if (boxed(n, el)) {
      r.selectNodeContents(n);
      const q = r.getBoundingClientRect();
      if (cur && q.top < cur.bottom + 1) cur.end = [n, n.data.length];
      else if (!pending) pending = [n, 0];
      continue;
    }
    // Characters of a text node keep their line until a rect's top moves on: split where it does (binary search).
    let from = 0;
    while (from < n.data.length) {
      r.setStart(n, from); r.setEnd(n, from + 1);
      const first = r.getClientRects()[0];
      if (!first) { from++; continue; }
      let lo = from, hi = n.data.length - 1;
      while (lo < hi) {
        const mid = (lo + hi + 1) >> 1;
        r.setStart(n, mid); r.setEnd(n, mid + 1);
        const q = r.getClientRects()[0];
        if (q && Math.abs(q.top - first.top) < 2) lo = mid; else hi = mid - 1;
      }
      // Same line when the fragments overlap vertically (text of another size sits higher or lower on it), starting above
      // its middle: under exact spacing tighter than the font, the next line's text overlaps this one's too (02-spacing's
      // 12pt-exact Calibri lines, 18px boxes 16px apart, had printed as one unwrapped line).
      if (cur && first.top < cur.bottom - 1 && first.bottom > cur.lo + 1 && first.top < (cur.lo + cur.bottom) / 2) { cur.end = [n, lo + 1]; cur.bottom = Math.max(cur.bottom, first.bottom); cur.lo = Math.min(cur.lo, first.top); }
      else {
        cur = { top: first.top, lo: first.top, bottom: first.bottom, left: first.left, start: pending || [n, from], end: [n, lo + 1] };
        out.push(cur);
      }
      pending = null;
      from = lo + 1;
    }
  }
  return pending ? null : out;
}

/** A ':nth-child' path from `root` (exclusive) down to `el`, for a rule that hides it without touching it. */
function pathOf(el, root) {
  const parts = [];
  for (let e = el; e && e !== root; e = e.parentElement) parts.unshift(`:nth-child(${[...e.parentElement.children].indexOf(e) + 1})`);
  return parts.join(' > ');
}

/**
 * Shallow copies of `el`'s ancestors up to (not including) the first `.odx-content` or `.odx-hf`, nested under `into`
 * and drawn as display:contents: the paragraph's rules that go by its context (a table style's "td > p", a header's)
 * still match its lines and their values still inherit, without a box of their own. Returns the innermost copy.
 */
function context(el, into) {
  const chain = [];
  for (let e = el.parentElement; e && !e.classList.contains('odx-content') && !e.classList.contains('odx-hf') && !e.classList.contains('odx-repeats'); e = e.parentElement) chain.unshift(e);
  let at = into;
  for (const a of chain) {
    const c = a.cloneNode(false);
    c.removeAttribute('id');
    c.style.cssText += ';display:contents';
    at.append(c);
    at = c;
  }
  return at;
}

/** Where a line's first text starts (screen x). */
const lineLeft = (ln) => ln.left;

// (Not a float placed against its page, lifted to the page's first paragraph (pages.js liftFloats), nor its hidden
// original: the layer draws the lifted ones itself, and those paragraphs' lines go round them.)
// (And equations: a line's copy is cut from its text, which would flatten their layout. Their paragraphs print as laid out.)
const SKIP = 'img:not(.ProseMirror-separator):not(.odx-lifted):not(.odx-lift > img), [style*="position: absolute"], .odx-opaque-inline[style*="inline-block"], .odx-math';

/**
 * Builds the print layer under `sheet`: the lines of the document's paragraphs (top level and in table cells) and of
 * the headers and footers, each placed exactly. Returns a function removing it.
 */
export function buildPrintLines(sheet, contentEl, shiftOf, scope) {
  removePrintLines(sheet);
  const layer = document.createElement('div');
  layer.className = LAYER;
  sheet.append(layer);
  // Inside an .odx-content of its own, so the editor's paragraph rules (list labels, line heights) apply to the lines;
  // not a direct child of the sheet, so the rules on the document's own blocks by position don't.
  const host = document.createElement('div');
  host.className = 'odx-content odx-pl-host';
  layer.append(host);
  const origin = sheet.getBoundingClientRect();
  const r = document.createRange();
  // Lifted floats: drawn in the layer where they are on screen (their paragraphs' own boxes are hidden in print), and
  // the lines beside them start and end where their text does.
  const floats = [...contentEl.querySelectorAll('.odx-lift > img')].map((e) => {
    const cs = getComputedStyle(e);
    const fr = e.getBoundingClientRect();
    const copy = e.cloneNode(false);
    copy.style.cssText = `position:absolute;margin:0;left:${(fr.left - origin.left).toFixed(2)}px;top:${(fr.top - origin.top).toFixed(2)}px;width:${fr.width}px;height:${fr.height}px`;
    layer.append(copy);
    return { r: fr, side: cs.float, ml: parseFloat(cs.marginLeft) || 0 };
  });
  // Every other float too (a picture wrapped square in a paragraph above, reaching down beside this one's lines): the
  // lines it pushes over or cuts short on screen start and end where their text does in the copy as well.
  for (const e of contentEl.querySelectorAll('img:not(.ProseMirror-separator)')) {
    if (e.closest('.odx-lift')) continue;
    const cs = getComputedStyle(e);
    if (cs.float === 'none') continue;
    const fr = e.getBoundingClientRect();
    floats.push({ r: { left: fr.left - (parseFloat(cs.marginLeft) || 0), right: fr.right + (parseFloat(cs.marginRight) || 0), top: fr.top - Math.max(0, parseFloat(cs.marginTop) || 0), bottom: fr.bottom + (parseFloat(cs.marginBottom) || 0) }, side: cs.float, ml: 0 });
  }
  // The editor owns its paragraphs' DOM (a class added to one is undone, and the element redrawn): read them all first,
  // and hide the originals by rules on their position instead.
  const paras = [...contentEl.children].filter((e) => e.tagName === 'P');
  for (const td of contentEl.querySelectorAll('td, th')) {
    if (getComputedStyle(td).writingMode !== 'horizontal-tb') continue;
    for (const pp of td.children) if (pp.tagName === 'P') paras.push(pp);
  }
  // Headers and footers, and repeated table header rows (drawn apart from the document, at each page's top).
  const hf = [...sheet.querySelectorAll('.odx-hf p, .odx-repeat p')];
  const hide = [];
  for (const el of [...paras, ...hf]) {
    const why = window.__plWhy; // measurement hook: why paragraphs are left to the normal print
    // (An inline picture inside a line of text goes with that line: the line's copy holds it. One on a line of its own,
    // or anything else of SKIP's, leaves the paragraph to the normal print.)
    const pics = [...el.querySelectorAll(SKIP)];
    const inline = pics.length && pics.every((e) => e.tagName === 'IMG' && !e.closest('.odx-lift') && getComputedStyle(e).float === 'none' && getComputedStyle(e).position === 'static');
    if ((pics.length && !inline) || el.closest('.odx-hf td, .odx-hf th')?.closest('[style*="writing-mode"]')) { if (why) why.push(['skip', pics[0]?.outerHTML.slice(0, 120)]); continue; }
    const ls = lines(el);
    if (!ls || !ls.length) { if (why) why.push([ls ? 'empty' : 'lines', el.textContent.slice(0, 60)]); continue; }
    if (inline && !pics.every((e) => ls.some((ln) => { r.setStart(...ln.start); r.setEnd(...ln.end); return r.intersectsNode(e); }))) { if (why) why.push(['picline', el.textContent.slice(0, 60)]); continue; }
    const cs = getComputedStyle(el);
    const box = el.getBoundingClientRect();
    // An exact row clips its paragraphs: lines past the clip aren't drawn.
    const cell = el.parentElement?.closest('tr[data-exact] > td');
    let clip = cell ? cell.getBoundingClientRect().bottom : cs.overflowY === 'clip' || cs.overflowY === 'hidden' ? box.bottom : Infinity;
    // A paragraph cut by an inset clip (split across columns: the part above and the copy below): only its lines inside.
    let clipTop = -Infinity;
    const inset = /^inset\(([-\d.]+)px [-\d.]+px ([-\d.]+)px/.exec(cs.clipPath || '');
    if (inset) {
      if (Number(inset[1]) > 0) clipTop = box.top + Number(inset[1]);
      if (Number(inset[2]) > 0) clip = Math.min(clip, box.bottom - Number(inset[2]));
    }
    // ... or a table holding it cut so (one split across columns, pages.js colSplitTable, and its copies): only its lines
    // inside the table's part.
    for (let e = el.parentElement; e && e !== contentEl && !e.classList.contains('odx-repeat'); e = e.parentElement) {
      const m = /^inset\(([-\d.]+)px [-\d.]+px ([-\d.]+)px/.exec(getComputedStyle(e).clipPath || '');
      if (!m) continue;
      const eb = e.getBoundingClientRect();
      if (Number(m[1]) > 0) clipTop = Math.max(clipTop, eb.top + Number(m[1]));
      if (Number(m[2]) > 0) clip = Math.min(clip, eb.bottom - Number(m[2]));
    }
    const left = box.left + parseFloat(cs.borderLeftWidth) + parseFloat(cs.paddingLeft) - origin.left;
    const width = box.width - parseFloat(cs.borderLeftWidth) - parseFloat(cs.paddingLeft) - parseFloat(cs.borderRightWidth) - parseFloat(cs.paddingRight);
    // The paragraph's print baseline shift, less what's already applied to it (measured under print styles, it is).
    const d = (shiftOf(el) || 0) - (cs.position === 'relative' ? parseFloat(cs.top) || 0 : 0);
    const justify = cs.textAlign === 'justify';
    const at = context(el, host);
    // A cell clips its text at its sides, as Word does (text pushed out by a negative indent isn't drawn).
    const td = el.parentElement?.closest('td, th');
    const side = td ? td.getBoundingClientRect() : null;
    // Lines beside a floated paragraph (a drop cap just before) start where their text does.
    const floatBefore = el.previousElementSibling && getComputedStyle(el.previousElementSibling).float !== 'none';
    ls.forEach((ln, k) => {
      if (ln.top >= clip - 1 || ln.top < clipTop - 1) return;
      r.setStart(...ln.start); r.setEnd(...ln.end);
      // A wrapped line's trailing spaces hang past its end, as on screen and in Word: in the copy (white-space: pre) they
      // would count, and a right-aligned or justified line would end a space short (07-paragraph-format: 2.8px).
      if (k < ls.length - 1 && ln.end[0].nodeType === 3) {
        let [en, eo] = ln.end;
        while (eo > 0 && en.data[eo - 1] === ' ' && !(en === ln.start[0] && eo - 1 <= ln.start[1])) eo--;
        if (eo !== ln.end[1]) r.setEnd(en, eo);
      }
      const line = el.cloneNode(false);
      line.style.removeProperty('margin-top'); // (a page's push, set on the block (pages.js), not the line's: !important, it beat margin:0 below)
      line.removeAttribute('id');
      // (Not the paragraph's inset clip: its lines were chosen by it above, and on one line it would cut the line's own
      // text, sample S045's column copies printing only their parentheses' descenders.)
      if (inset) line.style.clipPath = '';
      if (k > 0) line.classList.remove('odx-li'); // the list label is drawn on the first line only
      // A line pushed right by a float beside it (a drop cap) starts where its text does.
      const expected = left + (k === 0 ? parseFloat(cs.textIndent) || 0 : 0);
      const over = floats.filter((f) => ln.top < f.r.bottom && ln.bottom > f.r.top);
      // (So does one beside a shape with text on both sides (wrapboth.js) whose first words went past it: its spacer, before
      // the line's first text, isn't in the copy.)
      const beside = (floatBefore || over.some((f) => f.side === 'left') || el.querySelector('.odx-wb')) && !el.classList.contains('odx-li') && lineLeft(ln) - origin.left > expected + 2 ? lineLeft(ln) - origin.left - expected : 0;
      const cutF = Math.max(0, ...over.filter((f) => f.side === 'right').map((f) => left + width - (f.r.left - f.ml - origin.left)));
      line.style.cssText += `;position:absolute;margin:0;padding:0;border:0;top:0;left:${left + beside}px;width:${width - beside - cutF}px;white-space:pre;` +
        `text-indent:${k === 0 ? cs.textIndent : '0px'};text-align-last:${justify && k < ls.length - 1 ? 'justify' : 'auto'};min-height:0;max-height:none;overflow:visible`;
      // The range's contents, inside copies of its run elements: a line within one text node (or one span) would
      // otherwise come without them, losing its bold, font or size in print.
      let frag = r.cloneContents();
      frag.querySelectorAll('.odx-lift').forEach((e) => e.remove()); // (drawn once, by the layer)
      for (let e = r.commonAncestorContainer.nodeType === 1 ? r.commonAncestorContainer : r.commonAncestorContainer.parentNode; e && e !== el; e = e.parentNode) {
        const c = e.cloneNode(false);
        if (c.removeAttribute) c.removeAttribute('id');
        c.append(frag);
        frag = c;
      }
      line.append(frag);
      // A line broken at an automatic hyphenation point (hyphenate.js: the browser's own, or its soft hyphens): on screen
      // the browser draws the hyphen there, but in the copy (white-space: pre) nothing breaks, so the hyphen is written
      // in, in the run's font. (The browser's own: the next line goes on in the same word, the same text node.)
      const shyAt = (n) => !!n?.parentElement?.closest('.odx-shy');
      const next = ls[k + 1]?.start;
      const inWord = next && el.closest('[data-hyph]') && next[0] === ln.end[0] && next[1] === ln.end[1] && ln.end[0].nodeType === 3
        && /\p{L}/u.test(ln.end[0].data[ln.end[1] - 1] || '') && /\p{L}/u.test(ln.end[0].data[ln.end[1]] || '');
      if (k < ls.length - 1 && (shyAt(ln.end[0]) || shyAt(ls[k + 1].start[0]) || inWord)) {
        const own = [...line.querySelectorAll('.odx-shy')].pop();
        if (shyAt(ln.end[0]) && own) own.textContent = '-';
        else { const t = texts(line).pop(); if (t) t.data += '-'; }
      }
      if (side) {
        const lx = left + beside + origin.left;
        const cutL = Math.max(0, side.left - lx), cutR = Math.max(0, lx + width - beside - side.right);
        if (cutL >= width - beside) return;
        if (cutL || cutR) line.style.clipPath = `inset(-50px ${cutR}px -50px ${cutL}px)`;
      }
      at.append(line);
      // Where the clone's text sits in it, then move it so the text lands on the original line's, plus the shift.
      const lt = line.getBoundingClientRect().top;
      const t0 = texts(line).find((n) => !boxed(n, line));
      let off = 0;
      if (t0) { r.selectNodeContents(t0); off = (r.getClientRects()[0]?.top ?? lt) - lt; }
      // Edge rounds the line's text offset inside the clone when printing: compensate with the rounded one.
      const y = ln.top - origin.top - Math.round(off) + d;
      line.style.transform = `translateY(${y.toFixed(3)}px)`;
    });
    // Its borders and shading: the paragraph is hidden in print, so its box is drawn here, under its lines (a framed
    // paragraph's box, a ruled heading's line).
    const edges = ['Top', 'Right', 'Bottom', 'Left'];
    const shaded = !/^(transparent|rgba\(0, 0, 0, 0\))$/.test(cs.backgroundColor || 'transparent');
    if (shaded || edges.some((s) => parseFloat(cs[`border${s}Width`]) > 0 && cs[`border${s}Style`] !== 'none')) {
      const under = document.createElement('div');
      under.style.cssText = `position:absolute;box-sizing:border-box;left:${box.left - origin.left}px;top:${box.top - origin.top}px;width:${box.width}px;height:${box.height}px;`
        + edges.map((s) => `border-${s.toLowerCase()}:${cs[`border${s}Width`]} ${cs[`border${s}Style`]} ${cs[`border${s}Color`]}`).join(';')
        + (shaded ? `;background:${cs.backgroundColor}` : '') + (cs.clipPath && cs.clipPath !== 'none' ? `;clip-path:${cs.clipPath}` : '');
      layer.insertBefore(under, host);
    }
    hide.push(`.${scope} > ${pathOf(el, sheet)}`);
  }
  const style = document.createElement('style');
  // A rule each: one selector the browser can't read would drop a combined rule for all of them.
  style.textContent = hide.map((sel) => `${sel}{visibility:hidden !important}`).join('\n');
  if (hide.length) layer.append(style);
  return () => removePrintLines(sheet);
}

export function removePrintLines(sheet) {
  sheet.querySelectorAll(`:scope > .${LAYER}`).forEach((e) => e.remove());
}

/** On by default (with it 91.8% of lines were within 1px of Word, without 90.5%); off with window.odxEngine = false or ?noengine. */
export const engineOn = () => typeof window !== 'undefined' && window.odxEngine !== false && !/[?&]noengine\b/.test(window.location?.search || '');
