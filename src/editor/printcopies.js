// Printing pages of more than one size: one browser page size can't slice them, so each page is printed as its own
// sheet (a named @page of its size) holding a copy of the pages, shifted into place.

/**
 * The print copies of `sheet` for `pages` pages, placed by `pageW(p)`, `pageH(p)` and `pageTop(p)` (px): an element to
 * add to the body for the print, or null when the pages are all one size (the sheet prints as it is).
 */
export function printCopies(sheet, pages, { pageW, pageH, pageTop }) {
  const sizeOf = (p) => `${Math.round(pageW(p) * 100)}x${Math.round(pageH(p) * 100)}`;
  const sizes = new Set(Array.from({ length: pages }, (_, p) => sizeOf(p)));
  if (sizes.size < 2) return null;
  const printBox = document.createElement('div');
  printBox.className = 'odx-print-pages';
  const css = [];
  for (const k of sizes) { const [w, h] = k.split('x').map((v) => Number(v) / 100); css.push(`@page odx-h${k} { size: ${w}px ${h}px; margin: 0; }`); }
  const style = document.createElement('style');
  style.textContent = css.join('\n') + `\n@media print { body > :not(.odx-print-pages):not(#odx-none) { display: none !important; } html, body { margin: 0 !important; display: block !important; height: auto !important; overflow: visible !important; } }`;
  printBox.append(style);
  // A copy of the sheet for each page, shifted to show that page. (One copy for a run of pages of one size was tried:
  // Edge printed the run's pages over one another, an absolutely placed copy isn't cut across pages; corpus 2's c2-07
  // lost 75 lines. In the flow, cut into pages by the browser, it moved lines the page's own box doesn't.)
  // Each holding only what its page shows: the whole sheet in each, a 36-page document of large tables took the print
  // 230s. Where things are, read from the sheet before any copy is made (each read would lay out the page again after).
  const copyOf = copies(sheet, printBox);
  for (let p = 0; p < pages; p++) {
    const h = pageH(p);
    const one = document.createElement('div');
    // (Each page its own stacking context: the browser paints a page by going through every layer of the stacking
    // context it's in, and with the whole print one, each line of the print layer a layer of its own (its transform), it
    // went through them all for each page: 140s of a 524-page document's print.)
    one.style.cssText = `page:odx-h${sizeOf(p)};width:${pageW(p)}px;height:${h}px;overflow:hidden;position:relative;break-after:page;isolation:isolate`;
    const c = copyOf(pageTop(p), pageTop(p) + h);
    // (Not cut at its own width, the first section's: a landscape page's sheet is wider than a portrait one's, and printed
    // cut off at the portrait width, longfn-word fnland. The page's box above cuts it at the page's own size.)
    c.style.cssText += `;display:block;position:absolute;left:0;top:${-pageTop(p)}px;height:${pageTop(pages)}px;overflow:visible`;
    one.append(c);
    printBox.append(one);
  }
  return printBox;
}

// How far past its page something can still show (a line's letters above or below its box, a shadow), px.
const ROOM = 100;
// What can show outside its block's box (pictures, shapes, floats, equations): blocks holding any are kept whole.
const LOOSE = 'img:not(.ProseMirror-separator), svg, math, canvas, iframe, [style*="position: absolute"], [style*="position:absolute"], [style*="position: fixed"], [style*="float"]';
// Placed absolutely, nothing else moves for them: the page boxes (background, header, footer), repeated header rows and
// line numbers; the print layer's boxes, and its lines (printlines.js), in copies of their table and cell elements.
const PLACED = ['odx-sheets', 'odx-repeats', 'odx-lnnums'];
const LINE = '[style*="translateY("]';

/**
 * Where the sheet's parts are, read once, and a function making a copy of it with what shows between `from` and `to`
 * (px from the sheet's top): without the placed parts of other pages; each of the document's blocks wholly on other
 * pages an empty copy of itself at its height, so the page's own are where they were (their margins, and the rules on
 * their place among their siblings, the same). (Only what's kept is copied: the whole sheet copied for each page and
 * cut after, the copying grew as the square of the pages.)
 */
function copies(sheet, printBox) {
  const top0 = sheet.getBoundingClientRect().top;
  const spans = new Map();
  const span = (e) => { const r = e.getBoundingClientRect(); const s = [r.top - top0, r.bottom - top0]; spans.set(e, s); return s; };
  const layer = sheet.querySelector(':scope > .odx-printlines');
  const host = layer?.querySelector(':scope > .odx-pl-host');
  for (const c of PLACED) for (const e of sheet.querySelectorAll(`:scope > .${c} > *`)) span(e);
  if (layer) for (const e of layer.children) if (e !== host && e.tagName !== 'STYLE') span(e);
  const lines = new Map(); // a host child -> the lines in it (itself, when it is one)
  if (host) for (const e of host.children) { const ls = [...(e.matches(LINE) ? [e] : []), ...e.querySelectorAll(LINE)]; ls.forEach(span); lines.set(e, ls); }
  const content = sheet.querySelector(':scope > .odx-content');
  // Where each shows (v), and where it sits in the flow (a, b: less a relative move, a print baseline shift, and a
  // translation, a column's (pages.js placeColumns): 0663-like documents in columns had their blocks kept whole in each
  // page's copy). Moved otherwise (turned, scaled): kept.
  const blocks = new Map();
  if (content) for (const e of content.children) {
    if (e.matches(LOOSE) || e.querySelector(LOOSE)) continue;
    const cs = getComputedStyle(e);
    if (cs.float !== 'none' || !/^(static|relative)$/.test(cs.position) || !/^(block|flow-root|list-item)$/.test(cs.display) || (cs.translate && cs.translate !== 'none')) continue;
    let ty = 0;
    if (cs.transform !== 'none') {
      const m = new DOMMatrixReadOnly(cs.transform);
      if (!m.is2D || m.a !== 1 || m.b !== 0 || m.c !== 0 || m.d !== 1) continue;
      ty = m.f;
    }
    const r = e.getBoundingClientRect();
    const rel = cs.position === 'relative' ? parseFloat(cs.top) || 0 : 0;
    blocks.set(e, { va: r.top - top0, vb: r.bottom - top0, a: r.top - top0 - rel - ty, b: r.bottom - top0 - rel - ty, mt: cs.marginTop, mb: cs.marginBottom });
  }
  // Rules on the blocks by their place among their siblings (the print baseline shifts, baseline.js): with any, each
  // block left out keeps a place (hidden), so the page's own blocks keep theirs.
  const placed = [...document.querySelectorAll('style')].some((s) => s.textContent.includes('> .odx-content > :nth-child('));
  const style = document.createElement('style');
  style.textContent = '.odx-pc-cut::before, .odx-pc-cut::after { content: none !important; }';
  printBox.append(style);
  return (from, to) => {
    const away = (e) => { const s = spans.get(e); return !!s && s[1] > s[0] && (s[1] < from - ROOM || s[0] > to + ROOM); };
    // A copy of `e` without the lines of its own (`ls`, in document order) that are away.
    const trimmed = (e, ls) => {
      const c = e.cloneNode(true);
      const own = [...(c.matches(LINE) ? [c] : []), ...c.querySelectorAll(LINE)];
      ls.forEach((l, i) => { if (l !== e && away(l)) own[i].remove(); });
      return c;
    };
    const copy = sheet.cloneNode(false);
    copy.classList.add('odx-print-copy');
    for (const part of sheet.children) {
      if (part === content) {
        const c = part.cloneNode(false);
        const kids = [...part.children];
        // (Off the page where it shows and where it sits in the flow: a block in columns cut where only the one was,
        // 0663's pages 74 and 81 had lines moved and lost.)
        const gone = (e) => { const k = blocks.get(e); return !!k && (k.vb < from - ROOM || k.va > to + ROOM) && (k.b < from - ROOM || k.a > to + ROOM); };
        for (let i = 0; i < kids.length;) {
          if (!gone(kids[i])) { c.append(kids[i].cloneNode(true)); i++; continue; }
          // A run of blocks on other pages: one empty copy of its last (what the block after it follows, for the rules
          // on a block's neighbours), as tall as the run, with its first's margin above and its last's below. (A copy of
          // each block, a 535-page document's 3,190 in each page's copy, had made the print grow as the square of its
          // pages.)
          let j = i;
          while (j + 1 < kids.length && gone(kids[j + 1])) j++;
          if (placed) for (let k = i; k < j; k++) { const h = kids[k].cloneNode(false); h.removeAttribute('id'); h.classList.add('odx-pc-cut'); h.style.cssText += ';display:none !important'; c.append(h); }
          const first = blocks.get(kids[i]), last = blocks.get(kids[j]);
          const cut = kids[j].cloneNode(false);
          cut.removeAttribute('id');
          cut.classList.add('odx-pc-cut');
          cut.style.cssText += `;height:${last.b - first.a}px !important;min-height:0 !important;max-height:none !important;box-sizing:border-box !important;margin-top:${first.mt} !important;margin-bottom:${last.mb} !important;visibility:hidden !important`;
          c.append(cut);
          i = j + 1;
        }
        copy.append(c);
      } else if (part === layer) {
        const c = part.cloneNode(false);
        for (const e of part.children) {
          if (e !== host) { if (!away(e)) c.append(e.cloneNode(true)); continue; }
          const hc = e.cloneNode(false);
          for (const k of e.children) {
            const ls = lines.get(k);
            // (A line's table and cell elements (display: contents) with none of its lines left: not copied.)
            if ((k.matches(LINE) && away(k)) || !ls.some((l) => !away(l))) continue;
            hc.append(ls.some(away) ? trimmed(k, ls) : k.cloneNode(true));
          }
          c.append(hc);
        }
        copy.append(c);
      } else if (PLACED.some((n) => part.classList.contains(n))) {
        const c = part.cloneNode(false);
        for (const e of part.children) if (!away(e)) c.append(e.cloneNode(true));
        copy.append(c);
      } else copy.append(part.cloneNode(true));
    }
    return copy;
  };
}
