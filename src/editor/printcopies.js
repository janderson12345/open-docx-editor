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
  // A copy of the sheet for each page, shifted to show that page. (Costly for a long document: the whole sheet copied
  // once a page. One copy for a run of pages of one size was tried: Edge printed the run's pages over one another, an
  // absolutely placed copy isn't cut across pages; corpus 2's c2-07 lost 75 lines.)
  const copy = sheet.cloneNode(true);
  copy.classList.add('odx-print-copy');
  for (let p = 0; p < pages; p++) {
    const h = pageH(p);
    const one = document.createElement('div');
    one.style.cssText = `page:odx-h${sizeOf(p)};width:${pageW(p)}px;height:${h}px;overflow:hidden;position:relative;break-after:page`;
    const c = p ? copy.cloneNode(true) : copy;
    c.style.cssText += `;display:block;position:absolute;left:0;top:${-pageTop(p)}px;height:${pageTop(pages)}px`;
    one.append(c);
    printBox.append(one);
  }
  return printBox;
}
