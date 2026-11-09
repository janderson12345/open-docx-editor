// PDF export drawn from the editor's own page layout (pdf-lib, MIT): each sheet becomes a PDF page of the same size,
// and on it, at the places the layout put them, the text (real text: selectable and searchable), pictures, borders,
// shading, underlines and list labels. Separate from the main bundle: import 'open-docx-editor/pdf'.
//
// Fonts, run by run, the first that has the face: the document's own embedded fonts; the fonts installed on the
// computer, read with the browser's Local Font Access (Chromium browsers: it asks the user once); else the PDF standard
// fonts (Helvetica, Times, Courier and their bold and italic faces). Real fonts are embedded as subsets (fontkit). Each
// run is drawn at the exact width the layout gave it, so lines, line breaks and pages are the layout's in any font;
// with a standard font the letter shapes are its own, and characters it lacks are drawn as "?" (it covers
// Windows-1252's only).
import { PDFDocument, StandardFonts, rgb, PDFOperator, PDFOperatorNames, PDFNumber, pushGraphicsState, popGraphicsState, beginText, endText, setFontAndSize, setTextMatrix, showText, setFillingRgbColor, setWordSpacing } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import { embeddedFonts } from './ooxml/fonts.js';
import { formatNumber } from './ooxml/styles.js';

const PT = 0.75; // CSS px to PDF points
// Not drawn: what printing hides too (editing aids, comments beside the page, the screen-reader text).
const SKIP = '.odx-markup, .odx-ruler, .odx-imgsize, .odx-fm-layer, .odx-hfedit-tag, .odx-sr, .odx-live, .odx-nav, .odx-findbar, .odx-toolbar, .odx-ctx, .odx-pop, .odx-balloons, .ProseMirror-separator, .ProseMirror-trailingBreak, .odx-lifted';

/**
 * The editor's pages as a PDF. `editor` is mountDocxEditor's API or an editor view. Returns the PDF's bytes.
 * Options: `title`, `author` (document information); `localFonts` false not to ask for the installed fonts.
 */
export async function exportPdf(editor, { title, author, localFonts = true } = {}) {
  const view = editor.view || editor;
  const pageEl = view.dom.closest('.odx-page') || view.dom.parentElement;
  // Long documents skip drawing off-screen paragraphs: draw them all while exporting, as for printing.
  pageEl.classList.add('odx-exporting');
  await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
  // Lines where printing puts them: the print layer, each line exactly placed with its baseline shift (the originals
  // are hidden while it's there).
  const unlayer = view.odx?.pages?.printLayer ? view.odx.pages.printLayer() : () => {};
  try {
    const sources = await fontSources(view.odx?.session?.pkg, localFonts);
    return await draw(pageEl, { title, author, sources });
  } finally {
    unlayer();
    pageEl.classList.remove('odx-exporting');
  }
}

/**
 * Where the fonts come from: { embedded: Map(key -> bytes), local: Map(key -> FontData) }, keys "family|bold|italic"
 * in lower case. The installed fonts only when the browser has Local Font Access and the user allows it.
 */
async function fontSources(pkg, local) {
  const key = (family, bold, italic) => `${family.toLowerCase()}|${bold ? 1 : 0}|${italic ? 1 : 0}`;
  const embedded = new Map();
  if (pkg) for (const f of embeddedFonts(pkg)) embedded.set(key(f.family, (parseInt(f.weight, 10) || 400) >= 600, f.style === 'italic'), f.bytes);
  const installed = new Map();
  if (local && typeof window !== 'undefined' && typeof window.queryLocalFonts === 'function') {
    try {
      for (const f of await window.queryLocalFonts()) {
        const s = (f.style || '').toLowerCase();
        const b = /bold|black|heavy|semibold|demibold/.test(s), i = /italic|oblique/.test(s);
        const k = key(f.family, b, i);
        // (The plain face of each style: "Bold" rather than "Bold Condensed".)
        if (!installed.has(k) || /^(regular|bold|italic|bold italic)$/.test(s)) installed.set(k, f);
        // By full name too, less the style words: "Calibri Light" is the family Calibri, style Light, on Windows.
        const named = key((f.fullName || '').replace(/\s+(bold|italic|oblique|regular)\b/gi, '').trim(), b, i);
        if (f.fullName && !installed.has(named)) installed.set(named, f);
      }
    } catch { /* not allowed (a permission refused, or a page that may not ask): the standard fonts */ }
  }
  return { key, embedded, installed };
}

async function draw(pageEl, { title, author, sources }) {
  const pdf = await PDFDocument.create();
  pdf.registerFontkit(fontkit);
  if (title) pdf.setTitle(title);
  if (author) pdf.setAuthor(author);
  pdf.setProducer('open-docx-editor');
  const fonts = {};
  const fontFor = async (cs) => {
    const bold = (parseInt(cs.fontWeight, 10) || 400) >= 600;
    const italic = /italic|oblique/.test(cs.fontStyle);
    // The real font: the first family of the run's stack the document embeds or the computer has, in this face.
    for (const fam of cs.fontFamily.split(',').map((f) => f.trim().replace(/^["']|["']$/g, ''))) {
      const k = sources.key(fam, bold, italic);
      if (fonts[k] !== undefined) { if (fonts[k]) return fonts[k]; continue; }
      let bytes = sources.embedded.get(k);
      if (!bytes && sources.installed.has(k)) {
        const f = sources.installed.get(k);
        bytes = fromCollection(new Uint8Array(await (await f.blob()).arrayBuffer()), f.postscriptName);
      }
      fonts[k] = null;
      if (!bytes) continue;
      try {
        const font = await pdf.embedFont(bytes, { subset: true });
        fonts[k] = { font, chars: new Set(font.getCharacterSet()) };
        return fonts[k];
      } catch { /* a font fontkit can't read (a collection, a broken table): the next family */ }
    }
    const fam = cs.fontFamily.toLowerCase();
    const mono = /courier|consol|mono/.test(fam);
    const serif = !mono && /times|cambria|georgia|garamond|book|serif|palatino|constantia/.test(fam.split(',')[0]) && !/sans/.test(fam.split(',')[0]);
    const key =(mono ? 'Courier' : serif ? 'TimesRoman' : 'Helvetica') + (bold && italic ? (serif ? 'BoldItalic' : 'BoldOblique') : bold ? 'Bold' : italic ? (serif ? 'Italic' : 'Oblique') : '');
    if (!fonts[key]) {
      const font = await pdf.embedFont(StandardFonts[key]);
      fonts[key] = { font, chars: new Set(font.getCharacterSet()) };
    }
    return fonts[key];
  };
  const sheets = [...pageEl.querySelectorAll('.odx-sheets > .odx-sheet')].map((s) => ({ el: s, r: s.getBoundingClientRect() }));
  if (!sheets.length) throw new Error('No pages to export.');
  const pages = sheets.map(({ r }) => pdf.addPage([r.width * PT, r.height * PT]));
  const sheetAt = (y) => sheets.findIndex(({ r }) => y >= r.top && y < r.bottom);
  const ctx = document.createElement('canvas').getContext('2d');

  // The page in layers, as printing paints it.
  const d = { pdf, pageEl, sheets, pages, sheetAt, drawText, drawBullet, fontFor, ascent, firstLine };
  drawBoxes(d);
  await drawPictures(d);
  const mathDrawn = await drawMath(d);
  await drawTextLayer(d, mathDrawn);

  return pdf.save();

  /** A bullet the standard fonts lack, drawn as its shape (filled or hollow circle, square, check). */
  function drawBullet(label, r, cs) {
    const shape = { '●': 'disc', '•': null, '': 'disc', '○': 'circle', '◦': 'circle', 'o': null, '▪': 'square', '■': 'square', '': 'square', '✓': 'check', '✔': 'check', '': 'check' }[label];
    if (!shape) return false;
    const n = sheetAt(r.top + 2);
    if (n < 0) return false;
    const sr = sheets[n].r;
    const size = (parseFloat(cs.fontSize) || 16) * PT;
    const c = (color(cs.color) || { rgb: rgb(0, 0, 0) }).rgb;
    const x = (r.left - sr.left) * PT, base = (sr.bottom - (r.top + ascent(cs))) * PT;
    const d = size * 0.36, cy = base + size * 0.33, cx = x + d / 2 + size * 0.05;
    const page = pages[n];
    if (shape === 'disc') page.drawCircle({ x: cx, y: cy, size: d / 2, color: c });
    else if (shape === 'circle') page.drawCircle({ x: cx, y: cy, size: d / 2, borderColor: c, borderWidth: size / 16 });
    else if (shape === 'square') page.drawRectangle({ x: cx - d / 2, y: cy - d / 2, width: d, height: d, color: c });
    else page.drawSvgPath(`M0 ${-d * 0.5} L${d * 0.35} ${-d * 0.1} L${d} ${-d * 0.9}`, { x: cx - d / 2, y: cy + d / 2, borderColor: c, borderWidth: size / 12 });
    return true;
  }

  /** One run of text at its laid-out box: baseline from the font's ascent, width matched by scaling or word spacing. */
  async function drawText(text, r, cs, el) {
    if (!r) return;
    const n = sheetAt(r.top + Math.min(r.height || 0, 4));
    if (n < 0) return;
    const sr = sheets[n].r;
    let { font, chars } = await fontFor(cs);
    // Greek letters and math signs a standard font lacks, in a run of nothing else (an equation's π, ∑, ∫, θ): the
    // standard Symbol font has them.
    // (Not the dashes and minus the standard fonts take as a hyphen, below.)
    const missing = [...text].filter((ch) => ch.trim() && !/[\u2010-\u2012\u2015\u2212]/.test(ch) && !chars.has(ch.codePointAt(0)));
    if (missing.length && missing.length === [...text].filter((ch) => ch.trim()).length) {
      if (!fonts.Symbol) { const f = await pdf.embedFont(StandardFonts.Symbol); fonts.Symbol = { font: f, chars: new Set(f.getCharacterSet()) }; }
      if (missing.every((ch) => fonts.Symbol.chars.has(ch.codePointAt(0)))) ({ font, chars } = fonts.Symbol);
    }
    const safe = [...text.replace(/[\u00a0\u2000-\u200a\u202f\u205f]/g, ' ').replace(/[\u200b-\u200d\u2060\ufeff\u00ad]/g, '').replace(/[\u2010\u2011\u2012\u2212]/g, '-').replace(/\u2015/g, '\u2014')]
      .map((ch) => (chars.has(ch.codePointAt(0)) ? ch : ch.trim() ? '?' : ' ')).join(''); // (hyphens, spaces and joiners the standard fonts lack, as theirs)
    if (!safe.trim()) return;
    const size = (parseFloat(cs.fontSize) || 16) * PT;
    const c = color(cs.color) || { rgb: rgb(0, 0, 0) };
    const x = (r.left - sr.left) * PT;
    const baseline = r.top + ascent(cs);
    const y = (sr.bottom - baseline) * PT;
    const natural = font.widthOfTextAtSize(safe, size);
    const want = r.width != null ? r.width * PT : natural;
    const spaces = (safe.match(/ /g) || []).length;
    const page = pages[n];
    const key = page.node.newFontDictionary(font.name, font.ref);
    const ops = [pushGraphicsState(), beginText(), setFontAndSize(key, size), setFillingRgbColor(c.rgb.red, c.rgb.green, c.rgb.blue)];
    // Justified lines stretch their spaces; others scale to the width the layout measured (a different font's).
    if (spaces && want > natural * 1.03 && cs.textAlign === 'justify') ops.push(setWordSpacing((want - natural) / spaces));
    else if (natural > 0 && Math.abs(want - natural) > 0.1) ops.push(PDFOperator.of(PDFOperatorNames.SetTextHorizontalScaling, [PDFNumber.of(+(100 * want / natural).toFixed(2))]));
    ops.push(setTextMatrix(1, 0, 0, 1, x, y), showText(font.encodeText(safe)), endText(), popGraphicsState());
    page.pushOperators(...ops);
    // Underline and strikethrough, from the element that sets them (text-decoration isn't inherited, it's drawn through).
    for (let e = el; e && e !== pageEl; e = e.parentElement) {
      const d = getComputedStyle(e);
      if (d.display !== 'inline' && d.display !== 'contents') break;
      const lines = d.textDecorationLine;
      if (!lines || lines === 'none') continue;
      const dc = color(d.textDecorationColor) || c;
      const thick = Math.max(0.5, size / 18);
      if (lines.includes('underline')) page.drawLine({ start: { x, y: y - size * 0.12 }, end: { x: x + want, y: y - size * 0.12 }, thickness: thick, color: dc.rgb });
      if (lines.includes('line-through')) page.drawLine({ start: { x, y: y + size * 0.3 }, end: { x: x + want, y: y + size * 0.3 }, thickness: thick, color: dc.rgb });
    }
  }

  function ascent(cs) {
    ctx.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
    const m = ctx.measureText('Hg');
    return m.fontBoundingBoxAscent ?? (parseFloat(cs.fontSize) * 0.8);
  }

  /** The baseline of an element's first line of text. */
  function firstLine(el, rg) {
    const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    for (let t = w.nextNode(); t; t = w.nextNode()) {
      if (!t.data.trim() || t.parentElement.closest(SKIP)) continue;
      rg.setStart(t, 0); rg.setEnd(t, 1);
      const q = rg.getClientRects()[0];
      if (q) return { baseline: q.top + ascent(getComputedStyle(t.parentElement)) };
    }
    return null;
  }
}

// 1. Shading and borders, under everything else.
function drawBoxes(d) {
  const { pageEl, sheets, pages } = d;
  for (const el of pageEl.querySelectorAll('*')) {
    if (el.closest(SKIP) || el.classList.contains('odx-sheet') || el.classList.contains('odx-sheets')) continue;
    const cs = getComputedStyle(el);
    if (cs.visibility !== 'visible' || cs.display === 'none' || cs.display === 'contents') continue;
    const bg = color(cs.backgroundColor);
    const sides = ['Top', 'Right', 'Bottom', 'Left'].map((s) => ({ s, w: parseFloat(cs[`border${s}Width`]) || 0, c: color(cs[`border${s}Color`]), style: cs[`border${s}Style`] }))
      .filter((b) => b.w > 0 && b.c && b.style !== 'none' && b.style !== 'hidden');
    if (!bg && !sides.length) continue;
    for (const r of el.getClientRects()) {
      for (let n = 0; n < sheets.length; n++) {
        const sr = sheets[n].r;
        const top = Math.max(r.top, sr.top), bottom = Math.min(r.bottom, sr.bottom);
        if (bottom <= top || r.right <= sr.left || r.left >= sr.right) continue;
        const page = pages[n];
        const X = (x) => (x - sr.left) * PT, Y = (y) => (sr.bottom - y) * PT;
        if (bg) page.drawRectangle({ x: X(r.left), y: Y(bottom), width: r.width * PT, height: (bottom - top) * PT, color: bg.rgb, opacity: bg.a });
        for (const b of sides) {
          const t = b.w * PT;
          const dash = b.style === 'dashed' ? [3 * t, 2 * t] : b.style === 'dotted' ? [t, t] : undefined;
          const line = (x1, y1, x2, y2) => page.drawLine({ start: { x: x1, y: y1 }, end: { x: x2, y: y2 }, thickness: t, color: b.c.rgb, opacity: b.c.a, dashArray: dash });
          if (b.s === 'Top' && r.top >= sr.top) line(X(r.left), Y(r.top + b.w / 2), X(r.right), Y(r.top + b.w / 2));
          if (b.s === 'Bottom' && r.bottom <= sr.bottom) line(X(r.left), Y(r.bottom - b.w / 2), X(r.right), Y(r.bottom - b.w / 2));
          if (b.s === 'Left') line(X(r.left + b.w / 2), Y(top), X(r.left + b.w / 2), Y(bottom));
          if (b.s === 'Right') line(X(r.right - b.w / 2), Y(top), X(r.right - b.w / 2), Y(bottom));
        }
      }
    }
  }
}

// 2. Pictures.
async function drawPictures(d) {
  const { pdf, pageEl, sheets, pages, sheetAt } = d;
  const images = new Map();
  for (const img of pageEl.querySelectorAll('img')) {
    if (img.closest(SKIP) || !img.complete || !img.naturalWidth) continue;
    const r = clipped(img, img.getBoundingClientRect());
    if (!r || r.width < 1 || r.height < 1) continue;
    const n = sheetAt(r.top + r.height / 2);
    if (n < 0) continue;
    let embedded = images.get(img.src);
    if (!embedded) {
      embedded = await embedImage(pdf, img).catch(() => null);
      images.set(img.src, embedded);
    }
    if (!embedded) continue;
    const sr = sheets[n].r;
    pages[n].drawImage(embedded, { x: (r.left - sr.left) * PT, y: (sr.bottom - r.bottom) * PT, width: r.width * PT, height: r.height * PT });
  }
}

// 2b. Equations' drawn parts (MathML paints them, they aren't text): fraction bars, radical signs, and brackets
// stretched to their content; large operators (a sum, an integral) at the size the layout gave them.
async function drawMath(d) {
  const { pageEl, sheets, pages, sheetAt, drawText } = d;
  const mathDrawn = new Set();
  for (const m of pageEl.querySelectorAll('math')) {
    if (m.closest(SKIP)) continue;
    const mr = m.getBoundingClientRect();
    const n = sheetAt(mr.top + mr.height / 2);
    if (n < 0) continue;
    const sr = sheets[n].r, page = pages[n];
    const X = (x) => (x - sr.left) * PT, Y = (y) => (sr.bottom - y) * PT;
    const ink = (e) => (color(getComputedStyle(e).color) || { rgb: rgb(0, 0, 0) }).rgb;
    const em = (e) => parseFloat(getComputedStyle(e).fontSize) || 16;
    const box = (els) => els.map((e) => e.getBoundingClientRect()).reduce((a, r) => (a ? { left: Math.min(a.left, r.left), top: Math.min(a.top, r.top), right: Math.max(a.right, r.right), bottom: Math.max(a.bottom, r.bottom) } : { left: r.left, top: r.top, right: r.right, bottom: r.bottom }), null);
    for (const f of m.querySelectorAll('mfrac')) {
      if (f.getAttribute('linethickness') === '0' || f.children.length < 2) continue;
      const r = f.getBoundingClientRect(), num = f.children[0].getBoundingClientRect(), den = f.children[1].getBoundingClientRect();
      const y = (num.bottom + den.top) / 2, t = Math.max(0.5, (em(f) / 16) * 0.75);
      page.drawLine({ start: { x: X(r.left + 1), y: Y(y) }, end: { x: X(r.right - 1), y: Y(y) }, thickness: t, color: ink(f) });
    }
    for (const q of m.querySelectorAll('msqrt, mroot')) {
      const inner = q.tagName.toLowerCase() === 'mroot' ? box([q.children[0]]) : box([...q.children]);
      if (!inner) continue;
      const r = q.getBoundingClientRect(), e = em(q), c = ink(q), t = Math.max(0.5, (e / 16) * 0.75);
      const top = Math.min(inner.top, r.top + 1), x0 = inner.left - e * 0.55;
      // The sign: a short tick up to the left, down to the bottom, up to the top, then the bar over the radicand.
      page.drawSvgPath(`M${X(x0)} ${-Y(inner.bottom - (inner.bottom - top) * 0.4)} L${X(x0 + e * 0.12)} ${-Y(inner.bottom - (inner.bottom - top) * 0.45)} L${X(x0 + e * 0.3)} ${-Y(inner.bottom)} L${X(inner.left - e * 0.08)} ${-Y(top)} L${X(inner.right)} ${-Y(top)}`, { x: 0, y: 0, borderColor: c, borderWidth: t });
    }
    for (const o of m.querySelectorAll('mo')) {
      const r = o.getBoundingClientRect(), e = em(o), ch = o.textContent.trim();
      if (!ch || !r.height) continue;
      if (o.getAttribute('largeop') === 'true' && r.height > e * 1.3) {
        const cs = getComputedStyle(o);
        const big = { fontSize: `${r.height * 0.82}px`, fontFamily: cs.fontFamily, fontStyle: cs.fontStyle, fontWeight: cs.fontWeight, color: cs.color, textAlign: 'left' };
        await drawText(ch, { left: r.left, top: r.top, width: r.width, height: r.height }, big, o);
        mathDrawn.add(o);
      } else if (o.getAttribute('stretchy') === 'true' && r.height > e * 1.4 && '()[]{}|'.includes(ch)) {
        const c = ink(o), t = Math.max(0.6, (e / 16) * 0.8), w = Math.min(r.width, e * 0.35);
        const l = r.left + (r.width - w) / 2, rt = l + w, top = r.top + 1, bot = r.bottom - 1, mid = (top + bot) / 2;
        const P = (pts) => pts.map(([x, y], i) => `${i ? 'L' : 'M'}${X(x)} ${-Y(y)}`).join(' ');
        const d = ch === '[' ? P([[rt, top], [l, top], [l, bot], [rt, bot]])
          : ch === ']' ? P([[l, top], [rt, top], [rt, bot], [l, bot]])
          : ch === '|' ? P([[(l + rt) / 2, top], [(l + rt) / 2, bot]])
          : ch === '(' ? `M${X(rt)} ${-Y(top)} Q${X(l - w * 0.6)} ${-Y(mid)} ${X(rt)} ${-Y(bot)}`
          : ch === ')' ? `M${X(l)} ${-Y(top)} Q${X(rt + w * 0.6)} ${-Y(mid)} ${X(l)} ${-Y(bot)}`
          : ch === '{' ? `M${X(rt)} ${-Y(top)} Q${X(l + w * 0.4)} ${-Y(top)} ${X(l + w * 0.4)} ${-Y(mid - (mid - top) * 0.4)} Q${X(l + w * 0.4)} ${-Y(mid)} ${X(l)} ${-Y(mid)} Q${X(l + w * 0.4)} ${-Y(mid)} ${X(l + w * 0.4)} ${-Y(mid + (bot - mid) * 0.4)} Q${X(l + w * 0.4)} ${-Y(bot)} ${X(rt)} ${-Y(bot)}`
          : `M${X(l)} ${-Y(top)} Q${X(rt - w * 0.4)} ${-Y(top)} ${X(rt - w * 0.4)} ${-Y(mid - (mid - top) * 0.4)} Q${X(rt - w * 0.4)} ${-Y(mid)} ${X(rt)} ${-Y(mid)} Q${X(rt - w * 0.4)} ${-Y(mid)} ${X(rt - w * 0.4)} ${-Y(mid + (bot - mid) * 0.4)} Q${X(rt - w * 0.4)} ${-Y(bot)} ${X(l)} ${-Y(bot)}`;
        page.drawSvgPath(d, { x: 0, y: 0, borderColor: c, borderWidth: t });
        mathDrawn.add(o);
      }
    }
  }
  return mathDrawn;
}

// 3. Text, line by line: each text node's pieces on each line.
async function drawTextLayer(d, mathDrawn) {
  const { pageEl, sheetAt, drawText, drawBullet, fontFor, ascent, firstLine } = d;
  const range = document.createRange();
  const walker = document.createTreeWalker(pageEl, NodeFilter.SHOW_TEXT, {
    acceptNode: (t) => (!t.data.trim() || t.parentElement.closest(SKIP) || mathDrawn.has(t.parentElement) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  for (let t = walker.nextNode(); t; t = walker.nextNode()) {
    const el = t.parentElement;
    const cs = getComputedStyle(el);
    if (cs.visibility !== 'visible') continue;
    for (const piece of pieces(t, range)) await drawText(piece.text, clipped(el, piece.r), cs, el);
  }
  // Generated text the layout draws in CSS: list labels and note reference numbers.
  for (const p of pageEl.querySelectorAll('p.odx-li[data-label]')) {
    if (p.closest(SKIP) || !p.getClientRects().length || getComputedStyle(p).visibility !== 'visible') continue;
    const label = p.getAttribute('data-label').replace(/\t/g, ' ').trim();
    if (!label) continue;
    const ls = getComputedStyle(p, '::before');
    const ps = getComputedStyle(p);
    const r = p.getBoundingClientRect();
    const first = firstLine(p, range);
    const left = r.left + (parseFloat(ps.borderLeftWidth) || 0) + (parseFloat(ps.paddingLeft) || 0) + (parseFloat(ps.textIndent) || 0);
    const top = first ? first.baseline - ascent(ls) : r.top;
    const box = { left, top, width: null, height: parseFloat(ls.fontSize) };
    if (!drawBullet(label, box, ls)) await drawText(label, box, ls, p);
  }
  // Tab leaders drawn as characters (CSS content): as many as fit the tab's width.
  for (const tab of pageEl.querySelectorAll('.odx-tab[data-leader]')) {
    if (tab.closest(SKIP) || getComputedStyle(tab).visibility !== 'visible') continue;
    const r = clipped(tab, tab.getBoundingClientRect());
    if (!r || r.width < 2) continue;
    const ch = { dot: '.', middleDot: '·', hyphen: '-' }[tab.dataset.leader];
    const cs = getComputedStyle(tab, '::after');
    const { font } = await fontFor(cs);
    const size = (parseFloat(cs.fontSize) || 16) * PT;
    const count = Math.floor((r.width * PT) / font.widthOfTextAtSize(ch, size));
    if (count > 0) await drawText(ch.repeat(count), { left: r.left, top: r.top, width: null, height: r.height }, cs, tab);
  }
  // (Numbered in the order they show on the pages: the print layer's copies aren't in the document's order in the DOM.)
  const refs = [...pageEl.querySelectorAll('.odx-content [data-kind="footnoteReference"], .odx-content [data-kind="endnoteReference"]')]
    .filter((e) => !e.closest(SKIP) && e.getClientRects().length && getComputedStyle(e).visibility === 'visible')
    .map((e) => ({ e, r: e.getBoundingClientRect() }))
    .sort((a, b) => sheetAt(a.r.top) - sheetAt(b.r.top) || a.r.top - b.r.top || a.r.left - b.r.left);
  let fn = 0, en = 0;
  for (const { e: ref } of refs) {
    const kind = ref.getAttribute('data-kind');
    const label = kind === 'footnoteReference' ? String(++fn) : formatNumber(++en, 'lowerRoman');
    const s = getComputedStyle(ref, '::after');
    const r = ref.getBoundingClientRect();
    await drawText(label, { left: r.left, top: r.top, width: r.width || null, height: r.height }, s, ref);
  }
}

/**
 * A font file of its own from a collection (.ttc, as Windows ships Cambria): the font whose PostScript name is `ps`
 * (else the first), its tables copied into a new file. Anything else comes back as it is.
 */
function fromCollection(bytes, ps) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 12 || dv.getUint32(0) !== 0x74746366) return bytes; // 'ttcf'
  const offsets = Array.from({ length: dv.getUint32(8) }, (_, k) => dv.getUint32(12 + 4 * k));
  const tables = (at) => Array.from({ length: dv.getUint16(at + 4) }, (_, k) => {
    const r = at + 12 + 16 * k;
    return { tag: dv.getUint32(r), sum: dv.getUint32(r + 4), off: dv.getUint32(r + 8), len: dv.getUint32(r + 12) };
  });
  const psName = (at) => {
    const name = tables(at).find((t) => t.tag === 0x6e616d65); // 'name'
    if (!name) return null;
    const count = dv.getUint16(name.off + 2), strings = name.off + dv.getUint16(name.off + 4);
    for (let k = 0; k < count; k++) {
      const r = name.off + 6 + 12 * k;
      if (dv.getUint16(r + 6) !== 6) continue; // PostScript name
      const platform = dv.getUint16(r), len = dv.getUint16(r + 8), off = strings + dv.getUint16(r + 10);
      let s = '';
      if (platform === 3 || platform === 0) for (let j = 0; j < len; j += 2) s += String.fromCharCode(dv.getUint16(off + j));
      else for (let j = 0; j < len; j++) s += String.fromCharCode(bytes[off + j]);
      return s;
    }
    return null;
  };
  const at = offsets.find((o) => psName(o) === ps) ?? offsets[0];
  const list = tables(at);
  const head = 12 + 16 * list.length;
  const size = head + list.reduce((t, x) => t + ((x.len + 3) & ~3), 0);
  const out = new Uint8Array(size);
  const ov = new DataView(out.buffer);
  ov.setUint32(0, dv.getUint32(at));
  for (let k = 4; k < 12; k += 2) ov.setUint16(k, dv.getUint16(at + k));
  let pos = head;
  list.forEach((t, k) => {
    const r = 12 + 16 * k;
    ov.setUint32(r, t.tag); ov.setUint32(r + 4, t.sum); ov.setUint32(r + 8, pos); ov.setUint32(r + 12, t.len);
    out.set(bytes.subarray(t.off, t.off + t.len), pos);
    pos += (t.len + 3) & ~3;
  });
  return out;
}

/** A text node's pieces, one per line it's on: [{ text, r }], the line found where a character's box moves down. */
function pieces(t, rg) {
  const out = [];
  const n = t.data.length;
  let from = 0;
  while (from < n) {
    rg.setStart(t, from); rg.setEnd(t, from + 1);
    const first = rg.getClientRects()[0];
    if (!first) { from++; continue; }
    let lo = from, hi = n - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      rg.setStart(t, mid); rg.setEnd(t, mid + 1);
      const q = rg.getClientRects()[0];
      if (q && Math.abs(q.top - first.top) < 2) lo = mid; else hi = mid - 1;
    }
    // The piece's box, without a trailing space hanging past a wrapped line's end.
    let end = lo + 1;
    while (end > from + 1 && t.data[end - 1] === ' ' && end < n) end--;
    rg.setStart(t, from); rg.setEnd(t, end);
    const box = rg.getBoundingClientRect();
    if (box.width > 0) out.push({ text: t.data.slice(from, end), r: { left: box.left, top: box.top, width: box.width, height: box.height } });
    from = lo + 1;
  }
  return out;
}

/** `r` cut by an inset clip on the element or an ancestor (a paragraph or table split across columns), or null if gone. */
function clipped(el, r) {
  if (!r) return null;
  for (let e = el; e && !e.classList.contains('odx-page'); e = e.parentElement) {
    const m = /^inset\(([-\d.]+)px [-\d.]+px ([-\d.]+)px/.exec(getComputedStyle(e).clipPath || '');
    if (!m) continue;
    const b = e.getBoundingClientRect();
    if (+m[1] > 0 && r.top + (r.height || 0) / 2 < b.top + +m[1]) return null;
    if (+m[2] > 0 && r.top + (r.height || 0) / 2 > b.bottom - +m[2]) return null;
  }
  return r;
}

/** CSS color to pdf-lib's, with its alpha; null when transparent. */
function color(css) {
  const m = /rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?\)/.exec(css || '');
  if (!m) return null;
  const a = m[4] == null ? 1 : +m[4];
  if (a < 0.02) return null;
  return { rgb: rgb(+m[1] / 255, +m[2] / 255, +m[3] / 255), a };
}

async function embedImage(pdf, img) {
  const bytes = new Uint8Array(await (await fetch(img.src)).arrayBuffer());
  if (bytes[0] === 0x89 && bytes[1] === 0x50) return pdf.embedPng(bytes);
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return pdf.embedJpg(bytes);
  // Anything else the browser can draw (SVG from EMF, GIF, BMP...): drawn at twice its size, as a PNG.
  const c = document.createElement('canvas');
  const scale = 2;
  c.width = Math.max(1, Math.round(img.getBoundingClientRect().width * scale));
  c.height = Math.max(1, Math.round(img.getBoundingClientRect().height * scale));
  c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
  const png = await new Promise((r) => c.toBlob(r, 'image/png'));
  return pdf.embedPng(new Uint8Array(await png.arrayBuffer()));
}

