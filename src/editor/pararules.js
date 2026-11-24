// The paragraph pass's rules (editor.js paragraphPlugin): each sets classes and style on one paragraph, from its node,
// the session and what the pass worked out before it. `attrs` is the paragraph's decoration (class, data-label, role),
// `style` its list of CSS declarations; both are added to in the pass's order.
import { listLevel, resolveStyle } from '../ooxml/styles.js';
import { twipsToPt, fontCss, fontStack, lineFactor, mergeTabs, sizeCss, hexColor, highlightColor } from '../ooxml/props.js';
import { paragraphChanged, paragraphMarkRev } from './track.js';
import { outlineLevel } from './commands.js';
import { wordAscent } from './baseline.js';

// The paragraph mark's own font size (half-points), if it sets one.
export function markSize(node) {
  const mark = (node.attrs.extra || []).find((e) => e[0] === 'rPr');
  const sz = mark && /<w:sz w:val="(\d+)"/.exec(mark[1]);
  return sz ? Number(sz[1]) : null;
}

/** A list paragraph: its label, indents and the room after the label. `l` is its list, `label` its label text. */
export function listLabel(node, session, l, label, attrs, style) {
  if (label != null) {
    const lvl = listLevel(session.numbering, l.numId, l.ilvl);
    const ind = lvl ? lvl.pPr : {};
    // A list that comes from the paragraph's style: an indent that style sets itself (it, or a style based on it) is
    // the paragraph's, over the list level's, each part on its own (Word's prints of generated documents: a level at
    // 2.375in hanging 0.25in under a style with a 0.25in left indent, the label at the margin and the text 0.25in in;
    // with a hanging indent of its own, that one). With the list set on the paragraph itself the level's stands.
    const styleInd = {};
    if (node.attrs.numId == null && node.attrs.styleId) {
      const seen = new Set();
      for (let id = node.attrs.styleId; id && !seen.has(id); ) {
        seen.add(id);
        const st = session.styles.styles.get(id);
        if (!st) break;
        if (styleInd.indLeft == null && st.pPr.indLeft != null) styleInd.indLeft = st.pPr.indLeft;
        if (styleInd.indFirst == null && st.pPr.indFirst != null) styleInd.indFirst = st.pPr.indFirst;
        if (st.pPr.numId != null) break;
        id = st.basedOn;
      }
    }
    const left = node.attrs.view.indLeft ?? styleInd.indLeft ?? ind.indLeft ?? 720;
    const hang = -(node.attrs.view.indFirst ?? styleInd.indFirst ?? ind.indFirst ?? -360);
    attrs.class = lvl && lvl.fmt === 'bullet' ? 'odx-li odx-bullet' : 'odx-li';
    // A number formatted as hidden text (w:vanish on its level, or on the paragraph mark it takes its look from) isn't
    // shown, nor the room after it, as in Word.
    const mark = (node.attrs.extra || []).find((e) => e[0] === 'rPr');
    // (Nor a paragraph holding only a page break: Word counts it and shows no number for it, an "Appendix B." that had
    // stood alone at the foot of the page before the break. Its prints of generated documents.)
    let breakOnly = node.childCount > 0;
    node.forEach((c) => { if (!(c.type.name === 'hard_break' && c.attrs.type === 'page')) breakOnly = false; });
    const hidden = breakOnly || !!lvl?.hidden || (!!mark && /<w:vanish(?=[\s/>])(?![^>]*w:val="(?:0|false|off)")/.test(mark[1]));
    attrs['data-label'] = hidden ? '' : label;
    // A picture bullet (w:lvlPicBulletId): the picture in the label's place (editor.css odx-picbullet).
    // (Its picture is in the document's stylesheet once, docx.js picBulletCss, not in every item's style.)
    if (lvl && lvl.picBullet != null && session.pictureBullets?.has(lvl.picBullet)) attrs.class += ` odx-picbullet odx-pb-${lvl.picBullet}`;
    // The text after the label starts at the hanging indent; with none (a first-line indent instead), at the
    // next default tab stop after the label's start.
    // (The paragraph's own tab stops come first; then the default ones.)
    const start = left - hang;
    const dts = session.defaultTabStop || 720;
    const own = mergeTabs(session.styleProp(node.attrs.styleId, 'tabs'), node.attrs.view?.tabs)
      .filter((t) => t.val !== 'bar' && t.val !== 'clear' && t.pos > start).map((t) => t.pos);
    // (Nothing after the label, w:suff="nothing": the text follows it at once, doc 24's unnumbered headings starting
    // at the hanging position.)
    const next = (own.length ? Math.min(...own) : (Math.floor(start / dts) + 1) * dts) - start;
    // (Word 2003's rule, w:doNotUseIndentAsNumberingTabStop: the hanging indent is the label's stop only when it comes
    // before the next tab stop; past it, the text after the label goes to that stop.)
    const viaIndent = hang > 0 ? (session.compat?.has('doNotUseIndentAsNumberingTabStop') ? Math.min(hang, next) : hang) : next;
    // (A space after the label, w:suff="space": the label and a space, the text right after, as nothing's.)
    const spaced = !hidden && lvl && lvl.suff === 'space';
    if (spaced && label != null) attrs['data-label'] = `${label} `;
    // (After a right-aligned number, which ends at its place: at the paragraph's first stop of its own past that
    // place, where it comes before the indent. Word's prints of 18 generated lists, roman numbers i to xviii right-
    // and left-aligned with 180, 270 and 720 twips of hanging indent: a right-aligned number ends at its place
    // whatever its width, its text at the indent, or at a stop half way to it. A left-aligned one's text goes by
    // where the number ends: lhfix.js wideLabels.)
    const endsAtStart = !!lvl && (lvl.jc === 'right' || lvl.jc === 'end');
    const ownFirst = endsAtStart && hang > 0 && own.length ? Math.min(...own) : null;
    const gap = hidden || spaced || (lvl && lvl.suff === 'nothing') ? 0 : ownFirst != null && ownFirst < start + hang ? ownFirst - start : viaIndent;
    style.push(`margin-inline-start:${twipsToPt(left)}pt`, `text-indent:${-twipsToPt(hang)}pt`, `--odx-hang:${twipsToPt(gap)}pt`);
    // A label wider than that room: the text goes on from the next tab stop after the label's end, the paragraph's own
    // or, past its last, a default one (lhfix.js wideLabels measures the label and sets the room; where the label
    // starts, the room, the default stops' pitch and the paragraph's stops, twips). Word's prints of 20 generated
    // documents: a label of 83px with hanging indents of 0 to 720 twips, left-aligned and centred, with a stop of the
    // paragraph's own and none.
    if (gap > 0 && !(lvl && (lvl.jc === 'right' || lvl.jc === 'end' || lvl.jc === 'center'))) attrs['data-ltab'] = [start, gap, dts, ...own.sort((x, y) => x - y)].join(',');
    // A right-aligned number ends at its position, not starts there (roman "i." / "ii." levels): drawn in its box
    // that much to the left, the text still at the indent.
    if (lvl && (lvl.jc === 'right' || lvl.jc === 'end')) attrs.class += ' odx-lbl-right';
  } else if (node.attrs.numId === '0' && node.attrs.styleId && resolveStyle(session.styles.styles, node.attrs.styleId).pPr.numId != null) {
    // Switching off a style's list also drops the style's indent (Word treats it as the list's).
    style.push(`margin-inline-start:${twipsToPt(node.attrs.view.indLeft ?? 0)}pt`, `text-indent:${twipsToPt(node.attrs.view.indFirst ?? 0)}pt`);
  }
}

/** The line's size: from runs that share a size or font, a leading line break, the paragraph mark, and empty markers. */
export function lineSize(node, session, run, label, attrs, style) {
  // A paragraph opening with a bare line break (no size of its own): Word's first line is as tall as that break, in
  // the style's size, not the text's (a 12.5pt title under an 11pt break, doc 07); the text's runs keep their own.
  const fc = node.firstChild;
  const breakSize = fc && fc.type.name === 'hard_break' && !fc.marks.some((m) => m.type.name === 'size') && Number(session.styleRunProp(node.attrs.styleId, 'size') || session.styles.defaults.rPr.size || 20);
  if (run.size && breakSize && breakSize < Number(run.size)) style.push(sizeCss(breakSize / 2));
  else if (run.size) style.push(sizeCss(run.size / 2));
  // Every run raised too (a header line beside a logo, all its text lifted 3pt): no strut at all, as nothing of Word's line
  // hangs below the baseline (a strut would). Not with a line break (its empty line needs one) or a list label.
  else if (run.strut && run.raised && label == null && !node.content.content.some((c) => c.type.name === 'hard_break')) style.push('font-size:0');
  else if (run.strut && Number(run.strut) < Number(session.styleRunProp(node.attrs.styleId, 'size') || 20)) style.push(sizeCss(run.strut / 2));
  if (run.font) style.push(fontCss(run.font));
  // Word's last line is also as tall as the paragraph mark's own font size, when the mark sets one (a 4pt rule
  // line with a 12pt mark is 12pt tall; a mark that only inherits its style's size doesn't count).
  // The mark is a zero-width character at the end, so it counts on the last line only.
  // (Only when there's no visible text: under real text Word keeps the text's own line height.)
  const own = run.size && node.textContent && markSize(node);
  // (A line of spaces alone takes an inherited mark size too: a 4pt space above a rule, 12pt mark, is a 12pt line.)
  const blank = run.size && node.textContent && !/[^\s\u200b]/.test(node.textContent);
  const mark = blank && (own || Number(session.styleRunProp(node.attrs.styleId, 'size') || 20));
  // Under a leading break's smaller strut, a closing break's empty line is the mark's size again (doc 07: 11pt, 12.5pt
  // text, 12.5pt last line), not the strut's.
  const tail = run.size && breakSize && breakSize < Number(run.size) && node.lastChild?.type.name === 'hard_break' && (markSize(node) || Number(run.size));
  // A line of spaces over a smaller mark, or of spaces in several sizes: Word's line is the mark's alone (Word: an 18pt
  // space over a 12pt mark, a 12pt line). Plain spaces only: a no-break space is a letter to Word, sizing its line (a 12pt
  // one over an 11pt mark: a 12pt line, the samples' pasted mail). The paragraph takes the mark's size, and its larger
  // spaces no box (mixlines.js blankOnly).
  const spaces = node.textContent && !/[^ \u200b]/.test(node.textContent) && node.content.content.every((c) => c.isText);
  // (A mark of no size of its own over runs of none either: the size they share already, a table style's in a cell too.
  // Set from the paragraph style instead, a cell's lines of a space in a 9pt table style had come out 11pt.)
  const markAll = spaces && (markSize(node) || (run.size ? Number(session.styleRunProp(node.attrs.styleId, 'size') || session.styles.defaults.rPr.size || 20) : null));
  if (markAll && (!run.size || markAll < Number(run.size))) style.push(sizeCss(markAll / 2));
  if (tail && tail > breakSize && !(mark && mark > run.size)) {
    style.push(`--odx-mark:${tail / 2}pt`);
    attrs.class = [attrs.class, 'odx-mark odx-tailmark'].filter(Boolean).join(' ');
  } else if (mark && mark > run.size) {
    style.push(`--odx-mark:${mark / 2}pt`);
    attrs.class = [attrs.class, 'odx-mark'].filter(Boolean).join(' ');
  }
  // Ending in a line break: its empty last line is as tall as the paragraph's font (editor.css odx-endbreak; textspans.js).
  // (Not a page break, drawn as an element the browser lays out on its own: its paragraph mark stays on the page before.
  // Unless the document asks for Word 2003's rule (w:splitPgBreakAndParaMark): there the mark goes to the next page,
  // an empty line at its top, as in Word.)
  else if (node.lastChild?.type.name === 'hard_break' && (node.lastChild.attrs.type !== 'page' || session.compat?.has('splitPgBreakAndParaMark'))) {
    attrs.class = [attrs.class, 'odx-endbreak'].filter(Boolean).join(' ');
    // (That line holds only the paragraph mark, so it is the mark's size, its own or else the style's, not the break's or
    // the text's: a 14pt title closing with a break over an unsized mark in a 12pt style stood 3px too tall.)
    const end = markSize(node) || Number(session.styleRunProp(node.attrs.styleId, 'size') || session.styles.defaults.rPr.size || 20);
    if (end !== Number(run.size || 0)) style.push(`--odx-endmark:${end / 2}pt`);
    // (In the mark's font too, its own else the style's: 14pt Calibri closing with a break over a mark in the style's
    // Times New Roman, 1.1px a paragraph too tall in Word's prints of 48 generated documents.)
    const mk = (node.attrs.extra || []).find((e) => e[0] === 'rPr')?.[1] || '';
    const ofTheme = (t) => (t && session.theme ? session.theme[t.startsWith('major') ? 'major' : 'minor'] : null);
    const st = session.styleRunProp(node.attrs.styleId, 'themeFont');
    const mf = /\bw:ascii="([^"]+)"/.exec(/<w:rFonts\b[^>]*>/.exec(mk)?.[0] || '')?.[1] || ofTheme(/\bw:asciiTheme="([^"]+)"/.exec(/<w:rFonts\b[^>]*>/.exec(mk)?.[0] || '')?.[1])
      || session.styleRunProp(node.attrs.styleId, 'font') || ofTheme(st) || session.styles.defaults.rPr.font || ofTheme(session.styles.defaults.rPr.themeFont);
    // (The line's height is a multiple of its size, whatever the font: scaled by the two fonts' single lines, under
    // "auto" spacing only, an exact or at-least line staying as set.)
    const rule = node.attrs.view.lineRule ?? session.styleProp(node.attrs.styleId, 'lineRule') ?? session.styles.defaults.pPr.lineRule;
    const k = mf && run.font && (!rule || rule === 'auto') && String(mf).toLowerCase() !== String(run.font).toLowerCase() ? lineFactor(mf) / lineFactor(run.font) : 1;
    if (k > 0 && Math.abs(k - 1) > 0.002) style.push(`--odx-endfont:${fontStack(mf)}`, `--odx-endk:${+k.toFixed(5)}`);
  }
  // Empty markers alone, or before a bare line break that starts the paragraph: the browser gives that first line no
  // height (ProseMirror puts a zero-size image before a leading break), Word a full line. (A break inside formatting
  // has its span's strut, and its own size.)
  // (A floating picture takes no room on the line either: a paragraph holding only those is an empty line in Word,
  // sample S105's 20.6px under its pictures.)
  const marker = (c) => (c.type.name === 'opaque_inline' && !c.attrs.text && c.attrs.width == null) || (c.type.name === 'image' && /float:|position:\s*absolute/.test(c.attrs.place || ''));
  const lead = node.content.content.find((c) => !marker(c));
  if (node.childCount && (!lead || (lead.type.name === 'hard_break' && !lead.marks.length && !lead.attrs.type && !(lead.attrs.clear && lead.attrs.clear !== 'none')))) {
    attrs.class = [attrs.class, 'odx-markers'].filter(Boolean).join(' ');
  }
}

/** An inline picture on a line with text. `page` is the session's page. */
export function pictureLine(node, page, attrs) {
  // An inline picture on a line with text: Word's line is the picture, then the text's descent and the extra of
  // "multiple" spacing below it (corpus 3: a 40px picture in 11pt Calibri at 1.08, 45.4px), not the picture and the
  // extra alone, as on a line of its own (editor.css odx-pictext).
  // (On the same line: a picture with only a space beside it, its caption after a break, is a line of its own.)
  let img = false;
  let text = false;
  let both = false;
  let after = false; // (text after the last picture in line)
  node.forEach((c) => {
    if (c.type.name === 'hard_break') { img = text = false; return; }
    // (Nor one too wide to share a line, which goes to one of its own: sample S115's 665px screenshot after its title.)
    const col = page.width - page.left - page.right;
    if (c.type.name === 'image' && !/float:|position:|display:block/.test(c.attrs.place || '') && !(Number(c.attrs.width) >= col * 0.8)) { img = true; after = false; }
    else if (c.isText && c.text.trim()) { text = true; after = true; }
    if (img && text) both = true;
  });
  if (both) attrs.class = [attrs.class, 'odx-pictext'].filter(Boolean).join(' ');
  // Its picture's line the paragraph's last (nothing after the picture): the rest of the "multiple" extra under it, which
  // Word puts there (editor.css odx-pictext-end; picline-word: at 1.5 and double spacing the next line 3 and 7pt high).
  if (both && !after) attrs.class = `${attrs.class} odx-pictext-end`;
}

/**
 * A list label's bold, italic, colour and underline, as Word draws its number: the level's own (w:lvl/w:rPr), else
 * the paragraph mark's.
 */
export function labelFormat(node, session, l, label, style) {
  if (label == null) return;
  const lvl = listLevel(session.numbering, l.numId, l.ilvl) || {};
  const mark = (node.attrs.extra || []).find((e) => e[0] === 'rPr')?.[1] || '';
  const markOn = (tag) => { const m = new RegExp(`<w:${tag}(?=[\\s/>])([^>]*)>`).exec(mark); return m ? !/w:val="(?:0|false|off)"/.test(m[1]) : null; };
  const bold = lvl.bold ?? markOn('b');
  const italic = lvl.italic ?? markOn('i');
  const color = lvl.color ?? /<w:color\b[^>]*w:val="([^"]*)"/.exec(mark)?.[1] ?? null;
  const underline = lvl.underline ?? /<w:u\b[^>]*w:val="([^"]*)"/.exec(mark)?.[1] ?? null;
  if (bold != null) style.push(`--odx-lweight:${bold ? 'bold' : 'normal'}`);
  if (italic != null) style.push(`--odx-lstyle:${italic ? 'italic' : 'normal'}`);
  if (hexColor(color)) style.push(`--odx-lcolor:#${color}`);
  // (The automatic colour, set on the mark or the level: black, over the paragraph style's colour, as the runs'.)
  else if (color === 'auto') style.push('--odx-lcolor:var(--odx-autotext, #000)');
  if (underline) style.push(`--odx-ldeco:${underline === 'none' ? 'none' : 'underline'}`);
  // A highlighted paragraph mark highlights the number, and its room up to the text (the mark's shading does not).
  // Word's prints of five generated lists: yellow from the number's place to the text's start, a line tall, under a
  // mark highlighted yellow; none under one shaded yellow, or with only the text's runs highlighted.
  const hl = /<w:highlight\b[^>]*w:val="([A-Za-z]{1,20})"/.exec(mark)?.[1];
  if (hl && hl !== 'none' && highlightColor(hl) !== 'transparent') style.push(`--odx-lbg:${highlightColor(hl)}`);
}

const SYMBOL_ASCENT = 1.00537; // Symbol's Windows ascent, 2059 of 2048 units
const WINGDINGS_ASCENT = 0.8994; // Wingdings', 1842 of 2048
// (Fonts labels are set in whose ascent the text's table, baseline.js, doesn't hold: looked up by name, no prototype.)
const LABEL_ASCENT = Object.assign(Object.create(null), { 'arial black': 1.1006 });

/** A list label's own size and font, and the room a label reaching above the text adds to its line. */
export function labelLine(node, session, l, label, run, style) {
  // A Symbol bullet is taller than the text (ascent 1.005em), and Word's line takes the tallest
  // ascent and descent on it, so the bullet's line comes out ~4% taller.
  const lvlFont = label != null && listLevel(session.numbering, l.numId, l.ilvl)?.font;
  let markFont = null; // the font a label with none of its level's is drawn in, where it isn't the text's
  // With no font of its level's: the paragraph mark's (w:rFonts in the paragraph's own rPr, by name or the theme's),
  // else the paragraph style's, not the font its text's runs are set in. (Word's prints of 12 generated documents: the
  // level, the mark and the runs each with a font of their own or none; a sample's one heading number in Calibri, its
  // mark's font, among 35 in the style's Arial.)
  if (label != null && !lvlFont) {
    const mark = (node.attrs.extra || []).find((e) => e[0] === 'rPr')?.[1] || '';
    const fonts = /<w:rFonts\b[^>]*>/.exec(mark)?.[0] || '';
    const named = /\bw:ascii="([^"]+)"/.exec(fonts)?.[1];
    const themed = /\bw:asciiTheme="([^"]+)"/.exec(fonts)?.[1];
    const st = session.styleRunProp(node.attrs.styleId, 'themeFont');
    const ofTheme = (t) => (t && session.theme ? session.theme[t.startsWith('major') ? 'major' : 'minor'] : null);
    const own = named || ofTheme(themed) || (run.font ? session.styleRunProp(node.attrs.styleId, 'font') || ofTheme(st) || session.styles.defaults.rPr.font || ofTheme(session.styles.defaults.rPr.themeFont) : null);
    if (own) { style.push(`--odx-lfont:${fontStack(own)}`); markFont = own; }
  }
  // The label is in the level's own size, else the paragraph mark's, else the style's (12pt bullets over 10pt code
  // lines make Word's bullet lines taller).
  const lblSize = label != null && (listLevel(session.numbering, l.numId, l.ilvl)?.size || markSize(node) || session.styleRunProp(node.attrs.styleId, 'size') || 20);
  // (Also smaller than text in its style's size, no run setting one: a 10pt Symbol bullet over 12pt Times adds nothing
  // to Word's line, its top under the text's; drawn at the text's size it made each item 1.2px taller.)
  // (A size the level sets itself, any bullet font, where no run sets one: Courier New's "o" in 10pt over 12pt text too,
  // samples 1203 and 3751. A size from the paragraph mark for a Symbol bullet: a 10pt mark under 12pt or 11pt Times,
  // unsized, adds nothing to Word's line (a new Word test; samples 3128, 3851), nor under runs of mixed sizes, which
  // leave run.size empty (sample 1687). A number in a smaller size keeps the old drawing, unmeasured in Word.)
  let sized = false;
  node.descendants((c) => { if (!sized && c.isText && c.marks.some((m) => m.type.name === 'size')) sized = true; return !sized; });
  const lvlSize = label != null && Number(listLevel(session.numbering, l.numId, l.ilvl)?.size || 0);
  const styleSize = Number(session.styleRunProp(node.attrs.styleId, 'size') || session.styles.defaults.rPr.size || 20);
  const symbol = /^symbol/i.test(lvlFont || '');
  const smaller = lvlSize ? !sized && lvlSize < styleSize : symbol && Number(lblSize) < styleSize;
  // (And a size the level sets that is larger than its style's, whatever its runs set: Word draws the number at the
  // level's size. Its prints of 24 generated lists: a 12pt and a 14pt "Item 1." from their level over unsized text in
  // an 11pt document, and over 12pt runs; the line was as tall as Word's already, the number drawn at the text's size.)
  // (Not for a paragraph with no text whose number has a font of its own: there the room for that font's number is
  // added already, below, and the size taken as well made the line taller than Word's: a sample's empty "a." in
  // 12pt Times New Roman stood 1.5px over the lines after it. Word's prints of 16 generated documents, empty numbered
  // paragraphs under a 9, 12 or 14pt level: with a font of its own within 1px so; in the paragraph's font the 12pt
  // ones are 1.4px off with the size taken and 5.9px without, the 14pt ones 3.8px and 16px: Word's line grows above
  // the number only, which is still to be matched.)
  const larger = !!lvlSize && !run.size && lvlSize > styleSize && (!lvlFont || /\S/.test(node.textContent));
  if (lblSize && (run.size ? Number(lblSize) !== Number(run.size) : smaller || larger)) style.push(`--odx-lsize:${lblSize / 2}pt`);
  // The label's own letter-spacing, its level's (w:spacing in the level's rPr; the same 24 lists, condensed by 0.3 and
  // 1.1pt and expanded by 1pt, each number's last character within 0.2px of Word's): a converted document's numbers,
  // each level condensed by a point or so, were 1.5px a character wider than Word's.
  const lblSpacing = label != null ? Number(listLevel(session.numbering, l.numId, l.ilvl)?.spacing || 0) : 0;
  if (lblSpacing) style.push(`--odx-lsp:${+(lblSpacing / 20).toFixed(2)}pt`);
  // (A symbol font's bullet the editor has no plain character for, still its private-use code, is drawn in that font
  // too: a Wingdings 3 arrow had come out as the text font's empty box.)
  if (lvlFont && (!/^(symbol|wingdings)/i.test(lvlFont) || /[-]/.test(String(label || '')))) style.push(`--odx-lfont:${fontStack(lvlFont)}`);
  // A label in a font of its own (a Symbol bullet, a bullet or number set in Segoe UI over Calibri text): Word's line is
  // taller by what the label's top reaches above the text's, its font's ascent with its external leading (baseline.js
  // wordAscent) at the label's size; the label's descent never counts, and line spacing doesn't multiply it. Word's
  // PDFs of 368 generated lists, each within 0.1px: bullets, "o" and numbers in Calibri, Segoe UI, Courier New, Times
  // New Roman, Arial and Verdana over 12pt Times New Roman, Calibri and Arial, the label the text's size, 14pt (by its
  // level or the paragraph mark) and 10pt; Symbol bullets over ten fonts at 9 to 12pt, the bullet the text's size, 2pt
  // larger and 2pt smaller; single, 1.15 and double spacing. A Symbol bullet over 12pt Calibri is 0.85px a line, over
  // Times New Roman 1.15px, Courier New 2.77px, Verdana and Segoe UI nothing; a Segoe UI bullet over Calibri 2.03px; a
  // Courier New "o" nothing until it is larger than the text. The room goes above the first line's text (editor.css
  // odx-first), so wrapped lines keep the text's own height. A Wingdings bullet (ascent 0.899em) is under every text
  // font's at the text's size and adds nothing; 14pt over 12pt Calibri it adds 1.55px (a sample's Word PDF: 21.1px lines).
  const raise = label != null ? ((listLevel(session.numbering, l.numId, l.ilvl)?.position || 0) / 2) * (4 / 3) : 0;
  if (raise) style.push(`--odx-lraise:${+raise.toFixed(2)}px`);
  // (A label with no font of its level's counts the same, in the font it is drawn in: the paragraph mark's, else the
  // style's. Word's prints of 288 generated lists: a number in a 12 or 14pt level over 10 or 12pt Arial, Calibri and Times
  // New Roman, the item's line taller by exactly this; 9pt and equal sizes add nothing. A form's numbered headings, a 12pt
  // level over 10pt text, stood 2.6px high each.)
  if (label != null) {
    const tf = session.styleRunProp(node.attrs.styleId, 'themeFont');
    const fname = run.font || session.styleRunProp(node.attrs.styleId, 'font') || (tf && session.theme[tf.startsWith('major') ? 'major' : 'minor']) || session.theme.minor || 'calibri';
    const lblFont = lvlFont || markFont || fname;
    const dings = /^wingdings$/i.test(lblFont);
    const la = symbol ? SYMBOL_ASCENT : dings ? WINGDINGS_ASCENT : wordAscent(String(lblFont).toLowerCase()) || LABEL_ASCENT[String(lblFont).toLowerCase()];
    // (Text in a font whose ascent isn't known, under a Symbol bullet: its line less Calibri's descent, as fitted before.)
    const ta = wordAscent(String(fname).toLowerCase()) || (symbol ? Math.max(lineFactor(fname) - 0.2686, 0.93) : 0);
    const rule = node.attrs.view.lineRule ?? session.styleProp(node.attrs.styleId, 'lineRule') ?? session.styles.defaults.pPr.lineRule;
    // (Lines sized from their runs: an exact line stays as set. A text font's label under "at least" isn't measured.)
    if (la && ta && (symbol || dings ? rule !== 'exact' : !rule || rule === 'auto')) {
      const lbl = Number(lblSize) || Number(run.size) || 20;
      // (The text's size: its runs' one size, else its first text's own, where they set theirs one by one, else the
      // style's. A 12pt "(a)" before 12pt runs in an 11pt style, one run among them unsized, was counted as larger than
      // its text and given 1.25px of room above its line that Word doesn't add.)
      let firstSize = 0;
      if (!run.size) node.descendants((c) => { if (!firstSize && c.isText && c.text.trim()) firstSize = Number(c.marks.find((m) => m.type.name === 'size')?.attrs.val) || -1; return !firstSize; });
      // (Where it is larger than the style's, that is: under smaller text the line is the style's size already, an
      // unsized space or the mark its own, and room added for the number made it taller than Word's. Cells of numbered
      // 10pt text with one unsized space in an 11pt style stood 0.7px low with it.)
      const styleSize = Number(session.styleRunProp(node.attrs.styleId, 'size')) || Number(session.styles.defaults.rPr.size) || 0;
      const txt = Number(run.size) || (firstSize > styleSize ? firstSize : 0) || styleSize || lbl;
      // A label its level raises or lowers (w:position) is drawn that far off the baseline, and reaches that much
      // further up, or less far (48 generated lists: labels raised 1.5 to 6pt and lowered 1.5 and 4pt; a lowered
      // label never makes its line taller below).
      // (Runs setting fonts and sizes of their own, no one for the paragraph: from the first text's, the one the room is
      // put over, editor.css; a taller run later on the line stands above that room by itself. Word's prints of 24
      // generated lists, a Symbol bullet over 11pt Calibri and 10pt Arial in a 12pt Times New Roman style: the line is the
      // tallest top of bullet and runs and their deepest foot; against the style's Times it had no room and stood 0.75px
      // short. mixlines.js gives the runs their boxes.)
      let topA = ta, topSize = txt;
      if (!run.size || !run.font) {
        let seen = false;
        node.descendants((c) => {
          if (seen || !c.isText || !c.text.trim()) return !seen;
          seen = true;
          const a = wordAscent(String(c.marks.find((m) => m.type.name === 'font')?.attrs.val || fname).toLowerCase());
          if (a) { topA = a; topSize = Number(c.marks.find((m) => m.type.name === 'size')?.attrs.val) || styleSize || lbl; }
          return false;
        });
      }
      const up = (la * lbl - topA * topSize) / 2 * (4 / 3) + raise;
      // (And as Word has it, from the fonts' own units, unrounded: what a page's foot is fitted by, pages.js roundOff.)
      const units = (v) => Math.round(v * 2048) / 2048;
      const upExact = (units(la) * lbl - units(topA) * topSize) / 2 * (4 / 3) + raise;
      // (In px, to the nearest of the browser's 1/64px: it cuts a length down to that, a hundredth of a pixel a line.)
      if (up > 0.02) style.push(`--odx-lbl-own:${Math.round(up * 64) / 64}px`, `--odx-lbl-up:${+upExact.toFixed(5)}px`);
    }
  }
}

/** Deleted whole, text and mark (markers that show nothing, such as bookmarks or empty content controls, don't count). */
export function deletedWhole(node) {
  if (paragraphMarkRev(node.attrs) !== 'del') return false;
  let all = true;
  node.forEach((c) => { if (!c.marks.some((m) => m.type.name === 'del') && (c.isText ? c.text.trim() : c.type.name !== 'opaque_inline' || c.attrs.width != null)) all = false; });
  return all;
}

/** True when a paragraph was inserted whole: its mark and all its text (the Original view leaves it out). */
export function insertedWhole(node) {
  if (paragraphMarkRev(node.attrs) !== 'ins') return false;
  let all = true;
  node.forEach((c) => { if (!c.marks.some((m) => m.type.name === 'ins') && (c.isText ? c.text.trim() : c.type.name !== 'opaque_inline' || c.attrs.width != null)) all = false; });
  return all;
}

/**
 * Tracked changes' classes: the change bar, the paragraph mark's revision, a paragraph deleted whole, all text deleted,
 * a paragraph inserted whole.
 */
export function changeClasses(node) {
  // Tracked changes: a change bar in the margin, and a deleted or inserted paragraph mark.
  const changed = paragraphChanged(node) && ['odx-changed'];
  const pmark = paragraphMarkRev(node.attrs);
  // A paragraph that is deleted whole (text and mark) disappears in simple markup, as in Word.
  const allDel = deletedWhole(node);
  // All its text deleted, its mark kept: an empty line in simple markup (editor.css: odx-del-text), as in Word (doc 25's
  // struck paragraphs keep their 240-twip spacing and a line).
  let delText = !allDel && pmark !== 'del' && !!node.textContent.trim();
  if (delText) node.forEach((c) => { if (c.isText && c.text.trim() && !c.marks.some((m) => m.type.name === 'del')) delText = false; });
  return [...(changed || []), pmark && `odx-pmark-${pmark}`, allDel && 'odx-all-del', delText && 'odx-del-text', insertedWhole(node) && 'odx-all-ins'];
}

/** A heading to screen readers: its outline level (its own or its style's; Word's built-in Heading 1-9 have theirs). */
export function headingRole(node, session, attrs) {
  // A heading (an outline level, its own or its style's) is one to screen readers: NVDA read "Quarterly Report", not
  // "heading level 1". (Not an empty one: nothing to read.)
  // (Word's built-in Heading 1-9 have their level by name where styles.xml doesn't give it; one used without being
  // defined is Normal in Word, no heading.)
  const outline = outlineLevel(session, node);
  if (outline != null && outline >= 0 && outline < 9 && node.textContent.trim()) { attrs.role = 'heading'; attrs['aria-level'] = String(outline + 1); }
}
