// The paragraph pass's rules (editor.js paragraphPlugin): each sets classes and style on one paragraph, from its node,
// the session and what the pass worked out before it. `attrs` is the paragraph's decoration (class, data-label, role),
// `style` its list of CSS declarations; both are added to in the pass's order.
import { listLevel, resolveStyle } from '../ooxml/styles.js';
import { twipsToPt, fontCss, fontStack, lineFactor, mergeTabs, sizeCss, hexColor } from '../ooxml/props.js';
import { paragraphChanged, paragraphMarkRev } from './track.js';
import { outlineLevel } from './commands.js';

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
    const left = node.attrs.view.indLeft ?? ind.indLeft ?? 720;
    const hang = -(node.attrs.view.indFirst ?? ind.indFirst ?? -360);
    attrs.class = lvl && lvl.fmt === 'bullet' ? 'odx-li odx-bullet' : 'odx-li';
    // A number formatted as hidden text (w:vanish on its level, or on the paragraph mark it takes its look from) isn't
    // shown, nor the room after it, as in Word.
    const mark = (node.attrs.extra || []).find((e) => e[0] === 'rPr');
    const hidden = !!lvl?.hidden || (!!mark && /<w:vanish(?=[\s/>])(?![^>]*w:val="(?:0|false|off)")/.test(mark[1]));
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
    const gap = hidden || (lvl && lvl.suff === 'nothing') ? 0 : viaIndent;
    style.push(`margin-inline-start:${twipsToPt(left)}pt`, `text-indent:${-twipsToPt(hang)}pt`, `--odx-hang:${twipsToPt(gap)}pt`);
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
  else if (node.lastChild?.type.name === 'hard_break' && (node.lastChild.attrs.type !== 'page' || session.compat?.has('splitPgBreakAndParaMark'))) attrs.class = [attrs.class, 'odx-endbreak'].filter(Boolean).join(' ');
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
  node.forEach((c) => {
    if (c.type.name === 'hard_break') { img = text = false; return; }
    // (Nor one too wide to share a line, which goes to one of its own: sample S115's 665px screenshot after its title.)
    const col = page.width - page.left - page.right;
    if (c.type.name === 'image' && !/float:|position:|display:block/.test(c.attrs.place || '') && !(Number(c.attrs.width) >= col * 0.8)) img = true;
    else if (c.isText && c.text.trim()) text = true;
    if (img && text) both = true;
  });
  if (both) attrs.class = [attrs.class, 'odx-pictext'].filter(Boolean).join(' ');
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
  if (underline) style.push(`--odx-ldeco:${underline === 'none' ? 'none' : 'underline'}`);
}

/** A list label's own size and font, and the room a taller Symbol bullet adds to its line. */
export function labelLine(node, session, l, label, run, style) {
  // A Symbol bullet is taller than the text (ascent 1.005em), and Word's line takes the tallest
  // ascent and descent on it, so the bullet's line comes out ~4% taller.
  const lvlFont = label != null && listLevel(session.numbering, l.numId, l.ilvl)?.font;
  // The label is in the level's own size, else the paragraph mark's, else the style's (12pt bullets over 10pt code
  // lines make Word's bullet lines taller).
  const lblSize = label != null && (listLevel(session.numbering, l.numId, l.ilvl)?.size || markSize(node) || session.styleRunProp(node.attrs.styleId, 'size') || 20);
  if (lblSize && run.size && Number(lblSize) !== Number(run.size)) style.push(`--odx-lsize:${lblSize / 2}pt`);
  if (lvlFont && !/^(symbol|wingdings)/i.test(lvlFont)) style.push(`--odx-lfont:${fontStack(lvlFont)}`);
  // Wingdings (ascent 0.899em, descent 0.211em) fits inside Calibri's line and adds nothing.
  if (lvlFont && /^symbol/i.test(lvlFont)) {
    const tf = session.styleRunProp(node.attrs.styleId, 'themeFont');
    const fname = run.font || session.styleRunProp(node.attrs.styleId, 'font') || (tf && session.theme[tf.startsWith('major') ? 'major' : 'minor']) || session.theme.minor || 'calibri';
    const base = lineFactor(fname);
    // Only the line with the bullet: wrapped lines keep the text's own height (the label is padded above).
    // Aptos lines take a bullet the text's size (sample S101), not a smaller one (a 10pt bullet over 12pt text). The
    // extra goes above the text (sample S093: a heading's first bullet 2px lower than with it below;
    // Fitted to Word's PDFs, not a font rule.
    const lbl = Number(lblSize) || Number(run.size) || 20;
    const txt = Number(run.size) || Number(session.styleRunProp(node.attrs.styleId, 'size')) || lbl;
    const line = node.attrs.view.line ?? session.styleProp(node.attrs.styleId, 'line') ?? session.styles.defaults.pPr.line ?? 240;
    const aptos = /^aptos/i.test(fname);
    // (Verdana's ascent, 1.0054em, already reaches the bullet's: nothing added. file-sample's DejaVu Sans list, set in
    // Verdana: 0.11pt a bullet line over Word's.)
    const tall = /^(verdana|dejavu sans)/i.test(fname);
    let up = (aptos && lbl < txt) || tall ? 0 : (base >= 1.22 ? 1.2737 : Math.max(base, 1.2247)) - base;
    // A bullet larger than the text: its own ascent over the text's descent, in the label's em (sample S071's 12pt
    // Symbol bullets over 11pt Calibri, 24px lines in Word where the text's are 21.3).
    if (up > 0 && !aptos && lbl > txt) up = (1.005 * lbl + (up + base - 1.005) * txt - base * txt) / lbl;
    // (Exact line spacing: nothing makes Word's line taller, the bullet neither.)
    const rule = node.attrs.view.lineRule ?? session.styleProp(node.attrs.styleId, 'lineRule') ?? session.styles.defaults.pPr.lineRule;
    // (Multiple spacing stretches an Aptos bullet line's extra too: 1.25 lines add 1.25 times as much, Product
    // Management. Not Calibri's at 1.15: sample S047, sample S067 and sample S125 lose lines with it.)
    const mult = aptos && (!rule || rule === 'auto') ? Math.max(1, line / 240) : 1;
    if (up > 0 && rule !== 'exact') style.push(`--odx-lbl-up:${+(up * mult).toFixed(4)}em`);
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
