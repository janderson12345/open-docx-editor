// What a paragraph's text runs share, for sizing its lines.
import { schema } from './schema.js';

/** The font and size shared by every text run of a paragraph (each null if they differ or are unset). */
export function uniformRun(node, charStyle, theme) {
  let size;
  let font;
  // The largest size, while every run sets its own (null once one doesn't); whether every run is raised (w:position).
  let most;
  let raised;
  // (A run of spaces alone beside text doesn't break uniformity: "part 2"'s leading space in the default 11pt Calibri
  // before its 12pt Times, whose lines took Calibri's line factor, 2.7px a line too tall.)
  let text = false;
  node.forEach((c) => { if (c.isText && c.text.trim()) text = true; });
  node.forEach((c) => {
    if (!c.isText || (text && !c.text.trim())) return;
    const s = c.marks.find((m) => m.type === schema.marks.size);
    const f = c.marks.find((m) => m.type === schema.marks.font);
    const sv = s ? s.attrs.val : null;
    // A run naming a theme font (w:asciiTheme, kept raw) is in that font too.
    const tf = !f && theme && c.marks.find((m) => m.type.name === 'rextra')?.attrs.items?.find(([n, x]) => n === 'rFonts' && /w:(?:ascii|hAnsi)Theme="(major|minor)/.test(x));
    const fv = f ? f.attrs.val : tf ? theme[/w:(?:ascii|hAnsi)Theme="(major|minor)/.exec(tf[1])[1]] || null : null;
    most = most === null || sv == null ? null : Math.max(most || 0, Number(sv));
    const up = c.marks.some((m) => m.type.name === 'rextra' && (m.attrs.items || []).some(([n, x]) => n === 'position' && Number((/w:val="(-?\d+)"/.exec(x) || [])[1]) > 0));
    raised = raised !== false && up;
    // A stray Cambria Math character (Word's non-breaking hyphen) doesn't set the line, so it doesn't break uniformity.
    if (fv && /^cambria math$/i.test(fv)) { size = size === undefined || size === sv ? sv : null; return; }
    size = size === undefined || size === sv ? sv : null;
    font = font === undefined || font === fv ? fv : null;
  });
  // A line holding only a VML picture (a horizontal rule) takes its run's size, not the mark's.
  let pict = false;
  // (Not one placed apart from the text, position:absolute: it holds no line, so the mark sizes the line, sample S038 form.)
  if (size === undefined) node.descendants((c) => { if ((c.type.name === 'opaque_inline' || (c.type.name === 'image' && /o:hr="t"/.test(c.attrs.xml))) && /<w:pict\b/.test(c.attrs.xml) && !/position:\s*absolute/.test(c.attrs.xml)) pict = true; });
  if (size === undefined && !pict) {
    // No text: Word sizes the line from the paragraph mark's run properties (its size and its font).
    const mark = (node.attrs.extra || []).find((e) => e[0] === 'rPr');
    const sz = mark && /<w:sz w:val="(\d+)"/.exec(mark[1]);
    // The mark's character style (w:rStyle) too: an empty paragraph whose mark is "Intense Emphasis" is a 14pt line.
    const rs = mark && charStyle && /<w:rStyle w:val="([^"]+)"/.exec(mark[1]);
    const cs = rs ? charStyle(rs[1]) || {} : {};
    if (sz) size = String(sz[1]);
    else if (cs.size) size = String(cs.size);
    const rf = mark && /<w:rFonts\b[^>]*\bw:ascii="([^"]+)"/.exec(mark[1]);
    if (rf && font === undefined) font = rf[1];
    else if (cs.font && font === undefined) font = cs.font;
  }
  // Runs of different sizes, all set: the paragraph's own size (its style's) is in none of them, so Word's lines never
  // take it; the largest run's size stands in for it (a 14pt header style over 12pt and 10pt runs beside a logo).
  return { size: size || null, font: font || null, strut: !size && most ? String(most) : null, raised: !!raised };
}
