// Builds test/fixtures/table-row-notes.docx, table-row-band.docx, table-row-foot.docx, notes-mixed-font.docx and
// notes-separator-16pt.docx: table rows and footnotes as Word lays them (its prints of the same documents made in Word).
// table-row-notes: a 40-row table, a footnote reference in each row: each note at the foot of its own row's page (Word:
//   rows 1-15 and their notes on page 1, 16-30 on page 2, 31-40 on page 3).
// table-row-band: twelve rows, a 3pt top line over the first cell (one line of text) and 0.5pt over the second (three
//   lines): both cells' text 3pt down, the rows 62.67px apart (three 12pt Calibri lines and the 3pt line).
// table-row-foot: a page of exact-height text leaving 33.75pt above the foot, then a row of at least 32.25pt (four lines)
//   under 1pt lines over and under it: Word moves it whole to page 2 (it splits only from 34.25pt of room).
// notes-mixed-font: footnotes in a Times New Roman style, their text in Calibri 10pt after a Times quote mark: Word's
//   lines 16.28px apart (Calibri's line; the reference mark and the quote don't size it).
// notes-separator-16pt: Normal in Calibri 16pt, single spaced; 60 body lines of exactly 20pt in a style of their own and
//   a one-line note: page 1 keeps 30 of them (under Normal in 12pt, 31): the separator is a line of Normal.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SINGLE = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
const seps = '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr>' + SINGLE + '</w:pPr><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr>' + SINGLE + '</w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
const run = (t, rpr = '') => `<w:r><w:rPr>${rpr}</w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
const ref = (id) => `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
const CAL = (sz) => `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`;
const noteStyles = (font) => `<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:pPr>${SINGLE}</w:pPr><w:rPr>${font}<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>`
  + '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>';

/** A document from the blank one: its body, Normal's run and paragraph properties, more styles, its footnotes. */
function write(name, { body, normal = `<w:pPr>${SINGLE}</w:pPr><w:rPr>${CAL(24)}</w:rPr>`, styles = '', notes = null }) {
  const files = unzipSync(blankDocxBytes());
  const main = 'word/document.xml';
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
  let st = strFromU8(files['word/styles.xml']);
  st = st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, `$1${normal}$2`);
  files['word/styles.xml'] = strToU8(st.replace('</w:styles>', `${styles}</w:styles>`));
  if (notes) {
    files['word/footnotes.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${W}>${seps}${notes}</w:footnotes>`);
    files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'));
    const rels = 'word/_rels/document.xml.rels';
    files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>'));
  }
  fs.writeFileSync(new URL(`./${name}`, import.meta.url), zipSync(files));
}
const para = (content, ppr = SINGLE) => `<w:p><w:pPr>${ppr}</w:pPr>${content}</w:p>`;
const grid = (ws) => `<w:tblGrid>${ws.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
const line = (side, sz) => (sz ? `<w:${side} w:val="single" w:sz="${sz}" w:space="0" w:color="auto"/>` : `<w:${side} w:val="nil"/>`);
const cell = (w, content, { top = 0, bottom = null, valign = '' } = {}) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/><w:tcBorders>${line('top', top)}${bottom == null ? '' : line('bottom', bottom)}</w:tcBorders>${valign ? `<w:vAlign w:val="${valign}"/>` : ''}</w:tcPr>${content}</w:tc>`;
const ALL = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';

// table-row-notes
const noteText = (i) => `Note ${i} is long enough to take two lines at the foot of the page, so that the notes of a page take some room there and move the rows.`;
write('table-row-notes.docx', {
  body: para(run('Table test.')) + `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>${ALL}</w:tblPr>${grid([4680, 4680])}`
    + Array.from({ length: 40 }, (_, i) => `<w:tr>${cell(4680, para(run(`Row ${i + 1}`)))}${cell(4680, para(run(`Text ${i + 1}`) + ref(i + 1)))}</w:tr>`).join('')
    + '</w:tbl>' + para(''),
  styles: noteStyles(CAL(20)),
  notes: Array.from({ length: 40 }, (_, i) => `<w:footnote w:id="${i + 1}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>${run(' ' + noteText(i + 1))}</w:p></w:footnote>`).join(''),
});

// table-row-band
write('table-row-band.docx', {
  body: para(run('Table test.')) + `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr>${grid([4680, 4680])}`
    + Array.from({ length: 12 }, (_, i) => `<w:tr>${cell(4680, para(run(`Alpha${i + 1}`)), { top: 24 })}${cell(4680, ['one', 'two', 'three'].map((w) => para(run(`B${i + 1} ${w}`))).join(''), { top: 4 })}</w:tr>`).join('')
    + '</w:tbl>' + para(''),
});

// table-row-foot: 648pt of text height; 33.75pt left (the row's 32.25pt and 1.5pt, less than its lines' 2pt)
const exact = (pt) => `<w:spacing w:before="0" w:after="0" w:line="${pt * 20}" w:lineRule="exact"/>`;
write('table-row-foot.docx', {
  body: para(run('Lead'), exact(648 - 33.75)) + `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr>${grid([4680, 4680])}`
    + [1, 2].map((r) => `<w:tr><w:trPr><w:trHeight w:val="645"/></w:trPr>${cell(4680, para(run(`Side${r}`)), { top: 8, valign: 'center' })}${cell(4680, ['first', 'second', 'third', 'fourth'].map((w) => para(run(`R${r} ${w}`))).join(''), { top: 8 })}</w:tr>`).join('')
    + '</w:tbl>' + para(''),
});

// notes-mixed-font
const TNR = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>';
write('notes-mixed-font.docx', {
  body: para(run('A paragraph with three notes.') + ref(1) + ref(2) + ref(3)),
  normal: `<w:pPr>${SINGLE}</w:pPr><w:rPr>${TNR}<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>`,
  styles: noteStyles(TNR),
  notes: [1, 2, 3].map((i) => `<w:footnote w:id="${i}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>${run(' “')}${run(`Quoted” note ${i}, set in Calibri at ten points, long enough to wrap over three lines of the note at the foot of the page, with a few more words to be sure of it.`, '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>')}</w:p></w:footnote>`).join(''),
});

// notes-separator-16pt
write('notes-separator-16pt.docx', {
  body: Array.from({ length: 60 }, (_, i) => `<w:p><w:pPr><w:pStyle w:val="BodyX"/></w:pPr>${run(`Body line ${i + 1}`)}${i === 0 ? ref(1) : ''}</w:p>`).join(''),
  normal: `<w:pPr>${SINGLE}</w:pPr><w:rPr>${CAL(32)}</w:rPr>`,
  styles: `<w:style w:type="paragraph" w:customStyle="1" w:styleId="BodyX"><w:name w:val="BodyX"/><w:pPr>${exact(20)}</w:pPr><w:rPr>${CAL(24)}</w:rPr></w:style>` + noteStyles(CAL(20)),
  notes: `<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>${run(' The note, one line.')}</w:p></w:footnote>`,
});
