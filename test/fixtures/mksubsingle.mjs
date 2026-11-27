// Builds test/fixtures/sub-single.docx: paragraphs of text spaced below a single line (w:line under 240, auto). Word
// makes each line that multiple of the font's single line, shorter than its text. Single-line and two-line paragraphs
// at 0.5, 0.75 and 0.9 lines in Calibri 11pt, then a Bold Calibri Light 16pt block at 0.5 lines with 12pt before each
// paragraph (a heading style's spacing), then plain single-spaced lines below them to show where the page's text goes on.
// WORD_LINES, in features.mjs, holds where Word's PDF of it puts each line.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const LONG = 'Pack my box with five dozen liquor jugs; sphinx of black quartz, judge my vow, and the jaunty zebra waltzes past several heavy boxes of grapes.';
const P = (text, { line, before = 0, font = 'Calibri', size = 11, b = false } = {}) => `<w:p><w:pPr><w:spacing w:before="${before}" w:after="0"${line ? ` w:line="${line}" w:lineRule="auto"` : ''}/></w:pPr>`
  + `<w:r><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>${b ? '<w:b/>' : ''}<w:sz w:val="${size * 2}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const xml = [
  P('Single A', { line: 240 }),
  ...[120, 180, 216].flatMap((line) => [P(`Sub ${line} one`, { line }), P(`Sub ${line} two`, { line }), P(`${LONG} (${line})`, { line })]),
  P('Single B', { line: 240 }),
  ...['Head one', 'Head two', 'Head three'].map((t) => P(t, { line: 120, before: 240, font: 'Calibri Light', size: 16, b: true })),
  P('Single C', { line: 240 }),
  P('Single D', { line: 240 }),
].join('');
const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${xml}$1`));
fs.writeFileSync(new URL('./sub-single.docx', import.meta.url), zipSync(files));
