// Builds test/fixtures/semibold-bold.docx: Bold over the Microsoft 365 cloud SemiBold and Demi families, whose Bold Word
// fakes (they have no bold face). One wrapping paragraph each: Aptos SemiBold in Regular, Bold, Italic and Bold Italic
// at 11pt, Bold at 16pt, and the other three families (Avenir Next LT Pro Demi, Georgia Pro Semibold, Verdana Pro
// Semibold) in Bold. WORD_LINES, in features.mjs, holds the last word of each line in Word's PDF of it.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const TEXT = 'The quick brown fox jumps over the lazy dog while seven stored procedures were rewritten by Judge Kelly, '
  + 'who typed 0123456789 twice. Pack my box with five dozen liquor jugs; sphinx of black quartz, judge my vow. '
  + 'Bright vixens jump, dozy fowl quack, and the jaunty zebra waltzes past several heavy boxes of mixed grapes.';
const P = (font, size, b, i) => `<w:p><w:r><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>${b ? '<w:b/>' : ''}${i ? '<w:i/>' : ''}<w:sz w:val="${size * 2}"/></w:rPr><w:t xml:space="preserve">${TEXT}</w:t></w:r></w:p>`;
const xml = [
  P('Aptos SemiBold', 11, 0, 0),
  P('Aptos SemiBold', 11, 1, 0),
  P('Aptos SemiBold', 11, 0, 1),
  P('Aptos SemiBold', 11, 1, 1),
  P('Aptos SemiBold', 16, 1, 0),
  P('Avenir Next LT Pro Demi', 11, 1, 0),
  P('Georgia Pro Semibold', 11, 1, 0),
  P('Verdana Pro Semibold', 11, 1, 0),
].join('');
const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${xml}$1`));
fs.writeFileSync(new URL('./semibold-bold.docx', import.meta.url), zipSync(files));
