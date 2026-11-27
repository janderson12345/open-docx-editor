// Builds test/fixtures/aptos-narrow.docx: two paragraphs in Aptos Narrow (a cloud font browsers rarely have) at 12pt,
// four lines or so each. Word sets them at Aptos Narrow's own widths; drawn in the editor's Arial Narrow stand-in they
// were 6.5% narrower and every line took a word or two more than Word's (features.mjs: their lines end on Word's words,
// with the open-metric-fonts stand-in, which has Aptos Narrow's widths).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const R = '<w:rPr><w:rFonts w:ascii="Aptos Narrow" w:hAnsi="Aptos Narrow"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>';
const P = (t) => `<w:p><w:pPr>${R}</w:pPr><w:r>${R}<w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const xml = [
  P('The paragraph begins simply and carries some text for the layout, checks, including but not limited to lines, measures, spacing, punctuation and the trailing that ends with each of them, in keeping with the examples quoted at the close of the year.'),
  P('Afterward: the paragraphs might provide additional wording, simple lists and illustrated filler as following in example 1.0, sometimes on the right of the ruler, and might return to settings after two words above and one at the period.'),
].join('');
const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${xml}$1`));
fs.writeFileSync(new URL('./aptos-narrow.docx', import.meta.url), zipSync(files));
