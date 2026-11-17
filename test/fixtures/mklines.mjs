// Builds test/fixtures/lines.docx (Times New Roman: a line exactly 40pt tall over 12pt text; a line of 10pt text with a
// 12pt bold run in it; the same raised 3pt; lines of 10pt text with a 12pt run raised or lowered 3pt, Calibri's too, and at
// 1.15 lines) and lines-large.docx (1,600 one-line paragraphs in 10 to 13pt: a large
// document, whose blocks far off screen skip layout) with lines-small.docx (its first 400).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const F = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>';
const R = (t, rpr = '') => `<w:r><w:rPr>${F}${rpr}</w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
const C = (t, rpr = '') => `<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>${rpr}</w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
const P = (ppr, runs) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0"${ppr}/></w:pPr>${runs}</w:p>`;
const make = (name, body) => {
  const files = unzipSync(blankDocxBytes());
  files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
  fs.writeFileSync(new URL(`./${name}`, import.meta.url), zipSync(files));
};
make('lines.docx', [
  P('', R('Plain before.', '<w:sz w:val="24"/>')),
  P(' w:line="800" w:lineRule="exact"', R('Exactly forty.', '<w:sz w:val="24"/>')),
  P('', R('Mixed small ', '<w:sz w:val="20"/>') + R('Large', '<w:b/><w:sz w:val="24"/>') + R(' small.', '<w:sz w:val="20"/>')),
  P('', R('Raised small ', '<w:sz w:val="20"/>') + R('Large', '<w:b/><w:position w:val="6"/><w:sz w:val="24"/>') + R(' small.', '<w:sz w:val="20"/>')),
  P('', R('Plain after.', '<w:sz w:val="24"/>')),
  // (Lines of runs in two sizes, Word's heights: Times New Roman's 12pt lowered 3pt, Calibri's raised 3pt, single spaced
  // (the blank document's Normal is 1.08), and at 1.15.)
  P(' w:line="240" w:lineRule="auto"', R('Lowered small ', '<w:sz w:val="20"/>') + R('Large', '<w:position w:val="-6"/><w:sz w:val="24"/>') + R(' small.', '<w:sz w:val="20"/>')),
  P(' w:line="240" w:lineRule="auto"', C('Calibri small ', '<w:sz w:val="20"/>') + C('Large', '<w:position w:val="6"/><w:sz w:val="24"/>') + C(' small.', '<w:sz w:val="20"/>')),
  P(' w:line="276" w:lineRule="auto"', R('Spaced small ', '<w:sz w:val="20"/>') + R('Large', '<w:position w:val="6"/><w:sz w:val="24"/>') + R(' small.', '<w:sz w:val="20"/>')),
].join(''));
const sizes = [20, 21, 22, 26];
const many = (n) => Array.from({ length: n }, (_, i) => P('', R(`Line ${i + 1} of the document.`, `<w:sz w:val="${sizes[i % 4]}"/>`))).join('');
make('lines-large.docx', many(1600));
make('lines-small.docx', many(400)); // (the same lines, not a large document: what the large one's should match)
