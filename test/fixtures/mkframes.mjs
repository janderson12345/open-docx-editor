// Builds test/fixtures/frames.docx: paragraphs in frames (w:framePr) as Word places them: frame A set down 2in from the
// page's top and 4.5in from its left edge, 2in wide, two paragraphs of one frame (the same w:framePr), the text going
// around it; frame B at the right of the text where it comes, 1.5in wide, the paragraphs after it beside it; frame C a
// note in the left margin, beside the paragraph after it.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const R = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const P = (t, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${R(t)}</w:p>`;
const body = (n) => P(`Body ${n}. The quick brown fox jumps over the lazy dog while seven stored procedures were rewritten to remove a predicate that forced a full scan of an eleven million row table.`);
const frameA = '<w:framePr w:w="2880" w:hSpace="180" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="6480" w:y="2880"/>';
const frameB = '<w:framePr w:w="2160" w:hSpace="180" w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="right" w:y="1"/>';
const frameC = '<w:framePr w:w="1080" w:hSpace="180" w:wrap="around" w:vAnchor="text" w:hAnchor="page" w:x="180" w:y="1"/>';
const boxed = '<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="4" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="4" w:color="auto"/></w:pBdr>';

const xml = [
  body(1),
  P('Frame A, first paragraph: set down on the page.', frameA + boxed),
  P('Frame A, second paragraph, in the same frame.', frameA + boxed),
  body(2), body(3), body(4), body(5),
  P('Frame B: a sidebar at the right of the text.', frameB),
  body(6), body(7),
  P('Note C in the margin.', frameC),
  body(8), body(9),
].join('');
const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${xml}$1`));
fs.writeFileSync(new URL('./frames.docx', import.meta.url), zipSync(files));
