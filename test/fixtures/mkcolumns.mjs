// Builds test/fixtures/columns-unequal.docx: newspaper columns of their own widths (w:equalWidth="0"): 1.5in, a 0.25in
// gap, 3in, a 0.5in gap and 1.25in, with a line between, and numbered paragraphs running over two pages.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const words = 'The quick brown fox jumps over the lazy dog while seven stored procedures were rewritten to remove a predicate that forced a full scan'.split(' ');
const para = (i) => `<w:p><w:r><w:t xml:space="preserve">P${i} ${words.slice(0, 8 + (i * 7) % 16).join(' ')}.</w:t></w:r></w:p>`;
const cols = '<w:cols w:num="3" w:sep="1" w:space="360" w:equalWidth="0"><w:col w:w="2160" w:space="360"/><w:col w:w="4320" w:space="720"/><w:col w:w="1800"/></w:cols>';

const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
files[main] = strToU8(strFromU8(files[main])
  .replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${Array.from({ length: 70 }, (_, i) => para(i + 1)).join('')}$1`)
  .replace(/<w:cols\b[^>]*\/>/, cols));
fs.writeFileSync(new URL('./columns-unequal.docx', import.meta.url), zipSync(files));
