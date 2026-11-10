// Builds test/fixtures/original.docx: tracked changes of each kind, for the Original view (the document before them):
// text inserted and deleted in a paragraph, a paragraph inserted whole and one deleted whole, a paragraph split by an
// inserted paragraph mark, and two joined by a deleted one.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

let id = 1;
const who = () => `w:id="${id++}" w:author="Reviewer" w:date="2026-11-01T10:00:00Z"`;
const R = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const INS = (t) => `<w:ins ${who()}>${R(t)}</w:ins>`;
const DEL = (t) => `<w:del ${who()}><w:r><w:delText xml:space="preserve">${t}</w:delText></w:r></w:del>`;
const mark = (kind) => `<w:pPr><w:rPr><w:${kind} ${who()}/></w:rPr></w:pPr>`;
const P = (body, ppr = '') => `<w:p>${ppr}${body}</w:p>`;

const xml = [
  P(R('Alpha ') + INS('new ') + R('words ') + DEL('old ') + R('end.')),
  P(INS('A paragraph inserted whole.'), mark('ins')),
  P(DEL('A paragraph deleted whole.'), mark('del')),
  P(R('Split here') + INS(' by Enter'), mark('ins')),
  P(R(' and joined again.')),
  P(R('First of two'), mark('del')),
  P(R('Second of two.')),
].join('');
const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${xml}$1`));
fs.writeFileSync(new URL('./original.docx', import.meta.url), zipSync(files));
