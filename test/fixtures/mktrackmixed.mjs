// Builds test/fixtures/track-mixed-table.docx: Track Changes on, tables with vertically merged cells, and runs in another
// font than their paragraph's (Courier New and Cambria words in Times New Roman paragraphs), in the body and in the
// cells. Measuring such a run's line once put a probe inside its paragraph; the editor read that as an edit made in the
// page and drew the table again from it, and with tracking on an untouched document saved its tables twice over as
// tracked deletions and insertions (features.mjs checks it opens and saves unchanged).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const run = (t, font) => `<w:r><w:rPr>${font ? `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/>` : ''}</w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
const P = (...runs) => `<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr></w:pPr>${runs.join('')}</w:p>`;
const mixed = () => P(run('Plain text, then '), run('code words', 'Courier New'), run(' and '), run('Cambria words', 'Cambria'), run(' in one paragraph.'));
const tc = (body, merge) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${merge ? `<w:vMerge${merge === 'start' ? ' w:val="restart"' : ''}/>` : ''}</w:tcPr>${body}</w:tc>`;
const row = (a, b, c) => `<w:tr>${a}${b}${c}</w:tr>`;
// (Long enough to run over pages, merged cells crossing the page breaks: split across a page, the table's rows are laid
// out apart, and that is where the editor's own reading of it went wrong.)
const block = (n) => row(tc(mixed(), 'start'), tc(P(run(`Row ${n}`)), null), tc(mixed(), null))
  + row(tc(P(), 'cont'), tc(mixed(), null), tc(P(run(`Row ${n + 1}`)), 'start'))
  + row(tc(P(), 'cont'), tc(P(run(`Row ${n + 2}`)), null), tc(P(), 'cont'))
  + row(tc(mixed(), null), tc(mixed(), 'start'), tc(P(run(`Row ${n + 3}`)), null))
  + row(tc(P(run(`Row ${n + 4}`)), null), tc(P(), 'cont'), tc(mixed(), null));
const table = `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders><w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>`
  + Array.from({ length: 8 }, (_, i) => block(1 + 5 * i)).join('')
  + '</w:tbl>';
const body = mixed() + table + mixed() + table + mixed();
const files = unzipSync(blankDocxBytes());
files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
files['word/settings.xml'] = strToU8(strFromU8(files['word/settings.xml']).replace(/(<w:settings\b[^>]*>)/, '$1<w:trackRevisions/>'));
fs.writeFileSync(new URL('./track-mixed-table.docx', import.meta.url), zipSync(files));
