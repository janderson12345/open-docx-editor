// Builds two fixtures of vertical spacing between plain paragraphs (features.mjs), each checked against Word:
// - spaces-lines.docx: 12pt Times New Roman lines at 1.08 (w:line="259"), between "X" lines, holding nothing but spaces:
//   an 18pt and a 12pt space, an 18pt space, two, an 18pt lowered space, all over a 12pt mark; and a 12pt space over an
//   18pt mark. Word makes each the mark's line (15pt apart, as the X lines; 21.75pt for the 18pt mark), not its
//   spaces': runs of only spaces don't size a line. (An 18pt space had made its line 22.35pt.)
// - columns-empty-section.docx: two lines, then a continuous section in two columns holding only its empty
//   section-break paragraph, then two lines in one column again. Word gives that empty paragraph its line, as in one
//   column (the text after it 29.25pt under the line before, two lines; it had come up one, onto the empty paragraph's).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const F = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>';
const run = (sz, t, extra = '') => `<w:r><w:rPr>${F}<w:sz w:val="${sz}"/>${extra}</w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
const para = (mark, runs) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="259" w:lineRule="auto"/><w:rPr>${F}<w:sz w:val="${mark}"/></w:rPr></w:pPr>${runs}</w:p>`;
const sub = '<w:vertAlign w:val="subscript"/>';
const page = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
const sect = (type, cols) => `<w:sectPr>${type ? '<w:type w:val="continuous"/>' : ''}${page}<w:cols w:space="720"${cols ? ` w:num="${cols}"` : ''}/></w:sectPr>`;

function make(name, body, last) {
  const files = unzipSync(blankDocxBytes());
  let doc = strFromU8(files['word/document.xml']);
  doc = doc.replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${body}${last}</w:body>`);
  files['word/document.xml'] = strToU8(doc);
  fs.writeFileSync(new URL(`./${name}`, import.meta.url), zipSync(files));
}

const X = para(24, run(24, 'X'));
const blanks = [
  para(24, run(36, ' ') + run(24, ' ')),
  para(24, run(36, ' ')),
  para(24, run(36, '  ')),
  para(24, run(36, ' ', sub)),
  para(36, run(24, ' ')),
];
make('spaces-lines.docx', X + blanks.map((b) => b + X).join(''), sect(false, 0));

const line = (t) => para(24, run(24, t));
const body = line('First line.')
  + `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="259" w:lineRule="auto"/>${sect(false, 0)}</w:pPr>${run(24, 'Last line of the first section.')}</w:p>`
  + `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="259" w:lineRule="auto"/>${sect(true, 2)}</w:pPr></w:p>`
  + line('After the empty columns section.') + line('One more line.');
make('columns-empty-section.docx', body, sect(true, 0));
