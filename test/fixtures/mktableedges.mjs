// Builds test/fixtures/table-edges.docx, hide-mark.docx, float-no-anchor.docx and float-text-noroom.docx.
// table-edges: contextual spacing across a table's edges. A paragraph with "don't add space between paragraphs of the
// same style" above a one-cell table whose paragraph has the same style and the same setting, and one under it: Word
// drops the space after the first, the cell paragraph's spaces before and after, and the space before the last (a row
// ends with a mark of the default paragraph style, which the cell's paragraph and the one under the table meet). The
// same without the setting keeps all four (10pt each).
// float-no-anchor: a floating table with no w:vertAnchor, 2in down: Word places it against the margin (as the text after
// five paragraphs, it would be 2in under them), and the paragraphs after it fill the page above it.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const CTX = '<w:contextualSpacing/>';
const sp = (before, after) => `<w:spacing w:before="${before}" w:after="${after}" w:line="240" w:lineRule="auto"/>`;
const para = (text, ppr = '') => `<w:p><w:pPr>${ppr}</w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const borders = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
const table = (cells, tblpPr = '') => `<w:tbl><w:tblPr>${tblpPr}<w:tblW w:w="9000" w:type="dxa"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>`
  + cells.map((c) => `<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${c}</w:tc></w:tr>`).join('') + '</w:tbl>';

function write(name, body) {
  const files = unzipSync(blankDocxBytes());
  let doc = strFromU8(files['word/document.xml']);
  doc = doc.replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`);
  files['word/document.xml'] = strToU8(doc);
  fs.writeFileSync(new URL(`./${name}`, import.meta.url), zipSync(files));
}

write('table-edges.docx', [
  para('Above with', sp(0, 200) + CTX), table([para('Cell with', sp(200, 200) + CTX)]), para('Below with', sp(200, 0) + CTX),
  para('Gap', sp(0, 0)),
  para('Above without', sp(0, 200)), table([para('Cell without', sp(200, 200))]), para('Below without', sp(200, 0)),
].join(''));

// hide-mark: a three-row table whose middle row's cells are empty and leave their end marks out of the row's height
// (w:hideMark), and the same row in a second table without it. Word: the first table's middle row takes no room (its
// third row's line 14.25pt under the first's), the second's a line and its space after (36.75pt).
const hm = (tcpr, body) => `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/>${tcpr}</w:tcPr>${body}</w:tc>`;
const rows3 = (tcpr, tag) => `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>`
  + `<w:tr>${hm('', para(`First ${tag}`, sp(0, 0)))}${hm('', para('', sp(0, 0)))}</w:tr>`
  + `<w:tr>${hm(tcpr, '<w:p/>')}${hm(tcpr, '<w:p/>')}</w:tr>`
  + `<w:tr>${hm('', para(`Third ${tag}`, sp(0, 0)))}${hm('', para('', sp(0, 0)))}</w:tr></w:tbl>`;
write('hide-mark.docx', rows3('<w:hideMark/>', 'hidden') + para('Between', sp(0, 0)) + rows3('', 'shown') + para('End', sp(0, 0)));

write('float-no-anchor.docx', [
  ...Array.from({ length: 5 }, (_, i) => para(`Before ${i + 1}`, sp(0, 0))),
  table([para('Floating row 1', sp(0, 0)), para('Floating row 2', sp(0, 0))], '<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:horzAnchor="margin" w:tblpY="2880"/>'),
  ...Array.from({ length: 12 }, (_, i) => para(`After ${i + 1}`, sp(0, 0))),
].join(''));

// float-text-noroom: a table as wide as the text, placed 451tw down from the paragraph after it (w:vertAnchor="text"),
// then a paragraph, an empty one and two more. Word: the first above the table, the empty one under it (its line meets
// the table's top), then the others.
write('float-text-noroom.docx', [
  table([para('Placed row 1', sp(0, 0)), para('Placed row 2', sp(0, 0))], '<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="451"/>'),
  para('Text 1', sp(0, 0)), '<w:p><w:pPr>' + sp(0, 0) + '</w:pPr></w:p>', para('Text 2', sp(0, 0)), para('Text 3', sp(0, 0)),
].join(''));
