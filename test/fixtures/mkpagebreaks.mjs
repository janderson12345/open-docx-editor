// Builds test/fixtures/pagebreak-*.docx: where Word breaks pages (test/features.mjs), as its PDFs of these showed.
//   pagebreak-float-table.docx  two floating tables (vertAnchor text) 365tw narrower than the column, an empty paragraph
//     between them: nothing wraps beside them, and the second breaks at the page's foot row by row (Word: rows 1-23 on
//     page 1, 24-40 and the text after them on page 2, under the table).
//   pagebreak-tiny-footer.docx  a two-line footer reaching above the bottom margin, closed by an empty paragraph at 0.06
//     lines (w:line 14): that paragraph takes about 1px, so page 1 holds body paragraphs 1-54.
//   pagebreak-keep-long.docx  a heading kept with a table of 70 rows each keeping with the next, near page 1's foot: the
//     heading starts page 2 with the table's first 53 rows under it, the rest on page 3.
import fs from 'fs';
import { zipSync, strToU8 } from 'fflate';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const CT = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>';
const RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
const DRELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>';
const STY = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`;

const p = (text, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${text ? `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>` : ''}</w:p>`;
const paras = (n, pre = 'Body') => Array.from({ length: n }, (_, i) => p(`${pre} paragraph ${i + 1}`)).join('');
const BORDERS = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
function table(rows, { floating = false, width = 9000, bottomFromText = 0, ppr = '' } = {}) {
  const tp = floating ? `<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:bottomFromText="${bottomFromText}" w:vertAnchor="text" w:tblpY="1"/><w:tblOverlap w:val="never"/>` : '';
  let x = `<w:tbl><w:tblPr>${tp}<w:tblW w:w="${width}" w:type="dxa"/>${BORDERS}<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="${width - 3000}"/></w:tblGrid>`;
  for (let i = 1; i <= rows; i++) x += `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(`Row ${i}`, ppr)}</w:tc><w:tc><w:tcPr><w:tcW w:w="${width - 3000}" w:type="dxa"/></w:tcPr>${p(`Row ${i} second cell`, ppr)}</w:tc></w:tr>`;
  return x + '</w:tbl>';
}
function docx(name, body, { footer = p('Footer line'), bottom = 1440, footerDist = 720 } = {}) {
  const sect = `<w:sectPr><w:headerReference w:type="default" r:id="rId3"/><w:footerReference w:type="default" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="${bottom}" w:left="1440" w:header="720" w:footer="${footerDist}" w:gutter="0"/></w:sectPr>`;
  const files = {
    '[Content_Types].xml': CT, '_rels/.rels': RELS, 'word/_rels/document.xml.rels': DRELS, 'word/styles.xml': STY,
    'word/document.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${sect}</w:body></w:document>`,
    'word/footer1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${W}>${footer}</w:ftr>`,
    'word/header1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}>${p('Header line')}</w:hdr>`,
  };
  fs.writeFileSync(new URL(name, import.meta.url), zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, strToU8(v)]))));
}

docx('pagebreak-float-table.docx', paras(10) + table(20, { floating: true, width: 8995 }) + p('') + table(40, { floating: true, width: 8995, bottomFromText: 160 }) + paras(5, 'After'));
docx('pagebreak-tiny-footer.docx', paras(60), { footer: p('Footer line') + p('Footer two') + p('', '<w:spacing w:line="14" w:lineRule="auto"/>'), bottom: 1680, footerDist: 1495 });
docx('pagebreak-keep-long.docx', paras(50) + p('Heading text', '<w:keepNext/>') + table(70, { ppr: '<w:keepNext/><w:keepLines/>' }) + paras(3, 'After'));
