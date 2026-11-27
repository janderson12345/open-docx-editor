// Fixtures: floating tables placed up from the paragraph they're anchored to (tblpPr vertAnchor="text", tblpY -59), under
// 20 lines of text.
// float-up-moves.docx: 30 rows, past the page's foot, full width: Word's print puts it whole on page 2, its top 59tw above
// the page's top margin, and the text after it under it there.
// float-up-long-beside.docx: 70 rows (longer than a page), 7485tw wide and centred: Word's print starts it on page 2
// and breaks it at page 2's foot; the two empty paragraphs after it go beside its last rows on page 3, from that page's
// top, and the text after them on the third line there.
// usage: node test/fixtures/mkfloatup.mjs
import fs from 'node:fs';
import path from 'node:path';
import { zipSync, strToU8 } from 'fflate';

const out = path.dirname(new URL(import.meta.url).pathname.replace(/^\/(\w:)/, '$1'));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
  + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>';
const para = (text) => (text ? `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>` : '<w:p/>');
const table = (n, width) => {
  const b = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('');
  const c1 = Math.floor(width / 2);
  const rows = Array.from({ length: n }, (_, i) => `<w:tr><w:tc><w:tcPr><w:tcW w:w="${c1}" w:type="dxa"/></w:tcPr>${para(`Row ${i + 1}`)}</w:tc><w:tc><w:tcPr><w:tcW w:w="${width - c1}" w:type="dxa"/></w:tcPr>${para('cell')}</w:tc></w:tr>`).join('');
  return `<w:tbl><w:tblPr><w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="center" w:tblpY="-59"/><w:tblW w:w="${width}" w:type="dxa"/><w:tblBorders>${b}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="${c1}"/><w:gridCol w:w="${width - c1}"/></w:tblGrid>${rows}</w:tbl>`;
};
const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
const save = (name, n, width) => {
  const body = Array.from({ length: 20 }, (_, i) => para(`Filler line ${i + 1}`)).join('') + para('') + table(n, width) + para('') + para('') + para('After the table.');
  const files = {
    '[Content_Types].xml': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'),
    '_rels/.rels': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/_rels/document.xml.rels': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'),
    'word/document.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}${sect}</w:body></w:document>`),
    'word/styles.xml': strToU8(styles),
    'word/settings.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${W}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`),
  };
  fs.writeFileSync(path.join(out, `${name}.docx`), zipSync(files));
};
save('float-up-moves', 30, 9360);
save('float-up-long-beside', 70, 7485);
console.log('made float-up-moves.docx float-up-long-beside.docx');
