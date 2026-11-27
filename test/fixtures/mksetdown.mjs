// Fixtures: tables set down on their page (tblpPr vertAnchor="page") with text before them in the file.
// float-setdown-before.docx: a page of text reaching past the table's place, 8905tw down; Word's print puts the table on
// page 2 at that place, the paragraph after it starting page 2, and leaves page 1's text where it is.
// float-setdown-caption.docx: the table set down 1441tw from the top after a page and a half of text and its caption (kept
// with it); Word's print: page 2 opens with the table, the text before it and the caption under it.
// usage: node test/fixtures/mksetdown.mjs
import fs from 'node:fs';
import path from 'node:path';
import { zipSync, strToU8 } from 'fflate';

const out = path.dirname(new URL(import.meta.url).pathname.replace(/^\/(\w:)/, '$1'));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const TNR = '<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>';
const styles = (normal) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>${TNR}<w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>`
  + `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>${normal}</w:style>`
  + '<w:style w:type="paragraph" w:customStyle="1" w:styleId="paratext"><w:name w:val="paratext"/><w:pPr><w:spacing w:before="80" w:after="160" w:line="280" w:lineRule="exact"/><w:ind w:left="720"/></w:pPr><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:customStyle="1" w:styleId="tbltitle"><w:name w:val="tbl_title"/><w:pPr><w:keepNext/><w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="exact"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:sz w:val="22"/></w:rPr></w:style>'
  + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>';
const CALIBRI_NORMAL = '<w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>';
const theme = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements><a:clrScheme name="T"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="T"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="T"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="576" w:gutter="0"/><w:cols w:space="475"/><w:docGrid w:linePitch="360"/></w:sectPr>';
const bd = '<w:tcBorders><w:top w:val="single" w:sz="2" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="2" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="2" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="2" w:space="0" w:color="auto"/></w:tcBorders>';
const WING = '<w:rFonts w:ascii="Wingdings" w:eastAsia="Times New Roman" w:hAnsi="Wingdings" w:cs="Times New Roman"/>';
const cell = (text, w, { mar = true, font = TNR, sp = '<w:spacing w:before="40" w:after="40" w:line="240" w:lineRule="auto"/>', last = null } = {}, col = 0) => {
  const empty = last === 'empty' && col === 2, wing = last === 'wing' && col === 2;
  const run = empty ? '' : `<w:r><w:rPr>${wing ? WING : font}</w:rPr><w:t xml:space="preserve">${wing ? '' : text}</w:t></w:r>`;
  return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${bd}${mar ? '<w:tcMar><w:top w:w="29" w:type="dxa"/><w:bottom w:w="29" w:type="dxa"/></w:tcMar>' : ''}<w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr>${sp}<w:jc w:val="center"/><w:rPr>${font}</w:rPr></w:pPr>${run}</w:p></w:tc>`;
};
const table = (rows, { float = null, ...o } = {}) => `<w:tbl><w:tblPr>${float ? `<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="page" w:horzAnchor="margin" w:tblpY="${float}"/>` : ''}<w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="60" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000"/></w:tblPr><w:tblGrid><w:gridCol w:w="3960"/><w:gridCol w:w="2700"/><w:gridCol w:w="2700"/></w:tblGrid>`
  + Array.from({ length: rows }, (_, i) => `<w:tr>${cell(`Row ${i + 1} label`, 3960, o, 0)}${cell(`$${(i + 1) * 1111}`, 2700, o, 1)}${cell(`${i + 10}%`, 2700, o, 2)}</w:tr>`).join('') + '</w:tbl>';
const para = (text, style = 'paratext') => `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const filler = (n) => Array.from({ length: n }, (_, i) => para(`Filler paragraph ${i + 1} of plain text set at an exact fourteen point line, one line long.`)).join('');
const doc = (body) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}${sect}</w:body></w:document>`;
const save = (name, body, normal = CALIBRI_NORMAL) => {
  const files = {
    '[Content_Types].xml': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'),
    '_rels/.rels': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/_rels/document.xml.rels': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'),
    'word/document.xml': strToU8(doc(body)),
    'word/styles.xml': strToU8(styles(normal)),
    'word/theme/theme1.xml': strToU8(theme),
    'word/settings.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${W}"><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`),
  };
  fs.writeFileSync(path.join(out, `${name}.docx`), zipSync(files));
};

// Text before a table set down on its page, reaching into its place (8905tw from the page's top): a paragraph and the
// table's caption, kept with it, then more text after it.
const before = filler(22) + para('Text just before the set-down table, in the file before it.') + para('Table 1 The caption kept with the table', 'tbltitle');
save('float-setdown-before', before + table(6, { float: 8905 }) + para('Text after the table.') + filler(6));
// The same with the caption at the foot of a page, the table coming on the next page (set down 1441tw from its top).
save('float-setdown-caption', filler(36) + para('Table 2 A caption that keeps with its table', 'tbltitle') + table(6, { float: 1441 }) + para('Text after the table.') + filler(4));
console.log('made float-setdown-before.docx float-setdown-caption.docx');
