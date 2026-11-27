// Builds test/fixtures/table-style-size.docx, table-row-empty-part.docx, symbol-bullet-lines.docx and
// blank-tab-line.docx: row and line heights as Word lays them out (its prints of the same documents, made the same way).
// table-style-size: defaults Arial 11pt, Normal 12pt, a table style of 14pt; in its cells, paragraphs in Normal, in a
//   style of no base and no size ("Free"), one of 10pt of its own ("FreeSz") and one based on Normal ("OnNormal"):
//   Word draws them 12, 14, 10 and 12pt (Free takes the table style's size, the others as before).
// table-row-empty-part: a row of at least 485 twips whose long cell (51 lines) ends in an empty paragraph that only
//   just doesn't fit page 1: Word carries that paragraph to page 2 and gives the row's part there its whole minimum
//   height, the next row's text 47.4px under the top margin (15.4px lower than with no minimum height).
// symbol-bullet-lines: bullets typed by hand, a Symbol bullet, 7pt spaces, then Calibri 11pt text, in a Times New
//   Roman 12pt paragraph: Word's lines 17.92px apart (Calibri's), the first 18.72px (Symbol's ascent over Calibri's
//   descent).
// blank-tab-line: in a Times New Roman 12pt document, a cell paragraph of five 10pt Calibri spaces and a centre tab after
//   a line of text: Word gives it one line, the mark's, 18.4px (Times New Roman 12pt).
// With a folder (node test/fixtures/mkrowparts.mjs <dir>) it writes there instead.
import fs from 'fs';
import path from 'path';
import { zipSync, strToU8 } from 'fflate';

const out = process.argv[2] ? path.resolve(process.argv[2]) : null;
if (out) fs.mkdirSync(out, { recursive: true });
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
const TABLE_NORMAL = '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>';
const GRID = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';

/** The styles part: document defaults in `font` at `size` (half points), Normal's run properties, more styles. */
const stylesXml = (font, size, normalRpr, more = '') => `${HEAD}<w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}" w:eastAsia="${font}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>`
  + `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>${normalRpr ? `<w:rPr>${normalRpr}</w:rPr>` : ''}</w:style>${TABLE_NORMAL}${more}</w:styles>`;

function write(name, body, styles) {
  const files = {
    '[Content_Types].xml': `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`,
    '_rels/.rels': `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>`,
    'word/document.xml': `${HEAD}<w:document ${NS}><w:body>${body}${SECT}</w:body></w:document>`,
    'word/styles.xml': styles,
    'word/settings.xml': `${HEAD}<w:settings ${NS}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
  };
  const bytes = zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, strToU8(v)])));
  fs.writeFileSync(out ? path.join(out, name) : new URL(`./${name}`, import.meta.url), bytes);
}
const run = (t, rpr = '') => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const p = (t, ppr = '', rpr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${t ? run(t, rpr) : ''}</w:p>`;
const WORDS = 'alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november oscar papa quebec romeo sierra tango uniform victor whiskey xray yankee zulu'.split(' ');
function table(rows, widths, tblPr = '') {
  const grid = widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('');
  const trs = rows.map(([trPr, cells]) => `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells.map((c, i) => `<w:tc><w:tcPr><w:tcW w:w="${widths[i]}" w:type="dxa"/></w:tcPr>${c}</w:tc>`).join('')}</w:tr>`).join('');
  return `<w:tbl><w:tblPr>${tblPr}<w:tblW w:w="${widths.reduce((a, b) => a + b)}" w:type="dxa"/>${GRID}<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${trs}</w:tbl>`;
}
const pstyle = (id, based, rpr) => `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${id}"><w:name w:val="${id}"/>${based ? `<w:basedOn w:val="${based}"/>` : ''}${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}</w:style>`;

// table-style-size
{
  const ts = `<w:style w:type="table" w:styleId="TS"><w:name w:val="TS"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:rPr><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:tblPr>${GRID}</w:tblPr></w:style>`;
  const more = pstyle('Free') + pstyle('FreeSz', '', '<w:sz w:val="20"/>') + pstyle('OnNormal', 'Normal') + ts;
  const rows = ['Normal', 'Free', 'FreeSz', 'OnNormal'].flatMap((id) => [0, 1, 2].map((k) => ['', [p(`${id}${k} ${WORDS[k]} cell`, id === 'Normal' ? '' : `<w:pStyle w:val="${id}"/>`), p(`Side ${id}${k}`)]]));
  write('table-style-size.docx', p('Before the table') + table(rows, [4000, 3000], '<w:tblStyle w:val="TS"/>') + p('After the table'), stylesXml('Arial', 22, '<w:sz w:val="24"/>', more));
}

// table-row-empty-part
{
  const lines = Array.from({ length: 51 }, (_, i) => p(`Line${String(i).padStart(2, '0')} ${WORDS[i % 26]} text`)).join('');
  const rows = [['<w:trHeight w:val="485"/>', [p('Label') + p(''), lines + p('')]], ['', [p('NextRow marker'), p('NextRow right')]], ['', [p('ThirdRow marker'), p('ThirdRow right')]]];
  write('table-row-empty-part.docx', table(rows, [3000, 6000]) + p('After the table'), stylesXml('Arial', 22, ''));
}

// symbol-bullet-lines
{
  const cal = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/>';
  const text = Array.from({ length: 70 }, (_, i) => WORDS[i % 26]).join(' ');
  const para = `<w:p><w:pPr><w:ind w:left="720" w:hanging="360"/><w:rPr>${cal}</w:rPr></w:pPr>${run('', '<w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/><w:sz w:val="22"/>')}${run('       ', '<w:sz w:val="14"/>')}${run(text, cal)}</w:p>`;
  write('symbol-bullet-lines.docx', p('Start marker') + para.repeat(3) + p('End marker'), stylesXml('Calibri', 22, '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/>'));
}

// blank-tab-line
{
  const cal = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="20"/>';
  const blank = `<w:p><w:pPr><w:tabs><w:tab w:val="center" w:pos="5477"/></w:tabs></w:pPr>${run('     ', cal)}<w:r><w:rPr>${cal}</w:rPr><w:tab/></w:r></w:p>`;
  const first = `<w:p>${run('Firstline text in the cell', cal)}</w:p>`;
  const rows = [['<w:trHeight w:val="420"/>', [first + blank]], ['', [p('NextRow marker')]]];
  write('blank-tab-line.docx', p('Start marker') + table(rows, [9000]) + p('End marker'), stylesXml('Times New Roman', 24, ''));
}
