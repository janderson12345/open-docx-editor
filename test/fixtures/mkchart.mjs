// Builds test/fixtures/chart.docx: a column chart (two series, three categories) as Word writes one: the chart part
// with its cached values, and the workbook it's edited in embedded beside it (a sheet whose data is a table).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const cats = ['North', 'South', 'West'];
const series = [['2025', [4.2, 2.5, 3.1]], ['2026', [5.0, 3.4, 2.2]]];
const col = (i) => String.fromCharCode(66 + i);
const strCache = (xs) => `<c:strCache><c:ptCount val="${xs.length}"/>${xs.map((x, i) => `<c:pt idx="${i}"><c:v>${x}</c:v></c:pt>`).join('')}</c:strCache>`;
const numCache = (xs) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${xs.length}"/>${xs.map((x, i) => `<c:pt idx="${i}"><c:v>${x}</c:v></c:pt>`).join('')}</c:numCache>`;
const ser = ([name, vals], i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>Sheet1!$${col(i)}$1</c:f>${strCache([name])}</c:strRef></c:tx><c:invertIfNegative val="0"/>`
  + `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f>${strCache(cats)}</c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$${col(i)}$2:$${col(i)}$4</c:f>${numCache(vals)}</c:numRef></c:val></c:ser>`;
const chart = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:roundedCorners val="0"/><c:chart>`
  + `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
  + `<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series.map(ser).join('')}<c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="1001"/><c:axId val="1002"/></c:barChart>`
  + `<c:catAx><c:axId val="1001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="1002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`
  + `<c:valAx><c:axId val="1002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="1001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`
  + `</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`;

// The workbook: Sheet1 holding the data as Table1 (header row: a blank corner, then the series).
const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const cell = (ref, v) => (typeof v === 'number' ? `<c r="${ref}"><v>${v}</v></c>` : `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${v}</t></is></c>`);
const rows = [[' ', ...series.map((s) => s[0])], ...cats.map((c, k) => [c, ...series.map((s) => s[1][k])])];
const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${S}" xmlns:r="${R}"><dimension ref="A1:C4"/><sheetData>${rows.map((r, i) => `<row r="${i + 1}">${r.map((v, j) => cell(`${String.fromCharCode(65 + j)}${i + 1}`, v)).join('')}</row>`).join('')}</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>`;
const workbook = zipSync({
  '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>`),
  '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
  'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${S}" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
  'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
  'xl/worksheets/sheet1.xml': strToU8(sheet),
  'xl/worksheets/_rels/sheet1.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/table" Target="../tables/table1.xml"/></Relationships>`),
  'xl/tables/table1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="${S}" id="1" name="Table1" displayName="Table1" ref="A1:C4" totalsRowShown="0"><tableColumns count="3"><tableColumn id="1" name=" "/><tableColumn id="2" name="2025"/><tableColumn id="3" name="2026"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`),
});

const drawing = `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="5486400" cy="3200400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Chart 1"/><wp:cNvGraphicFramePr/>`
  + `<a:graphic xmlns:a="${A}"><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="rIdChart1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
const files = unzipSync(blankDocxBytes());
const doc = strFromU8(files['word/document.xml']);
files['word/document.xml'] = strToU8(doc.replace(/<w:body>/, `<w:body><w:p><w:r><w:t>A chart:</w:t></w:r></w:p><w:p>${drawing}</w:p><w:p><w:r><w:t>After the chart.</w:t></w:r></w:p>`));
files['word/charts/chart1.xml'] = strToU8(chart);
files['word/charts/_rels/chart1.xml.rels'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/></Relationships>`);
files['word/embeddings/Microsoft_Excel_Worksheet.xlsx'] = workbook;
const rels = strFromU8(files['word/_rels/document.xml.rels']);
files['word/_rels/document.xml.rels'] = strToU8(rels.replace('</Relationships>', `<Relationship Id="rIdChart1" Type="${R}/chart" Target="charts/chart1.xml"/></Relationships>`));
const ct = strFromU8(files['[Content_Types].xml']);
files['[Content_Types].xml'] = strToU8(ct.replace('</Types>', `${/Extension="xlsx"/.test(ct) ? '' : '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'}<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`));
fs.writeFileSync(new URL('./chart.docx', import.meta.url), zipSync(files));
