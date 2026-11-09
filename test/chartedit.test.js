import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { DocxSession } from '../src/ooxml/docx.js';
import { readChart, writeChart, workbookRel, updateWorkbook } from '../src/ooxml/chartedit.js';

const bytes = fs.readFileSync(new URL('./fixtures/chart.docx', import.meta.url));
const part = () => strFromU8(unzipSync(bytes)['word/charts/chart1.xml']);

test('a chart\'s data is read from its cached values', () => {
  const d = readChart(part());
  assert.equal(d.type, 'column');
  assert.equal(d.title, 'Sales');
  assert.deepEqual(d.cats, ['North', 'South', 'West']);
  assert.deepEqual(d.series, [{ name: '2025', vals: [4.2, 2.5, 3.1] }, { name: '2026', vals: [5, 3.4, 2.2] }]);
  assert.deepEqual(d.types, ['column', 'bar', 'line', 'area']);
});

test('new data, title and type are written as cached values and references to the sheet', () => {
  const data = { type: 'line', title: 'Revenue', hasTitle: true, cats: ['N', 'S', 'W', 'E'], series: [{ name: 'A', vals: [1, 2, null, 4] }, { name: 'B', vals: [5, 6, 7, 8] }, { name: 'C', vals: [9, 9, 9, 9] }] };
  const xml = writeChart(part(), data);
  const back = readChart(xml);
  assert.equal(back.type, 'line');
  assert.equal(back.title, 'Revenue');
  assert.deepEqual(back.cats, data.cats);
  assert.deepEqual(back.series, data.series);
  assert.match(xml, /<c:f>Sheet1!\$D\$2:\$D\$5<\/c:f>/); // the third series' values, column D
  assert.doesNotMatch(xml, /invertIfNegative/); // (a bar series' element, not a line's)
  assert.equal((/<c:lineChart>[\s\S]*<\/c:lineChart>/.exec(xml)[0].match(/<c:axId /g) || []).length, 2, 'the axes go with it');
  // Fewer series again, and no title.
  const less = readChart(writeChart(xml, { ...data, title: '', hasTitle: false, series: data.series.slice(0, 1) }));
  assert.equal(less.series.length, 1);
  assert.equal(less.hasTitle, false);
});

test('the embedded workbook gets the data, its table resized to it', () => {
  const files = unzipSync(bytes);
  const rel = workbookRel(part(), [{ id: 'rId1', target: 'word/embeddings/Microsoft_Excel_Worksheet.xlsx', external: false }]);
  assert.equal(rel.target, 'word/embeddings/Microsoft_Excel_Worksheet.xlsx');
  const data = { cats: ['N', 'S', 'W', 'E'], series: [{ name: 'A', vals: [1, 2, 3, 4] }, { name: 'B', vals: [5, 6, 7, 8] }, { name: 'C', vals: [9, 9, 9, 9] }] };
  const wb = unzipSync(updateWorkbook(files[rel.target], data));
  const sheet = strFromU8(wb['xl/worksheets/sheet1.xml']);
  assert.match(sheet, /<dimension ref="A1:D5"\/>/);
  assert.match(sheet, /<c r="D1" t="inlineStr"><is><t xml:space="preserve">C<\/t><\/is><\/c>/);
  assert.match(sheet, /<c r="B5"><v>4<\/v><\/c>/);
  const table = strFromU8(wb['xl/tables/table1.xml']);
  assert.match(table, /ref="A1:D5"/);
  assert.match(table, /<tableColumns count="4">.*name="C"/);
});

test('an edited chart is saved into its part and workbook; the drawing that shows it is unchanged', () => {
  const s = new DocxSession(bytes);
  let node = null;
  s.doc.descendants((n) => { if (!node && n.type.name === 'image') node = n; });
  assert.ok(node && /<c:chart\b/.test(node.attrs.xml), 'the chart is a picture node');
  const edited = writeChart(s.chartPart(node.attrs.xml).xml, { type: 'bar', title: 'Edited', hasTitle: true, cats: ['X', 'Y'], series: [{ name: 'Only', vals: [3, 4] }] });
  // (The editor sets the attribute with setNodeMarkup; here the node is rebuilt with it.)
  const out = s.doc.type.create(s.doc.attrs, mapImages(s.doc, (n) => n.type.create({ ...n.attrs, chart: edited }, n.content, n.marks)));
  const files = unzipSync(s.save(out));
  const chart = strFromU8(files['word/charts/chart1.xml']);
  assert.match(chart, /barDir val="bar"/);
  assert.match(chart, /<a:t>Edited<\/a:t>/);
  const wb = unzipSync(files['word/embeddings/Microsoft_Excel_Worksheet.xlsx']);
  assert.match(strFromU8(wb['xl/worksheets/sheet1.xml']), /<dimension ref="A1:B3"\/>/);
  assert.match(strFromU8(files['word/document.xml']), /<c:chart [^>]*r:id="rIdChart1"/);
});

/** A copy of `node` with its image nodes replaced by `fn(image)`. */
function mapImages(node, fn) {
  const kids = [];
  node.forEach((c) => {
    if (c.type.name === 'image') kids.push(fn(c));
    else if (c.isLeaf) kids.push(c);
    else kids.push(c.type.create(c.attrs, mapImages(c, fn), c.marks));
  });
  return kids;
}

test('a title the dialog left alone stays as it was (Word\'s automatic title has no text to read)', () => {
  const auto = part().replace(/<c:title>[\s\S]*?<\/c:title>/, '<c:title><c:overlay val="0"/></c:title>');
  const d = readChart(auto);
  assert.equal(d.title, '');
  const xml = writeChart(auto, { ...d, title: null, cats: d.cats, series: d.series });
  assert.match(xml, /<c:title>/);
  assert.doesNotMatch(xml, /autoTitleDeleted val="1"/);
});

test('a combination chart (bars and a line) is shown but not edited; a literal series name is read', () => {
  const combo = part().replace('</c:barChart>', '</c:barChart><c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="2"/><c:order val="2"/><c:tx><c:v>Target</c:v></c:tx><c:val><c:numRef><c:f>Sheet1!$D$2:$D$4</c:f><c:numCache><c:ptCount val="3"/></c:numCache></c:numRef></c:val></c:ser><c:axId val="1001"/><c:axId val="1002"/></c:lineChart>');
  assert.equal(readChart(combo).editable, false);
  const literal = part().replace(/<c:tx><c:strRef><c:f>Sheet1!\$B\$1<\/c:f>[\s\S]*?<\/c:strRef><\/c:tx>/, '<c:tx><c:v>Last year</c:v></c:tx>');
  assert.equal(readChart(literal).series[0].name, 'Last year');
});

test('a quoted sheet name with a quote in it is read whole', () => {
  const quoted = part().replace(/Sheet1!/g, "'Bob''s data'!");
  assert.equal(readChart(quoted).sheet, "Bob's data");
});

test('saving again after the edit is undone puts the chart and its workbook back as they were', () => {
  const s = new DocxSession(bytes);
  const edited = writeChart(part(), { type: 'bar', title: 'Edited', hasTitle: true, cats: ['X'], series: [{ name: 'Only', vals: [3] }] });
  const withEdit = s.doc.type.create(s.doc.attrs, mapImages(s.doc, (n) => n.type.create({ ...n.attrs, chart: edited }, n.content, n.marks)));
  const first = unzipSync(s.save(withEdit));
  assert.match(strFromU8(first['word/charts/chart1.xml']), /Edited/);
  const again = unzipSync(s.save(s.doc)); // the edit undone: the node has no chart of its own
  assert.equal(strFromU8(again['word/charts/chart1.xml']), part());
  assert.deepEqual(again['word/embeddings/Microsoft_Excel_Worksheet.xlsx'], unzipSync(bytes)['word/embeddings/Microsoft_Excel_Worksheet.xlsx']);
  let node = null;
  s.doc.descendants((n) => { if (!node && n.type.name === 'image') node = n; });
  assert.equal(s.chartPart(node.attrs.xml).xml, part(), 'and the dialog reads the chart as opened');
});
