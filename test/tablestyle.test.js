import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import { readStyles, tableCellCss, tableStylesCss } from '../src/ooxml/styles.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const cond = (type, fill) => `<w:tblStylePr w:type="${type}"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${fill}"/></w:tcPr></w:tblStylePr>`;
const STYLES = `<w:styles ${W}><w:style w:type="table" w:styleId="Corners"><w:name w:val="Corners"/>`
  + cond('firstRow', '111111') + cond('lastRow', '222222') + cond('firstCol', '333333') + cond('lastCol', '444444')
  + cond('nwCell', 'AA0000') + cond('neCell', '00AA00') + cond('swCell', '0000AA') + cond('seCell', 'AAAA00') + '</w:style></w:styles>';
const pkg = { partByType: () => 'styles.xml', xml: () => new DOMParser().parseFromString(STYLES, 'text/xml') };
const parseXml = (s) => new DOMParser().parseFromString(s, 'text/xml');

const css = (look) => {
  const cells = [];
  for (let row = 0; row < 3; row++) for (let col = 0; col < 3; col++) cells.push({ row, col, rowspan: 1, colspan: 1, tcPr: [] });
  const tblPr = `<w:tblPr><w:tblStyle w:val="Corners"/><w:tblLook ${look}/></w:tblPr>`;
  return tableCellCss(readStyles(pkg), tblPr, 3, 3, cells, parseXml).css.map((c) => /background-color:#?(\w+)/i.exec(c)?.[1]?.toUpperCase());
};

test('corner cells take the table style\'s corner formatting over their row and column', () => {
  const bg = css('w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="1" w:noHBand="1" w:noVBand="1"');
  assert.deepEqual([bg[0], bg[2], bg[6], bg[8]], ['AA0000', '00AA00', '0000AA', 'AAAA00']);
  assert.equal(bg[1], '111111'); // the first row between them
  assert.equal(bg[3], '333333'); // the first column between them
});

test('a corner applies only when both its row and its column are switched on', () => {
  const bg = css('w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"');
  assert.equal(bg[0], '111111');
  assert.equal(bg[2], '111111');
});

// The font size of a cell's text in the default paragraph style: the size the table style's rule sets on it, or null
// where it leaves the Normal style's (Word's prints of generated documents: defaults, Normal, table style; each mode).
const cellSize = (dd, normal, tableRPr, oldSizes = false, other = '', baseRPr = '') => {
  const sz = (v) => (v ? `<w:rPr><w:sz w:val="${v}"/></w:rPr>` : '');
  const xml = `<w:styles ${W}><w:docDefaults><w:rPrDefault>${sz(dd)}</w:rPrDefault></w:docDefaults>`
    + `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>${sz(normal)}</w:style>`
    + `<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body Text"/><w:basedOn w:val="Normal"/></w:style>`
    + `<w:style w:type="paragraph" w:styleId="Own"><w:name w:val="Own"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>`
    + `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>${baseRPr}</w:style>`
    + `<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/>${tableRPr ? `<w:rPr>${tableRPr}</w:rPr>` : ''}${other}</w:style></w:styles>`;
  const info = readStyles({ partByType: () => 'styles.xml', xml: () => new DOMParser().parseFromString(xml, 'text/xml') });
  const rules = tableStylesCss('.d', info, {}, { oldSizes }).split('\n');
  const size = (re) => { const r = rules.find((x) => re.test(x.split('{')[0]) && /font-size/.test(x)); return r ? Number(/font-size:([\d.]+)pt/.exec(r)[1]) : null; };
  return { cell: size(/data-tstyle="TableGrid"\] td > p:not\(\[data-style\]\)/), based: size(/td > p\[data-style="Body"\]/), own: size(/td > p\[data-style="Own"\]/) };
};

test('a size the Normal style sets is the cell\'s, over the table style\'s and the document defaults\'', () => {
  assert.equal(cellSize(22, 24, '').cell, null); // defaults 11pt, Normal 12pt: Normal's
  assert.equal(cellSize(22, 20, '<w:color w:val="000000"/>').cell, null);
  assert.equal(cellSize(22, 24, '<w:sz w:val="20"/>').cell, null); // nor the table style's 10pt
  assert.equal(cellSize(22, 22, '<w:sz w:val="20"/>').cell, null); // nor when Normal's is the defaults' size
  assert.equal(cellSize(22, 0, '<w:sz w:val="20"/>').cell, 10); // the table style's only where Normal sets none
  assert.equal(cellSize(0, 0, '<w:sz w:val="28"/>').cell, 14);
});

test('laid out as Word 2010 or older, a 12pt Normal gives way to a table style that formats its text', () => {
  assert.deepEqual(cellSize(22, 24, '<w:color w:val="000000"/>', true), { cell: 11, based: 11, own: null }); // the defaults' size
  assert.equal(cellSize(22, 24, '<w:lang w:val="en-GB"/>', true).cell, 11); // a language is enough
  assert.equal(cellSize(22, 24, '<w:color w:val="000000"/><w:sz w:val="28"/>', true).cell, 14); // the table style's own
  assert.equal(cellSize(22, 24, '', true).cell, null); // a table style of no text formatting: Normal's
  assert.equal(cellSize(22, 24, '', true, '<w:pPr><w:spacing w:after="0"/></w:pPr>').cell, 11); // a paragraph setting is enough
  assert.equal(cellSize(22, 24, '', true, '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>').cell, 11); // or a part's
  assert.equal(cellSize(22, 24, '', true, '<w:tcPr><w:shd w:val="clear" w:fill="FFFFFF"/></w:tcPr>').cell, null); // not a cell setting
  assert.equal(cellSize(22, 24, '', true, '<w:rPr/>').cell, null); // nor an empty one
  assert.equal(cellSize(22, 24, '', true, '', '<w:rPr><w:color w:val="000000"/></w:rPr>').cell, null); // nor the default table style's
  assert.equal(cellSize(22, 24, '<w:sz w:val="20"/>', true).cell, null); // 10pt is no size: Normal's 12pt stays
  assert.equal(cellSize(20, 24, '<w:color w:val="000000"/>', true).cell, null);
  assert.equal(cellSize(0, 24, '<w:color w:val="000000"/>', true).cell, null);
  assert.equal(cellSize(22, 20, '<w:color w:val="000000"/>', true).cell, null); // only a Normal of 12pt
  assert.equal(cellSize(22, 26, '<w:sz w:val="28"/>', true).cell, null);
  assert.equal(cellSize(22, 24, '<w:color w:val="000000"/>', false).cell, null); // Word 2013's layout, or the override set
});

test('fonts Word substitutes when missing, and faces named as fonts', async () => {
  const { fontStack, lineFactor, faceCss, fontCss } = await import('../src/ooxml/props.js');
  assert.match(fontStack('Palatino'), /^"Palatino", "Palatino Linotype"/);
  assert.match(fontStack('Helvetica Neue'), /^"Helvetica Neue", "Arial"/);
  assert.equal(lineFactor('Helvetica Neue'), lineFactor('Arial'));
  assert.equal(lineFactor('Futura Md BT'), 1.2261);
  assert.match(fontStack('Courier'), /^"Courier New"/);
  assert.equal(faceCss('Arial Bold'), ';font-weight:bold');
  assert.equal(faceCss('Times New Roman Bold Italic'), ';font-weight:bold;font-style:italic');
  assert.equal(faceCss('Arial'), '');
  assert.match(fontCss('Arial Bold'), /font-weight:bold$/);
});

// A cell's own bottom line against the top line of the row under it (Word's prints of generated documents).
const padTop = (c) => Number(/(?:^|;)padding-top:([\d.]+)px/.exec(c)?.[1]);
const grid = (tcPrOf, tblBorders = '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>') => {
  const cells = [];
  for (let row = 0; row < 3; row++) for (let col = 0; col < 3; col++) cells.push({ row, col, rowspan: 1, colspan: 1, tcPr: tcPrOf(row, col) });
  return tableCellCss(readStyles(pkg), `<w:tblPr><w:tblBorders>${tblBorders}</w:tblBorders></w:tblPr>`, 3, 3, cells, parseXml).css;
};
const bottom18 = [['tcBorders', '<w:tcBorders><w:bottom w:val="single" w:sz="18" w:space="0" w:color="000000"/></w:tcBorders>']];

test('a cell\'s own thick bottom line is the line over the row below it, wider than the table\'s inside line', () => {
  const c = grid((row) => (row === 0 ? bottom18 : []));
  assert.equal(padTop(c[3]), 3); // 2.25pt
  assert.equal(padTop(c[6]), 0.671875); // the next row: the table's 0.5pt line (0.667px, as the nearest 64th of a pixel)
});

test('one cell\'s thick bottom line lowers the whole row below it', () => {
  const c = grid((row, col) => (row === 0 && col === 1 ? bottom18 : []));
  assert.deepEqual([padTop(c[3]), padTop(c[4]), padTop(c[5])], [3, 3, 3]);
  assert.match(c[4], /box-shadow:inset 0px 3px/); // drawn under that cell only
  assert.doesNotMatch(c[3], /0px 3px/);
});

test('the table\'s own inside line gives way to a thinner top line a cell sets itself', () => {
  const thinTop = [['tcBorders', '<w:tcBorders><w:top w:val="single" w:sz="2" w:space="0" w:color="000000"/></w:tcBorders>']];
  const c = grid((row) => (row === 1 ? thinTop : []), '<w:insideH w:val="single" w:sz="12" w:space="0" w:color="000000"/>');
  assert.equal(padTop(c[3]), 0.328125);
});

test('a cell that draws its own top line keeps its place beside a cell under a thicker one', () => {
  const top = (sz) => [['tcBorders', `<w:tcBorders><w:top w:val="single" w:sz="${sz}" w:space="0" w:color="000000"/></w:tcBorders>`]];
  const c = grid((row, col) => (row === 1 ? (col === 0 ? top(2) : col === 1 ? top(12) : []) : []));
  assert.equal(padTop(c[3]), 0.328125); // its own 0.25pt line
  assert.equal(padTop(c[4]), 2);
  assert.equal(padTop(c[5]), 2); // the table's 0.5pt line: lowered to the row's widest
});

test('a wider line from the row above takes its room out of the top margin of a cell with a top line of its own', () => {
  const mar = '<w:tcMar><w:top w:w="80" w:type="dxa"/></w:tcMar>';
  const cell = (top, bottom) => [['tcBorders', `<w:tcBorders><w:top w:val="single" w:sz="${top}" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="${bottom}" w:space="0" w:color="000000"/></w:tcBorders>`], ['tcMar', mar]];
  const c = grid(() => cell(2, 8));
  assert.equal(padTop(c[3]), 5.671875); // 4pt margin and its own 0.25pt line, not the 1pt line's 6.667
  const none = grid((row) => (row === 0 ? cell(2, 8) : [['tcMar', mar]]), '');
  assert.equal(padTop(none[3]), 6.667); // no top line of its own: the line from above in full
});

test('a cell\'s own "no fill" clears the shade its band has from the table style', () => {
  const cells = [];
  for (let row = 0; row < 3; row++) for (let col = 0; col < 3; col++) cells.push({ row, col, rowspan: 1, colspan: 1, tcPr: row === 1 && col === 1 ? [['shd', '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>']] : [] });
  const css = tableCellCss(readStyles(pkg), '<w:tblPr><w:tblStyle w:val="Corners"/><w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>', 3, 3, cells, parseXml).css;
  assert.match(css[1], /background-color:#111111/i); // the first row keeps its shade
  assert.doesNotMatch(css[4], /background-color/); // the cell that says "no fill" has none
});

test('the top line a merged cell\'s lower part sets counts in that part\'s row', () => {
  const cells = [
    { row: 0, col: 0, rowspan: 2, colspan: 1, tcPr: [], vmTops: [8] }, { row: 0, col: 1, rowspan: 1, colspan: 1, tcPr: [] },
    { row: 1, col: 1, rowspan: 1, colspan: 1, tcPr: [['tcBorders', '<w:tcBorders><w:top w:val="nil"/></w:tcBorders>']] },
  ];
  const css = tableCellCss(readStyles(pkg), '<w:tblPr></w:tblPr>', 2, 2, cells, parseXml).css;
  assert.equal(padTop(css[2]), 1.328125); // a 1pt line's room, though the cell itself draws none
});

test('a line under a pixel wide is drawn at its own width on paper', () => {
  const c = grid(() => []);
  assert.match(c[3], /box-shadow:inset 0px calc\(1 \* max\(0\.667px, calc\(var\(--odx-hair, 1px\) \* 1\)\)\)/);
});
