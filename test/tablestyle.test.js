import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import { readStyles, tableCellCss } from '../src/ooxml/styles.js';

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
