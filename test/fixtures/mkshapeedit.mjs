// Builds test/fixtures/shape-edit.docx: shapes as Insert > Shapes makes them, for moving and resizing with the mouse
// (test/shapeedit.mjs): in line with the text an arrow, a callout (its pointer below its box) and a line with an
// arrowhead; floating beside a paragraph (placed by offsets from the column and the paragraph) a line and a rounded
// rectangle.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { DocxSession } from '../../src/ooxml/docx.js';
import { blankDocxBytes } from '../../src/index.js';

const s = new DocxSession(blankDocxBytes());
const float = (xml, x, y) => xml
  .replace(/<wp:inline distT="0" distB="0" distL="0" distR="0">/, `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>${x * 9525}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${y * 9525}</wp:posOffset></wp:positionV>`)
  .replace(/(<wp:effectExtent [^>]*\/>)/, '$1<wp:wrapNone/>')
  .replace('</wp:inline>', '</wp:anchor>');
const P = (...x) => `<w:p>${x.join('')}</w:p>`;
const R = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const lorem = 'The quick brown fox jumps over the lazy dog while seven stored procedures were rewritten to remove a predicate. ';
const body = [
  P(R('Shapes in line: '), s.addShape('rightArrow', 96, 72).xml, R(' '), s.addShape('wedgeRoundRectCallout', 96, 72).xml, R(' '), s.addShape('line', 144, 1, { arrow: true }).xml),
  P(R('Line after the shapes.')),
  P(float(s.addShape('line', 144, 48).xml, 360, 10), float(s.addShape('roundRect', 120, 80).xml, 400, 90), R(lorem.repeat(4))),
  P(R('End.')),
].join('');
const files = unzipSync(blankDocxBytes());
files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/<w:body>/, `<w:body>${body}`));
fs.writeFileSync(new URL('./shape-edit.docx', import.meta.url), zipSync(files));
