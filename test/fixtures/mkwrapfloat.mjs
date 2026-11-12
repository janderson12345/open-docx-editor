// Builds test/fixtures/wrap-float.docx: floating shapes wrapped square as Word wraps them (test/wrapfloat.mjs): one in
// the middle of the column with text on both sides (wrapText bothSides), one with text on its left only (left), one
// centred in the column (align center, both sides), and one near the page's foot reaching past its edge (which Word
// keeps on the page, drawn up to end at the edge).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { DocxSession } from '../../src/ooxml/docx.js';
import { blankDocxBytes } from '../../src/index.js';

const s = new DocxSession(blankDocxBytes());
const float = (xml, h, y, side) => xml
  .replace(/<wp:inline distT="0" distB="0" distL="0" distR="0">/, `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>${h}<wp:positionV relativeFrom="paragraph"><wp:posOffset>${Math.round(y * 9525)}</wp:posOffset></wp:positionV>`)
  .replace(/(<wp:effectExtent [^>]*\/>)/, `$1<wp:wrapSquare wrapText="${side}"/>`)
  .replace('</wp:inline>', '</wp:anchor>');
const at = (x) => `<wp:positionH relativeFrom="column"><wp:posOffset>${Math.round(x * 9525)}</wp:posOffset></wp:positionH>`;
const P = (...x) => `<w:p>${x.join('')}</w:p>`;
const R = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const lorem = 'The quick brown fox jumps over the lazy dog while seven stored procedures were rewritten to remove a predicate that forced a full scan. ';
const body = [
  P(R('Both sides:')),
  P(float(s.addShape('rect', 100, 60).xml, at(150), 10, 'bothSides'), R(lorem.repeat(2))),
  P(R('Left only:')),
  P(float(s.addShape('ellipse', 100, 60).xml, at(200), 10, 'left'), R(lorem.repeat(2))),
  P(R('Centred:')),
  P(float(s.addShape('roundRect', 100, 60).xml, '<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>', 10, 'bothSides'), R(lorem.repeat(2))),
];
for (let i = 1; i <= 14; i++) body.push(P(R(`Line ${i}.`)));
body.push(P(float(s.addShape('rect', 120, 240).xml, at(0), 10, 'right'), R('The foot: a float reaching past the page\'s edge.')));
body.push(P(R('After.')));
const files = unzipSync(blankDocxBytes());
files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/<w:body>/, `<w:body>${body.join('')}`));
fs.writeFileSync(new URL('./wrap-float.docx', import.meta.url), zipSync(files));
