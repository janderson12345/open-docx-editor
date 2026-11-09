// Builds test/fixtures/textbox.docx: a text box in line with the text, and one floating beside a paragraph (placed by
// offsets from the column and the paragraph, text wrapping around it), each with the VML copy older Word versions read.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { DocxSession } from '../../src/ooxml/docx.js';
import { blankDocxBytes } from '../../src/index.js';

const s = new DocxSession(blankDocxBytes());
const box = (w, h, text) => s.addTextBox(w, h).xml.replace(/<w:txbxContent><w:p\/><\/w:txbxContent>/g, `<w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent>`);
const inline = box(220, 60, 'An inline text box whose text wraps when it is resized.');
const floating = box(180, 90, 'A floating text box, placed by offsets.')
  .replace(/<wp:inline distT="0" distB="0" distL="0" distR="0">/, '<wp:anchor distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>3657600</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>95250</wp:posOffset></wp:positionV>')
  .replace(/(<wp:effectExtent [^>]*\/>)/, '$1<wp:wrapSquare wrapText="bothSides"/>')
  .replace('</wp:inline>', '</wp:anchor>')
  .replace(/style="width:/, 'style="position:absolute;margin-left:288pt;margin-top:7.5pt;width:');
const P = (...x) => `<w:p>${x.join('')}</w:p>`;
const R = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const lorem = 'The quick brown fox jumps over the lazy dog while seven stored procedures were rewritten to remove a predicate that forced a full scan of an eleven million row table. ';
const body = [P(R('Text boxes')), P(R('Inline: '), inline), P(floating, R(lorem.repeat(3))), P(R('End.'))].join('');
const files = unzipSync(blankDocxBytes());
const doc = strFromU8(files['word/document.xml']);
files['word/document.xml'] = strToU8(doc.replace(/<w:body>/, `<w:body>${body}`));
fs.writeFileSync(new URL('./textbox.docx', import.meta.url), zipSync(files));
