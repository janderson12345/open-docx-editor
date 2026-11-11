// Builds test/fixtures/altchunk.docx: imported content (w:altChunk), each a part Word converts to ordinary paragraphs
// when it opens the document, as mail-merge systems write rich-text fields: XHTML parts (relationship type aFChunk,
// .dat files of type application/xhtml+xml) in table cells, one blank (only white space, as a merge field with no
// value leaves it), one plain text, one rich; in the body a rich HTML part (headings, bold, italic, underline, a
// link, lists, a table, inline styles), an HTML part (.htm, text/html) and a plain-text part (.txt).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const P = (t, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${t ? `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>` : ''}</w:p>`;
const chunk = (id) => `<w:altChunk r:id="${id}"><w:altChunkPr><w:matchSrc w:val="false"/></w:altChunkPr></w:altChunk>`;
const cell = (w, inner) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
const border = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
const row = (label, id) => `<w:tr>${cell(2400, P(label, '<w:spacing w:after="0"/>'))}${cell(6960, chunk(id) + P('', '<w:spacing w:after="0"/>'))}</w:tr>`;
const table = `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>${border}<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="6960"/></w:tblGrid>`
  + row('Blank field', 'chunkBlank') + row('Plain field', 'chunkPlain') + row('Rich field', 'chunkRichCell') + '</w:tbl>';

const xml = [
  P('Imported content: before the table.'),
  table,
  P('A rich part in the body:'),
  chunk('chunkRich'),
  P('An HTML part:'),
  chunk('chunkHtm'),
  P('A plain-text part:'),
  chunk('chunkTxt'),
  P('The end.'),
].join('');

const blank = '<html>\n  <head/>\n  <body>\n    \n  </body>\n</html>';
const plain = '<html>\n  <head/>\n  <body>\n    Basis of Estimate: hours from the last three task orders.\n  </body>\n</html>';
const richCell = '<html><head/><body><p>First line of the field, <b>bold</b> and <i>italic</i>.</p><p>Second line.<br/>After a line break.</p></body></html>';
const rich = `<html xmlns="http://www.w3.org/1999/xhtml"><head><title>Rich</title></head><body>
<h1>Heading one</h1>
<p>A paragraph with <b>bold</b>, <strong>strong</strong>, <i>italic</i>, <em>em</em>, <u>underlined</u> and <s>struck</s> text, H<sub>2</sub>O and x<sup>2</sup>.</p>
<h2>Heading two</h2>
<p>A <a href="https://example.com/">link to example.com</a> in a sentence.</p>
<ul><li>First bullet</li><li>Second bullet<ul><li>Nested bullet</li></ul></li></ul>
<ol><li>First step</li><li>Second step</li></ol>
<h3>Heading three</h3>
<p><span style="color:#C00000">Red text</span>, <span style="font-size:16pt">16-point text</span>, <span style="font-family:Arial">Arial text</span> and <font color="#0070C0" face="Courier New">blue Courier</font>.</p>
<table border="1"><tr><td>Cell one</td><td>Cell two</td></tr><tr><td>Cell three</td><td>Cell four</td></tr></table>
<p style="text-align:center">Centred paragraph.</p>
</body></html>`;
const htm = '<!DOCTYPE html><html><head><meta charset="utf-8"><title>Htm</title></head><body><p>From an <b>.htm</b> part &amp; an entity: caf&eacute;.</p></body></html>';
const txt = 'Plain text part, first line.\r\nSecond line.';

const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${xml}$1`));
const parts = { chunkBlank: ['afchunk.dat', blank], chunkPlain: ['afchunk2.dat', plain], chunkRichCell: ['afchunk3.dat', richCell], chunkRich: ['afchunk4.dat', rich], chunkHtm: ['chunk5.htm', htm], chunkTxt: ['chunk6.txt', txt] };
const rels = 'word/_rels/document.xml.rels';
files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', Object.entries(parts).map(([id, [name]]) => `<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="/word/${name}" Id="${id}"/>`).join('') + '</Relationships>'));
for (const [name, body] of Object.values(parts)) files[`word/${name}`] = strToU8(body);
const ct = '[Content_Types].xml';
files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Default Extension="dat" ContentType="application/xhtml+xml"/><Default Extension="htm" ContentType="text/html"/><Default Extension="txt" ContentType="text/plain"/></Types>'));
fs.writeFileSync(new URL('./altchunk.docx', import.meta.url), zipSync(files));
