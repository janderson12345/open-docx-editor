// Builds test/fixtures/notes-long.docx: one paragraph with four footnotes of 44 items each, far more than a page holds,
// then a dozen short paragraphs. Word continues such notes at the foot of the following pages.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const item = (i) => `<w:p><w:r><w:t>Item ${i} of the long note list, with a few more words to fill the line.</w:t></w:r></w:p>`;
const note = (id) => `<w:footnote w:id="${id}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Note ${id}:</w:t></w:r></w:p>${Array.from({ length: 44 }, (_, i) => item(i + 1)).join('')}</w:footnote>`;
const seps = '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
const ref = (id) => `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
const para = (t) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;

const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
const body = `<w:p><w:r><w:t xml:space="preserve">One paragraph with four notes</w:t></w:r>${[1, 2, 3, 4].map(ref).join('')}</w:p>`
  + Array.from({ length: 12 }, (_, i) => para(`Body paragraph ${i + 1} after the notes.`)).join('');
files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
files['word/footnotes.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${W}>${seps}${[1, 2, 3, 4].map(note).join('')}</w:footnotes>`);
files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'));
const rels = 'word/_rels/document.xml.rels';
files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>'));
fs.writeFileSync(new URL('./notes-long.docx', import.meta.url), zipSync(files));
