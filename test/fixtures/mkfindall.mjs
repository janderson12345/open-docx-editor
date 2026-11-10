// Builds test/fixtures/find-all.docx: the word "apple" in each part Word's Find goes through: twice in the main text,
// once each in a text box, the header, the footer, a footnote and an endnote.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { DocxSession } from '../../src/ooxml/docx.js';
import { blankDocxBytes } from '../../src/index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const P = (...x) => `<w:p>${x.join('')}</w:p>`;
const s = new DocxSession(blankDocxBytes());
const box = s.addTextBox(220, 60).xml.replace(/<w:txbxContent><w:p\/><\/w:txbxContent>/g, `<w:txbxContent>${P(R('A text box apple.'))}</w:txbxContent>`);
const ref = (kind, id) => `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:${kind}Reference w:id="${id}"/></w:r>`;
const note = (kind, id, t) => `<w:${kind} w:id="${id}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:${kind}Ref/></w:r>${R(` ${t}`)}</w:p></w:${kind}>`;
const seps = (kind) => `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}><w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`;

const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
const body = P(R('First apple in the body.'), ref('footnote', 1)) + P(R('Inline: '), box) + P(R('Second apple in the body.'), ref('endnote', 1));
files[main] = strToU8(strFromU8(files[main])
  .replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)
  .replace(/<w:sectPr>/, '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>'));
files['word/header1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}>${P(R('Header apple'))}</w:hdr>`);
files['word/footer1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${W}>${P(R('Footer apple'))}</w:ftr>`);
files['word/footnotes.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${W}>${seps('footnote')}${note('footnote', 1, 'A footnote apple.')}</w:footnotes>`);
files['word/endnotes.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes ${W}>${seps('endnote')}${note('endnote', 1, 'An endnote apple.')}</w:endnotes>`);
const ct = (part, type) => `<Override PartName="/word/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`;
files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace('</Types>', ct('header1.xml', 'header') + ct('footer1.xml', 'footer') + ct('footnotes.xml', 'footnotes') + ct('endnotes.xml', 'endnotes') + '</Types>'));
const rel = (id, type, target) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
const rels = 'word/_rels/document.xml.rels';
files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', rel('rIdH', 'header', 'header1.xml') + rel('rIdF', 'footer', 'footer1.xml') + rel('rIdFn', 'footnotes', 'footnotes.xml') + rel('rIdEn', 'endnotes', 'endnotes.xml') + '</Relationships>'));
fs.writeFileSync(new URL('./find-all.docx', import.meta.url), zipSync(files));
