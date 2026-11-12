// Builds test/fixtures/hyphenation.docx: automatic hyphenation on (w:autoHyphenation), two narrow paragraphs of long
// words, the second with w:suppressAutoHyphens.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const TEXT = 'Internationalization and characterization of telecommunications infrastructure requires extraordinary organizational responsibilities.';
const P = (ppr) => `<w:p><w:pPr>${ppr}<w:ind w:right="4320"/></w:pPr><w:r><w:t xml:space="preserve">${TEXT}</w:t></w:r></w:p>`;
const files = unzipSync(blankDocxBytes());
files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${P('')}${P('<w:suppressAutoHyphens/>')}$1`));
// (The blank document has no settings part: one with the setting, its relationship and content type.)
files['word/settings.xml'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:autoHyphenation/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>');
files['word/_rels/document.xml.rels'] = strToU8(strFromU8(files['word/_rels/document.xml.rels']).replace('</Relationships>', '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'));
files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'));
fs.writeFileSync(new URL('./hyphenation.docx', import.meta.url), zipSync(files));
