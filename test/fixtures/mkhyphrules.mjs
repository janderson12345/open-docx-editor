// Builds test/fixtures/hyphenation-{free,zone,limit}.docx: automatic hyphenation on, four narrow paragraphs of long
// words; free has no zone or limit (Word 2013's layout), zone a 1in hyphenation zone in Word 2010's layout
// (compatibility mode 14, where Word follows it), limit at most one line in a row ending in a hyphen.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const WORDS = 'the international organization of telecommunications infrastructure and its characteristically extraordinary responsibilities for administration require considerable understanding of environmental sustainability'.split(' ');
const text = (k) => Array.from({ length: 60 }, (_, j) => WORDS[(k * 5 + j * 3) % WORDS.length]).join(' ') + '.';
const P = (k) => `<w:p><w:pPr><w:ind w:right="3600"/></w:pPr><w:r><w:t xml:space="preserve">${text(k)}</w:t></w:r></w:p>`;
for (const [name, mode, extra] of [['free', 15, ''], ['zone', 14, '<w:hyphenationZone w:val="1440"/>'], ['limit', 15, '<w:consecutiveHyphenLimit w:val="1"/>']]) {
  const files = unzipSync(blankDocxBytes());
  files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${[0, 1, 2, 3].map(P).join('')}$1`));
  // (The blank document has no settings part: one with the settings, its relationship and content type.)
  files['word/settings.xml'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + `<w:autoHyphenation/>${extra}<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/></w:compat></w:settings>`);
  files['word/_rels/document.xml.rels'] = strToU8(strFromU8(files['word/_rels/document.xml.rels']).replace('</Relationships>', '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'));
  files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'));
  fs.writeFileSync(new URL(`./hyphenation-${name}.docx`, import.meta.url), zipSync(files));
}
