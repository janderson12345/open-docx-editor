// Builds test documents for Word's compatibility rules (w:compat in settings.xml), each in Word 2003's compatibility
// mode (11), where they apply, twice: compat-<rule>-on.docx with the rule set, compat-<rule>-off.docx without it. Word's
// own PDFs of each pair show what the rule changes. (Written to the folder given, not test/fixtures.)
// usage: node test/fixtures/mkcompat.mjs <outDir>
import '../setup.js';
import fs from 'fs';
import path from 'path';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const out = process.argv[2];
fs.mkdirSync(out, { recursive: true });
const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;');
const run = (t, rpr = '') => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
const para = (runs, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${runs}</w:p>`;
const LONG = 'The quick brown fox jumps over the lazy dog, and then it runs on to the end of the line so that the paragraph wraps onto a second line. ';
const SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>';
const SETTINGS_REL = '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>';
const SETTINGS_TYPE = '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>';

function build(name, body, rule) {
  for (const on of [true, false]) {
    const files = unzipSync(blankDocxBytes());
    files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${body}${SECT}</w:body>`));
    files['word/settings.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:defaultTabStop w:val="720"/><w:compat>${on ? `<w:${rule}/>` : ''}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="11"/></w:compat></w:settings>`);
    files['word/_rels/document.xml.rels'] = strToU8(strFromU8(files['word/_rels/document.xml.rels']).replace('</Relationships>', SETTINGS_REL + '</Relationships>'));
    files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace('</Types>', SETTINGS_TYPE + '</Types>'));
    fs.writeFileSync(path.join(out, `compat-${name}-${on ? 'on' : 'off'}.docx`), zipSync(files));
  }
}

// A page break at the end of a paragraph: its paragraph mark on the page before or after.
build('splitpg', [para(run('Before the break.')), para(run('Text, then a page break at the end of the paragraph:') + '<w:r><w:br w:type="page"/></w:r>'), para(run('The next paragraph.')), ...Array.from({ length: 3 }, () => para(run(LONG)))].join(''), 'splitPgBreakAndParaMark');

// A numbered list whose hanging indent is narrower than the label: the indent as the tab stop after the number, or not.
const numPr = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>';
build('indenttab', [para(run('A numbered list with a narrow hanging indent:')), ...Array.from({ length: 3 }, (_, i) => para(run(`Item ${i + 1} of the list, long enough to wrap onto a second line so the indent shows. ${LONG}`), `${numPr}<w:ind w:left="360" w:hanging="360"/>`)),
  para(run('With a hanging indent wider than the label:')), ...Array.from({ length: 2 }, (_, i) => para(run(`Item ${i + 4} of the list. ${LONG}`), `${numPr}<w:ind w:left="1080" w:hanging="1080"/>`))].join(''), 'doNotUseIndentAsNumberingTabStop');

// HTML-style automatic spacing (w:beforeAutospacing / w:afterAutospacing), as pasted from web pages.
build('htmlspacing', [para(run('First, with automatic spacing before and after.'), '<w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"/>'),
  para(run('Second, the same.'), '<w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"/>'),
  para(run('A list item with it:'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"/>'),
  para(run('Another list item with it.'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"/>'),
  para(run('A plain paragraph after them.')), para(run(LONG))].join(''), 'doNotUseHTMLParagraphAutoSpacing');

// A floating table near the foot of a page: broken across the page, or moved whole.
const cell = (t) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${para(run(t))}</w:tc>`;
const ftable = `<w:tbl><w:tblPr><w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:tblpY="1"/><w:tblW w:w="8000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${Array.from({ length: 12 }, (_, i) => `<w:tr>${cell(`Row ${i + 1}, first cell`)}${cell(`Row ${i + 1}, second cell`)}</w:tr>`).join('')}</w:tbl>`;
build('wrappedtables', [...Array.from({ length: 30 }, (_, i) => para(run(`Filler line ${i + 1}.`))), ftable, para(run('After the table. ' + LONG)), para(run(LONG))].join(''), 'doNotBreakWrappedTables');

// East Asian layout rules on a Latin document (Word sets this in many documents): lines and spacing.
build('felayout', [para(run('A document with East Asian layout rules turned on. ' + LONG + LONG)), para(run('Tabs:\tone\ttwo\tthree')), para(run(LONG), '<w:jc w:val="both"/>'), para(run('Text with ') + run('superscript', '<w:vertAlign w:val="superscript"/>') + run(' and ') + run('subscript', '<w:vertAlign w:val="subscript"/>') + run(' in it. ' + LONG))].join(''), 'useFELayout');

// The last tab stop on a line, right-aligned, with text past it.
build('lasttab', [para(run('Name') + '<w:r><w:tab/></w:r>' + run('A right-aligned tab with more text than fits before the margin, so it runs past the stop'), '<w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs>'), para(run('Short') + '<w:r><w:tab/></w:r>' + run('fits'), '<w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs>')].join(''), 'forgetLastTabAlignment');

// A numbered list whose number is underlined: the tab after it underlined too, or not.
build('ultab', [para(run('Underlined numbers:')), ...Array.from({ length: 2 }, (_, i) => para(run(`Item ${i + 1}.`, '<w:u w:val="single"/>'), `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:rPr><w:u w:val="single"/></w:rPr>`))].join(''), 'underlineTabInNumList');
