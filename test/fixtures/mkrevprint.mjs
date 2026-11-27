// Builds test/fixtures/revprint.docx: tracked changes and hidden paragraph marks as Word prints them by default (its
// PDF export with the default item, made from the same documents built in Word):
// - a paragraph with an inserted word and a tracked bold: printed without a change bar or a tint;
// - a paragraph indented 2in whose mark is hidden (w:vanish), then one at the margin: one paragraph, at the margin
//   (compatibility mode 15, set here; with no mode set Word keeps the pair at 2in);
// - a paragraph whose mark is hidden and w:specVanish (a style separator), indented 2in, then one at the margin: one
//   paragraph, at 2in;
// - a four-item list, the second item's mark deleted: 1. (one), 2. (two and three run on), 3. (four);
// - a three-item list, the second item's text all deleted: 1., 2. on an empty line, 3.;
// - a paragraph whose mark was inserted and then deleted, its text too: gone.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
let id = 1;
const who = () => `w:id="${id++}" w:author="Reviewer" w:date="2026-11-01T10:00:00Z"`;
const R = (t, rpr = '') => `<w:r>${rpr}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const INS = (t) => `<w:ins ${who()}>${R(t)}</w:ins>`;
const DEL = (t) => `<w:del ${who()}><w:r><w:delText xml:space="preserve">${t}</w:delText></w:r></w:del>`;
const P = (body, ppr = '', mark = '') => `<w:p><w:pPr>${ppr}${mark ? `<w:rPr>${mark}</w:rPr>` : ''}</w:pPr>${body}</w:p>`;
const NUM = (n) => `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${n}"/></w:numPr>`;
const IND = (left) => `<w:ind w:left="${left}"/>`;

const body = [
  P(R('Changed ') + INS('inserted ') + R('and bold.', `<w:rPr><w:b/><w:rPrChange ${who()}><w:rPr/></w:rPrChange></w:rPr>`)),
  P(R('Hidden mark A. '), IND(2880), '<w:vanish/>'),
  P(R('Runs on at the margin.'), IND(0)),
  P(R('Separator B. '), IND(2880), '<w:vanish/><w:specVanish/>'),
  P(R('Runs on at 2in.'), IND(0)),
  P(R('Item one.'), NUM(1)),
  P(R('Item two.'), NUM(1), `<w:del ${who()}/>`),
  P(R('Item three.'), NUM(1)),
  P(R('Item four.'), NUM(1)),
  P(R('Second one.'), NUM(2)),
  P(DEL('Second two.'), NUM(2)),
  P(R('Second three.'), NUM(2)),
  P(`<w:ins ${who()}>${DEL('Inserted and deleted.')}</w:ins>`, '', `<w:ins ${who()}/><w:del ${who()}/>`),
  P(R('The end.')),
].join('');
const lvl = '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>';
const numbering = `<w:abstractNum w:abstractNumId="0">${lvl}</w:abstractNum><w:abstractNum w:abstractNumId="1">${lvl}</w:abstractNum>`
  + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>';

const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
const ct = '[Content_Types].xml', rels = 'word/_rels/document.xml.rels';
files['word/numbering.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}>${numbering}</w:numbering>`);
if (!strFromU8(files[ct]).includes('/word/numbering.xml')) files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'));
if (!strFromU8(files[rels]).includes('numbering.xml')) files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'));
// Compatibility mode 15 (Word 2013 and later), where a hidden mark's paragraph runs on in the next one's formatting.
files['word/settings.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`);
if (!strFromU8(files[ct]).includes('/word/settings.xml')) files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'));
if (!strFromU8(files[rels]).includes('settings.xml')) files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'));
fs.writeFileSync(new URL('./revprint.docx', import.meta.url), zipSync(files));
