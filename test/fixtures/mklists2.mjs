// Builds test/fixtures/lists-more.docx, checked against Word's PDF of it:
// - list numbers formatted by their level (bold red, italic) and, with none, by the paragraph mark (bold);
// - numbering formats beyond decimal, letters and Roman: ordinal, cardinalText, ordinalText, enclosed numbers,
//   full-width digits, Arabic and Russian letters, ideographs; items 1-24 and 98-101 of each;
// - a paragraph with a bar border (w:pBdr/w:bar).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = (t, rpr = '') => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const P = (inner, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${inner}</w:p>`;
const item = (numId, t, markRpr = '') => P(R(t), `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>${markRpr ? `<w:rPr>${markRpr}</w:rPr>` : ''}`);

export const FORMATS = ['ordinal', 'cardinalText', 'ordinalText', 'decimalEnclosedCircle', 'decimalEnclosedParen', 'decimalEnclosedFullstop', 'decimalFullWidth', 'arabicAlpha', 'arabicAbjad', 'russianLower', 'russianUpper', 'ideographTraditional', 'ideographZodiac'];
const lvl = (fmt, rpr = '', text = '%1.') => `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="1080"/></w:pPr>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}</w:lvl>`;
const abstracts = [];
const nums = [];
let body = '';
// Label formatting: 1 bold red (the level's), 2 italic (the level's), 3 none but the paragraph mark bold.
abstracts.push(`<w:abstractNum w:abstractNumId="1">${lvl('decimal', '<w:b/><w:color w:val="C00000"/>')}</w:abstractNum>`, `<w:abstractNum w:abstractNumId="2">${lvl('upperLetter', '<w:i/>')}</w:abstractNum>`, `<w:abstractNum w:abstractNumId="3">${lvl('lowerRoman')}</w:abstractNum>`);
nums.push('<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>', '<w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>', '<w:num w:numId="3"><w:abstractNumId w:val="3"/></w:num>');
body += P(R('Label formatting'), '<w:pStyle w:val="Heading1"/>');
body += item(1, 'Bold red number (the level)') + item(1, 'Bold red number again');
body += item(2, 'Italic letter (the level)') + item(2, 'Italic letter again');
body += item(3, 'Bold numeral from the paragraph mark', '<w:b/>') + item(3, 'Plain numeral');
// The formats: one list each, 1-24, then the same list from 98.
FORMATS.forEach((fmt, k) => {
  const a = 10 + k, n1 = 100 + 2 * k, n2 = n1 + 1;
  abstracts.push(`<w:abstractNum w:abstractNumId="${a}">${lvl(fmt)}</w:abstractNum>`);
  nums.push(`<w:num w:numId="${n1}"><w:abstractNumId w:val="${a}"/></w:num>`, `<w:num w:numId="${n2}"><w:abstractNumId w:val="${a}"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="98"/></w:lvlOverride></w:num>`);
  body += P(R(`Format ${fmt}`), '<w:pStyle w:val="Heading1"/>');
  for (let i = 1; i <= 24; i++) body += item(n1, `${fmt} ${i}`);
  for (let i = 98; i <= 101; i++) body += item(n2, `${fmt} ${i}`);
});
body += P(R('Bar border'), '<w:pStyle w:val="Heading1"/>');
body += P(R('A paragraph with a bar border, which Word draws as a line beside it in the margin. It runs on to a second line so the line is two lines tall.'), '<w:pBdr><w:bar w:val="single" w:sz="18" w:space="4" w:color="FF00FF"/></w:pBdr><w:ind w:left="720"/>');

const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
files['word/numbering.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}>${abstracts.join('')}${nums.join('')}</w:numbering>`);
const ct = '[Content_Types].xml', rels = 'word/_rels/document.xml.rels';
if (!strFromU8(files[ct]).includes('/word/numbering.xml')) files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'));
if (!strFromU8(files[rels]).includes('numbering.xml')) files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'));
fs.writeFileSync(new URL('./lists-more.docx', import.meta.url), zipSync(files));
