// Builds test/fixtures/lists-spacing.docx, checked against Word's PDF of it (Times New Roman 12pt body, single
// spacing, no space after):
// - HTML-style auto spacing (w:beforeAutospacing / w:afterAutospacing) between items of one list: Word drops each side
//   on its own, an item's auto space after under the next item even when that one has no auto space of its own, and
//   the next's auto space before under an item with a plain 6pt after (which stays);
// - two numIds over one abstract list count as one list for it;
// - a web style's auto spacing turned off outright (w:afterAutospacing="0", w:after="150"): the items keep 7.5pt;
//   a plain w:after over that style, not turning it off, is still the auto 14pt;
// - an empty paragraph whose mark is in the theme's font (w:asciiTheme, Calibri) under the Times style: a Calibri line.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const P = (t, ppr = '', mark = '') => `<w:p><w:pPr>${ppr}${mark ? `<w:rPr>${mark}</w:rPr>` : ''}</w:pPr>${t ? R(t) : ''}</w:p>`;
const num = (n) => `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${n}"/></w:numPr>`;
const BA = '<w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"/>';
const B = '<w:spacing w:before="100" w:beforeAutospacing="1" w:after="0"/>';
const WEB = '<w:pStyle w:val="NormalWeb"/>';
const OFF = '<w:spacing w:before="0" w:beforeAutospacing="0" w:after="150" w:afterAutospacing="0"/>';
const THEME = '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/>';

const body = [
  P('Auto spacing on one side'),
  P('auto-a1', num(1) + BA), P('auto-a2', num(1)), P('auto-a3', num(1) + '<w:spacing w:after="120"/>'), P('auto-a4', num(1) + B), P('auto-a5'),
  P('Two nums, one abstract'),
  P('auto-b1', num(1) + BA), P('auto-b2', num(2) + BA), P('auto-b3', num(2) + BA), P('auto-b4'),
  P('Turned off'),
  P('off-1', WEB + num(1) + OFF), P('off-2', WEB + num(1) + OFF), P('off-3', WEB + num(1) + OFF), P('off-4', WEB + '<w:spacing w:after="150"/>'), P('off-5'),
  P('Empty theme mark'),
  P('mark-1'), P('', '', THEME + '<w:sz w:val="24"/>'), P('mark-2'), P('', ''), P('mark-3'),
].join('');

const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
const stylesXml = strFromU8(files['word/styles.xml']);
const TNR = '<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/>';
const normal = `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr>${TNR}</w:rPr></w:style>`;
const web = '<w:style w:type="paragraph" w:styleId="NormalWeb"><w:name w:val="Normal (Web)"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"/></w:pPr></w:style>';
files['word/styles.xml'] = strToU8(stylesXml.replace(/<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?<\/w:style>/, normal).replace('</w:styles>', `${web}</w:styles>`));
const lvl = '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>';
files['word/numbering.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}><w:abstractNum w:abstractNumId="0">${lvl}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num></w:numbering>`);
const ct = '[Content_Types].xml', rels = 'word/_rels/document.xml.rels';
if (!strFromU8(files[ct]).includes('/word/numbering.xml')) files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'));
if (!strFromU8(files[rels]).includes('numbering.xml')) files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'));
// (A theme part, for the mark's theme font: Calibri, as in Word's default.)
files['word/theme/theme1.xml'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements><a:clrScheme name="C"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="F"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>');
files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>'));
files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>'));
fs.writeFileSync(new URL('./lists-spacing.docx', import.meta.url), zipSync(files));
