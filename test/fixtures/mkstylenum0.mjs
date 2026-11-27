// Builds test/fixtures/style-numid0.docx: paragraph styles that switch a numbered base style's list off (w:numId 0 in
// the style itself, as a "Table TEXT" style based on a bullet style does), as Word lays them out (its print of the
// same document, made here).
// Each paragraph starts with its variant's name and wraps to a second line, so its first and other lines show where
// Word puts them:
//   base styles: L (bulleted, list linked to the style, own left 360), U (the same, the list not linked), N (linked, no
//   indent of its own), P (no list, left 360);
//   derived with numId 0: dL, dU, dN, dP; dL540 (own left 540), dLr (own right 364 only), dLf (own first line 200);
//   ddL (based on dL, nothing of its own); and dL with a direct left 200, a direct right 364;
//   Q (no list, left 360, right 1440, hanging 200), dQ, dQh (own hanging 300); R (left 360, first line 400), dR.
// With a folder (node test/fixtures/mkstylenum0.mjs <dir>) it writes there instead.
import '../setup.js';
import fs from 'fs';
import path from 'path';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const out = process.argv[2] ? path.resolve(process.argv[2]) : null;
if (out) fs.mkdirSync(out, { recursive: true });
const SINGLE = '<w:spacing w:before="0" w:after="120" w:line="240" w:lineRule="auto"/>';

// numId n -> a bullet level at left 720, hanging 360, linked to style link (or none).
const lists = { 1: 'L', 2: '', 3: 'N' };
const numbering = Object.entries(lists).map(([id, link]) => `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>${link ? `<w:pStyle w:val="${link}"/>` : ''}<w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>`).join('')
  + Object.keys(lists).map((id) => `<w:num w:numId="${id}"><w:abstractNumId w:val="${id}"/></w:num>`).join('');

const num = (id) => `<w:numPr><w:numId w:val="${id}"/></w:numPr>`;
const style = (id, based, ppr) => `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${id}"><w:name w:val="${id}"/><w:basedOn w:val="${based}"/><w:qFormat/><w:pPr>${ppr}</w:pPr></w:style>`;
const styles = [
  style('L', 'Normal', `${num(1)}<w:ind w:left="360"/>`),
  style('U', 'Normal', `${num(2)}<w:ind w:left="360"/>`),
  style('N', 'Normal', num(3)),
  style('P', 'Normal', '<w:ind w:left="360"/>'),
  style('dL', 'L', num(0)),
  style('dU', 'U', num(0)),
  style('dN', 'N', num(0)),
  style('dP', 'P', num(0)),
  style('dL540', 'L', `${num(0)}<w:ind w:left="540"/>`),
  style('dLr', 'L', `${num(0)}<w:ind w:right="364"/>`),
  style('dLf', 'L', `${num(0)}<w:ind w:firstLine="200"/>`),
  style('ddL', 'dL', ''),
  style('Q', 'Normal', '<w:ind w:left="360" w:right="1440" w:hanging="200"/>'),
  style('dQ', 'Q', num(0)),
  style('dQh', 'Q', `${num(0)}<w:ind w:hanging="300"/>`),
  style('R', 'Normal', '<w:ind w:left="360" w:firstLine="400"/>'),
  style('dR', 'R', num(0)),
].join('');

const TEXT = 'lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua ut enim ad minim veniam quis';
const p = (name, styleId, ppr = '') => `<w:p><w:pPr><w:pStyle w:val="${styleId}"/>${ppr}</w:pPr><w:r><w:t xml:space="preserve">${name} ${TEXT}</w:t></w:r></w:p>`;
const body = [
  p('Lx', 'L'), p('Ux', 'U'), p('Nx', 'N'), p('Px', 'P'),
  p('dLx', 'dL'), p('dUx', 'dU'), p('dNx', 'dN'), p('dPx', 'dP'),
  p('dL540x', 'dL540'), p('dLrx', 'dLr'), p('dLfx', 'dLf'), p('ddLx', 'ddL'),
  p('dLind200x', 'dL', '<w:ind w:left="200"/>'), p('dLright364x', 'dL', '<w:ind w:right="364"/>'),
  p('Qx', 'Q'), p('dQx', 'dQ'), p('dQhx', 'dQh'), p('Rx', 'R'), p('dRx', 'dR'),
].join('');

const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
let st = strFromU8(files['word/styles.xml']);
st = st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, `$1<w:pPr>${SINGLE}</w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr>$2`);
files['word/styles.xml'] = strToU8(st.replace('</w:styles>', `${styles}</w:styles>`));
const ct = '[Content_Types].xml', rels = 'word/_rels/document.xml.rels';
files['word/numbering.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}>${numbering}</w:numbering>`);
if (!strFromU8(files[ct]).includes('/word/numbering.xml')) files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'));
if (!strFromU8(files[rels]).includes('numbering.xml')) files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'));
fs.writeFileSync(out ? path.join(out, 'style-numid0.docx') : new URL('./style-numid0.docx', import.meta.url), zipSync(files));
