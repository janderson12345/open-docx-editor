// Builds test/fixtures/autofit-indents.docx and cell-auto-space.docx: table cells as Word lays them (its prints of the
// same tables, made here and printed by Word).
// autofit-indents: three autofit tables, a first column of 600 twips (40px) holding (1) empty paragraphs bulleted at a
//   720 indent, (2) "x" after a 720 first-line indent, (3) "x" bulleted at 720: Word keeps (1) and (2) at 40px (an
//   empty paragraph needs nothing, a first-line indent isn't counted) and widens (3) to its indent, word and margins.
// cell-auto-space: a Normal (Web) paragraph (14pt auto spacing) alone in a cell, then first of two in another table,
//   then two list items of it ending a cell:
//   Word drops the auto space after the last paragraph of a cell (the first row as high as the second), not the first's,
//   nor a list's last item's.
// With a folder (node test/fixtures/mkautofit.mjs <dir>) it writes every variant tried in Word there instead.
import '../setup.js';
import fs from 'fs';
import path from 'path';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const out = process.argv[2] ? path.resolve(process.argv[2]) : null;
if (out) fs.mkdirSync(out, { recursive: true });
const SINGLE = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const ALL = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';

function write(name, { body, numbering, styles = '' }) {
  const files = unzipSync(blankDocxBytes());
  const main = 'word/document.xml';
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
  let st = strFromU8(files['word/styles.xml']);
  st = st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, `$1<w:pPr>${SINGLE}</w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>$2`);
  files['word/styles.xml'] = strToU8(st.replace('</w:styles>', `${styles}</w:styles>`));
  const ct = '[Content_Types].xml', rels = 'word/_rels/document.xml.rels';
  files['word/numbering.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}>${numbering}</w:numbering>`);
  if (!strFromU8(files[ct]).includes('/word/numbering.xml')) files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'));
  if (!strFromU8(files[rels]).includes('numbering.xml')) files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'));
  fs.writeFileSync(out ? path.join(out, name) : new URL(`./${name}`, import.meta.url), zipSync(files));
}

// numId n -> one bullet level of its own indent.
const lvls = { 6: [720, 360, '', 'Symbol'], 7: [720, 360, '•', ''], 1: [720, 720, '□', 'Times New Roman'], 2: [360, 360, '□', 'Times New Roman'], 3: [720, 360, '•', 'Symbol'], 4: [1440, 720, '□', 'Times New Roman'], 5: [720, 720, '%1.', ''] };
const numbering = Object.entries(lvls).map(([id, [left, hang, text, font]]) => `<w:abstractNum w:abstractNumId="${id}"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${text.includes('%') ? 'decimal' : 'bullet'}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="${hang}"/></w:pPr>${font ? `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/></w:rPr>` : ''}</w:lvl></w:abstractNum>`).join('')
  + Object.keys(lvls).map((id) => `<w:num w:numId="${id}"><w:abstractNumId w:val="${id}"/></w:num>`).join('');

const num = (id) => `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${id}"/></w:numPr>`;
// (In the schema's order: numbering before spacing, indents after it.)
const p = (text, ppr = '') => `<w:p><w:pPr>${ppr.startsWith('<w:ind') ? SINGLE + ppr : ppr + SINGLE}</w:pPr>${text ? run(text) : ''}</w:p>`;
const tbl = (first, w1 = 600) => {
  const ws = [w1, 2400, 9360 - 2400 - w1];
  const c = (w, x) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${x}</w:tc>`;
  const rows = [0, 1, 2].map((i) => `<w:tr>${c(ws[0], first)}${c(ws[1], p(`Second ${i}`))}${c(ws[2], p(`Third column text ${i}`))}</w:tr>`).join('');
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${ALL}</w:tblPr><w:tblGrid>${ws.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>${p('After')}`;
};
const variants = {
  'v01-bullet720-empty': tbl(p('', num(1))),
  'v02-bullet720-x': tbl(p('x', num(1))),
  'v03-ind720-x': tbl(p('x', '<w:ind w:left="720"/>')),
  'v04-ind720-empty': tbl(p('', '<w:ind w:left="720"/>')),
  'v05-bullet360-empty': tbl(p('', num(2))),
  'v06-longword': tbl(p('Abcdefghijkl')),
  'v07-bullet720-longword': tbl(p('Abcdefghijkl', num(1))),
  'v08-bullet720-empty-w1200': tbl(p('', num(1)), 1200),
  'v09-symbol720h360-empty': tbl(p('', num(3))),
  'v10-bullet1440-empty': tbl(p('', num(4))),
  'v11-decimal720-empty': tbl(p('', num(5))),
  'v12-ind1440-x': tbl(p('x', '<w:ind w:left="1440"/>')),
  'v13-first720-x': tbl(p('x', '<w:ind w:firstLine="720"/>')),
  'v14-right720-x': tbl(p('x', '<w:ind w:right="720"/>')),
  'v15-bullet720-ab': tbl(p('ab', num(1))),
};
const WEB = '<w:style w:type="paragraph" w:styleId="NormalWeb"><w:name w:val="Normal (Web)"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"/></w:pPr></w:style>';
const web = (t, id) => `<w:p><w:pPr><w:pStyle w:val="NormalWeb"/>${id ? num(id) : ''}</w:pPr>${run(t)}</w:p>`;
const row3 = (a, b, c) => `<w:tr>${[a, b, c].map((x) => `<w:tc><w:tcPr><w:tcW w:w="3120" w:type="dxa"/></w:tcPr>${x}</w:tc>`).join('')}</w:tr>`;
const tbl3 = (rows) => `${p('Before')}<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${ALL}</w:tblPr><w:tblGrid><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/></w:tblGrid>${rows}</w:tbl>${p('After')}`;
const spacing = {
  's01-web-alone': tbl3(row3(p('A1'), web('B1'), p('C1')) + row3(p('A2'), p('B2'), p('C2'))),
  's02-web-last-of-two': tbl3(row3(p('A1'), p('B1') + web('B1b'), p('C1')) + row3(p('A2'), p('B2'), p('C2'))),
  's03-web-first-of-two': tbl3(row3(p('A1'), web('B1') + p('B1b'), p('C1')) + row3(p('A2'), p('B2'), p('C2'))),
  's04-web-all-cells': tbl3(row3(web('A1'), web('B1'), web('C1')) + row3(p('A2'), p('B2'), p('C2'))),
  's05-web-two-in-cell': tbl3(row3(p('A1'), web('B1') + web('B1b'), p('C1')) + row3(p('A2'), p('B2'), p('C2'))),
  's06-web-last-row': tbl3(row3(p('A1'), p('B1'), p('C1')) + row3(p('A2'), web('B2'), p('C2'))),
  's07-web-body': web('Web one') + web('Web two') + p('Plain'),
  's09-web-list-two': tbl3(row3(p('A1'), web('B1', 7) + web('B1b', 7), p('C1')) + row3(p('A2'), p('B2'), p('C2'))),
  's10-web-list-one': tbl3(row3(p('A1'), web('B1', 7), p('C1')) + row3(p('A2'), p('B2'), p('C2'))),
  's11-plain-then-web-item': tbl3(row3(p('A1'), p('B1') + web('B1b', 7), p('C1')) + row3(p('A2'), p('B2'), p('C2'))),
  's12-web-then-web-item': tbl3(row3(p('A1'), web('B1') + web('B1b', 7), p('C1')) + row3(p('A2'), p('B2'), p('C2'))),
  's13-web-list-two-lastrow': tbl3(row3(p('A1'), p('B1'), p('C1')) + row3(p('A2'), web('B2', 7) + web('B2b', 7), p('C2'))),
  's08-web-tall-neighbour': tbl3(row3(p('A1') + p('A1b') + p('A1c'), web('B1'), p('C1')) + row3(p('A2'), p('B2'), p('C2'))),
};
const LP = '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>';
const NORM = (rule) => `<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:pPr><w:spacing w:after="260" w:line="260" w:lineRule="${rule}"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style>`;
const body = (t) => `<w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr>${run(t)}</w:p>`;
const item = (t, id, line = 259) => `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>${id ? num(id) : ''}<w:spacing w:after="160" w:line="${line}" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const items = (id, line) => body('First body paragraph.') + body('Second body paragraph:') + [1, 2, 3, 4].map((k) => item(`Item ${k} text`, id, line)).join('') + body('After the list.');
const lists = {
  'l01-atleast-symbol': [items(6), 'atLeast'],
  'l02-atleast-nobullet': [items(0), 'atLeast'],
  'l03-atleast-calibri-bullet': [items(7), 'atLeast'],
  'l04-auto-symbol': [items(6), 'auto'],
  'l05-atleast-symbol-single': [items(6, 240), 'atLeast'],
  'l06-exact-symbol': [items(6), 'exact'],
};
if (out) {
  for (const [n, body] of Object.entries(variants)) write(`${n}.docx`, { body, numbering });
  for (const [n, body] of Object.entries(spacing)) write(`${n}.docx`, { body, numbering, styles: WEB });
  for (const [n, [body, rule]] of Object.entries(lists)) write(`${n}.docx`, { body, numbering, styles: LP + NORM(rule) });
  console.log(Object.keys(variants).length + Object.keys(spacing).length + Object.keys(lists).length, 'documents in', out);
} else {
  write('autofit-indents.docx', { body: tbl(p('', num(1))) + tbl(p('x', '<w:ind w:firstLine="720"/>')) + tbl(p('x', num(1))), numbering });
  write('cell-auto-space.docx', { body: spacing['s01-web-alone'] + spacing['s03-web-first-of-two'] + spacing['s13-web-list-two-lastrow'], numbering, styles: WEB });
}
