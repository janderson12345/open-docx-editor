// Builds test/fixtures/vml-inline-boxes.docx: inline VML text boxes, and an inline VML group of two text boxes (a form's
// titled box: the title's box over the body's), as Word lays them out (its prints of the same documents, made here and
// printed by Word): each box at its place in the group, the group's line its height and the paragraph font's descent.
// With a folder (node test/fixtures/mkvmlgroup.mjs <dir>) it writes every variant tried in Word there instead.
import '../setup.js';
import fs from 'fs';
import path from 'path';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const out = process.argv[2] ? path.resolve(process.argv[2]) : null;
if (out) fs.mkdirSync(out, { recursive: true });
const V = 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"';

function write(name, body, font = 'Arial', size = 20) {
  const files = unzipSync(blankDocxBytes());
  const main = 'word/document.xml';
  let d = strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`);
  for (const ns of V.split(' ')) { const k = ns.split('=')[0]; if (!d.includes(k + '=')) d = d.replace('<w:document ', `<w:document ${ns} `); }
  files[main] = strToU8(d);
  let st = strFromU8(files['word/styles.xml']);
  st = st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, `$1<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:sz w:val="${size}"/></w:rPr>$2`);
  files['word/styles.xml'] = strToU8(st);
  fs.writeFileSync(out ? path.join(out, name) : new URL(`./${name}`, import.meta.url), zipSync(files));
}

const r = (t, rpr = '') => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const p = (t, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${t ? r(t) : ''}</w:p>`;
const txbx = (paras) => `<v:textbox inset="0,0,0,0"><w:txbxContent>${paras}</w:txbxContent></v:textbox>`;
const box = (id, l, t, w, h, paras, stroke = 'strokecolor="#231f20" strokeweight="1.44pt"') => `<v:shape id="${id}" type="#_x0000_t202" style="position:absolute;left:${l};top:${t};width:${w};height:${h}" filled="f" ${stroke}>${txbx(paras)}</v:shape>`;
const TYPE = '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
const title = p('Form Instruction', '<w:spacing w:before="40"/><w:ind w:left="2900" w:right="2900"/><w:jc w:val="center"/><w:rPr><w:b/><w:sz w:val="28"/></w:rPr>').replace('<w:r>', '<w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr>');
const body = ['The Placeholders Example passage in the sample wraps through to a further two sentences:', 'a. Sample quantities listed', 'b. Expanded sample marker', 'c. Another letter', '', 'Additional sample is applicable.']
  .map((t) => p(t, '<w:spacing w:before="40" w:line="186" w:lineRule="exact"/><w:ind w:left="42"/>')).join('');
// The titled box (sample 1113 page 5): 450pt wide here (541.45pt there), coordsize 10829 x 2929, the title's box 540 high on top.
const group = (h = 2929, stroke, w = 450) => `<w:pict>${TYPE}<v:group id="_x0000_s9001" style="width:${w}pt;height:${(h * w / 10829).toFixed(2)}pt;mso-position-horizontal-relative:char;mso-position-vertical-relative:line" coordsize="10829,${h}">${box('_x0000_s9003', 14, 554, 10801, h - 569, body, stroke)}${box('_x0000_s9002', 14, 14, 10801, 540, title, stroke)}<w10:anchorlock/></v:group></w:pict>`;
const single = `<w:pict>${TYPE}<v:shape id="_x0000_s9010" type="#_x0000_t202" style="width:153pt;height:99.85pt;mso-left-percent:-10001;mso-top-percent:-10001;mso-position-horizontal:absolute;mso-position-horizontal-relative:char;mso-position-vertical:absolute;mso-position-vertical-relative:line" filled="f" stroked="f">${txbx(p('EXAMPLES PLACEHOLDER', '<w:ind w:left="259"/>') + p('KIND: A - MAIN') + p('B - SPECIAL'))}<w10:anchorlock/></v:shape></w:pict>`;
const run = (pict) => `<w:r>${pict}</w:r>`;
const para = (inner, ppr = '<w:ind w:left="105"/>') => `<w:p><w:pPr>${ppr}</w:pPr><w:r/>${inner}</w:p>`;
const around = (x) => p('Before the group.') + x + p('After the group.') + p('Second line after.');
const variants = {
  'g01-group-arial10': [around(para(run(group())))],
  'g02-group-tnr12': [around(para(run(group()))), 'Times New Roman', 24],
  'g03-group-arial14-mark': [around(para(run(group()), '<w:ind w:left="105"/><w:rPr><w:sz w:val="28"/></w:rPr>'))],
  'g04-single-box': [around(para(run(single)))],
  'g05-box-tab-group': [around(para(run(single) + '<w:r><w:tab/></w:r>' + run(group(1997, undefined, 230)), '<w:tabs><w:tab w:val="left" w:pos="4500"/></w:tabs><w:ind w:left="300"/>'))],
  'g06-group-unstroked': [around(para(run(group(2929, 'stroked="f"'))))],
  'g07-group-in-cell': [p('Before the table.') + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="11000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="11000" w:type="dxa"/></w:tcPr>${para(run(group()), '<w:spacing w:before="59"/><w:jc w:val="center"/>')}</w:tc></w:tr></w:tbl>` + p('After the table.')],
  // A table whose third row's cells and the second's set nil lines between them, under the table's 1pt inside line, in
  // the body and in a text box (1113 page 60's codes table).
  'g11-nil-rows': [(() => {
    const B = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((k) => `<w:${k} w:val="single" w:sz="8" w:space="0" w:color="000000"/>`).join('');
    const c = (t, b) => `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/>${b ? `<w:tcBorders>${b}</w:tcBorders>` : ''}</w:tcPr><w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p></w:tc>`;
    const t = (k) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${B}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="227"/></w:trPr>${c(`${k} row one`)}${c('x')}</w:tr><w:tr><w:trPr><w:trHeight w:val="239"/></w:trPr>${c(`${k} row two`, '<w:bottom w:val="nil"/>')}${c('y', '<w:bottom w:val="nil"/>')}</w:tr><w:tr><w:trPr><w:trHeight w:val="400"/></w:trPr>${c(`${k} row three`, '<w:top w:val="nil"/>')}${c('z', '<w:top w:val="nil"/>')}</w:tr></w:tbl>`;
    const boxed = `<w:p><w:r><w:pict>${TYPE}<v:shape id="_x0000_s9020" type="#_x0000_t202" style="width:160pt;height:70pt;mso-position-horizontal-relative:char;mso-position-vertical-relative:line" filled="f" stroked="f"><v:textbox inset="0,0,0,0"><w:txbxContent>${t('Box')}<w:p/></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`;
    return p('Before') + t('Body') + p('Between') + boxed + p('After');
  })()],
  'g08-group-with-text': [around(para(r('Text ') + run(group())))],
  // (Its body at 256/240 lines, multiple, its lines wrapping: 1113 page 60's codes box.)
  'g09-group-multiple': [around(para(run(group().replace(/w:line="186" w:lineRule="exact"/g, 'w:line="256" w:lineRule="auto"').replace('a further two sentences:', 'a further two sentences, which wraps onto more lines than one in the box:'))))],
  'g10-box-multiple-bold': [around(para(run(single.replace(/<w:p>/g, '<w:p><w:pPr><w:spacing w:line="256" w:lineRule="auto"/><w:ind w:right="800"/><w:jc w:val="both"/></w:pPr>').replace('KIND: A - MAIN', 'A -  FIRST SAMPLE ITEM C -  LINE NUMBER SEQUENCE D - OTHER CASE'))))],
};
if (out) {
  for (const [n, [b, f, s]] of Object.entries(variants)) write(`${n}.docx`, b, f, s);
  console.log(Object.keys(variants).length, 'documents in', out);
} else {
  write('vml-inline-boxes.docx', variants['g01-group-arial10'][0] + variants['g05-box-tab-group'][0]);
}
