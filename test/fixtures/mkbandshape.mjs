// Builds test/fixtures/band-shapes.docx: floating VML shapes with no text, wrapped top and bottom (a grey band, a ruled
// line), placed down from their paragraph, as Word lays out the text round them (its prints of the same documents, made
// here and printed by Word): the paragraph's own line above a band that starts below it, the next paragraph under it.
// With a folder (node test/fixtures/mkbandshape.mjs <dir>) it writes every variant tried in Word there instead.
import '../setup.js';
import fs from 'fs';
import path from 'path';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const out = process.argv[2] ? path.resolve(process.argv[2]) : null;
if (out) fs.mkdirSync(out, { recursive: true });
const NS = 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"';

function write(name, body) {
  const files = unzipSync(blankDocxBytes());
  const main = 'word/document.xml';
  let d = strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`);
  for (const ns of NS.split(' ')) { const k = ns.split('=')[0]; if (!d.includes(k + '=')) d = d.replace('<w:document ', `<w:document ${ns} `); }
  files[main] = strToU8(d);
  let st = strFromU8(files['word/styles.xml']);
  st = st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, '$1<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>$2');
  files['word/styles.xml'] = strToU8(st);
  fs.writeFileSync(out ? path.join(out, name) : new URL(`./${name}`, import.meta.url), zipSync(files));
}

const p = (t, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${t ? `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>` : ''}</w:p>`;
// A band `h` pt high `top` pt down from its paragraph's top (vertical-relative text), across the page from 55pt.
const rect = (top, h) => `<w:r><w:pict><v:rect style="position:absolute;margin-left:55pt;margin-top:${top}pt;width:439pt;height:${h}pt;z-index:-1;mso-wrap-distance-left:0;mso-wrap-distance-right:0;mso-position-horizontal-relative:page;mso-position-vertical-relative:text" fillcolor="#7f7f7f" stroked="f"><w10:wrap type="topAndBottom" anchorx="page"/></v:rect></w:pict></w:r>`;
const rule = (top) => `<w:r><w:pict><v:shape style="position:absolute;left:0;text-align:left;margin-left:36pt;margin-top:${top}pt;width:537.6pt;height:.1pt;z-index:-1;mso-wrap-distance-left:0;mso-wrap-distance-right:0;mso-position-horizontal-relative:page" coordorigin="720,307" coordsize="10752,0" path="m11472,307l720,307e" filled="f" strokecolor="#231f20" strokeweight=".5pt"><v:path arrowok="t"/><w10:wrap type="topAndBottom" anchorx="page"/></v:shape></w:pict></w:r>`;
const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const variants = {
  // 1113 page 66: an empty paragraph, the band 14pt down, 29pt high
  'b01-empty-rect14': p('Before') + `<w:p>${rect(14.05, 29.15)}</w:p>` + p('After band') + p('Second after'),
  // 1113 page 6: a line of text, a ruled line 15pt down
  'b02-text-rule15': p('Before') + `<w:p>${rule(15.35)}${run('Text of the anchor paragraph')}</w:p>` + p('After rule') + p('Second after'),
  'b03-empty-rect2': p('Before') + `<w:p>${rect(2, 29.15)}</w:p>` + p('After band') + p('Second after'),
  'b04-text-rect30': p('Before') + `<w:p>${rect(30, 20)}${run('Text of the anchor paragraph')}</w:p>` + p('After band') + p('Second after'),
  'b05-text-rect5': p('Before') + `<w:p>${rect(5, 20)}${run('Text of the anchor paragraph')}</w:p>` + p('After band') + p('Second after'),
  'b06-twolines-rect30': p('Before') + `<w:p>${rect(30, 20)}${run('Text of the anchor paragraph')}<w:r><w:br/></w:r>${run('Its second line')}</w:p>` + p('After band') + p('Second after'),
};
if (out) {
  for (const [n, b] of Object.entries(variants)) write(`${n}.docx`, b);
  console.log(Object.keys(variants).length, 'documents in', out);
} else {
  write('band-shapes.docx', variants['b01-empty-rect14'] + variants['b02-text-rule15']);
}
