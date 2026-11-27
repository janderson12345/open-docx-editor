// Builds test/fixtures/tbox-baselines.docx: floating text boxes (VML, placed on the page, no insets) of one line each in
// several fonts and sizes, as Word sets their text (its prints of the same boxes, made here and printed by Word): the
// first baseline the font's line gap and ascent below the box's top (and its space before).
// With a folder (node test/fixtures/mktboxbase.mjs <dir>) it writes every variant tried in Word there instead.
import '../setup.js';
import fs from 'fs';
import path from 'path';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const out = process.argv[2] ? path.resolve(process.argv[2]) : null;
if (out) fs.mkdirSync(out, { recursive: true });
const V = 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"';

function write(name, body) {
  const files = unzipSync(blankDocxBytes());
  const main = 'word/document.xml';
  let d = strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`);
  for (const ns of V.split(' ')) { const k = ns.split('=')[0]; if (!d.includes(k + '=')) d = d.replace('<w:document ', `<w:document ${ns} `); }
  files[main] = strToU8(d);
  let st = strFromU8(files['word/styles.xml']);
  st = st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, '$1<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>$2');
  files['word/styles.xml'] = strToU8(st);
  fs.writeFileSync(out ? path.join(out, name) : new URL(`./${name}`, import.meta.url), zipSync(files));
}

// A box at `top` pt down the page, 72pt in, holding one line "<label> ..." in `font` at `hp` half-points, `before` twips
// before it; `kind`: VML (v:shape) or DrawingML (wps), its insets none.
const vml = (top, label, font, hp, before = 0, exact = 0) => `<w:r><w:pict><v:shape id="b${Math.round(top)}" type="#_x0000_t202" style="position:absolute;margin-left:72pt;margin-top:${top}pt;width:300pt;height:${Math.ceil(hp)}pt;z-index:-1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page" filled="f" stroked="f"><v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:pPr><w:spacing w:before="${before}" w:after="0" w:line="${exact || 240}" w:lineRule="${exact ? 'exact' : 'auto'}"/><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:sz w:val="${hp}"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:sz w:val="${hp}"/></w:rPr><w:t xml:space="preserve">${label} box text line</w:t></w:r></w:p></w:txbxContent></v:textbox><w10:wrap anchorx="page" anchory="page"/></v:shape></w:pict></w:r>`;
const FONTS = [['Arial', [14, 16, 20, 24, 28]], ['Times New Roman', [16, 24]], ['Calibri', [16, 22]]];
const boxes = (before) => {
  let top = 72, runs = '';
  for (const [f, sizes] of FONTS) for (const hp of sizes) { runs += vml(top, `${f.split(' ')[0]}${hp}b${before}`, f, hp, before); top += 40; }
  return `<w:p><w:pPr><w:spacing w:line="14" w:lineRule="auto"/></w:pPr>${runs}</w:p>`;
};
// Lines of an exact height in Arial 10pt and Times New Roman 11pt, shorter and taller than the font's own line.
const exacts = () => {
  let top = 72, runs = '';
  for (const [f, hp] of [['Arial', 20], ['Times New Roman', 22]]) for (const ex of [150, 186, 207, 230, 260, 320]) { runs += vml(top, `${f.split(' ')[0]}${hp}x${ex}`, f, hp, 0, ex); top += 36; }
  return `<w:p><w:pPr><w:spacing w:line="14" w:lineRule="auto"/></w:pPr>${runs}</w:p>`;
};
const variants = { 'b01-before0': boxes(0), 'b02-before14': boxes(14), 'b03-before52': boxes(52), 'b04-exact': exacts() };
if (out) {
  for (const [n, b] of Object.entries(variants)) write(`${n}.docx`, b);
  console.log(Object.keys(variants).length, 'documents in', out);
} else {
  write('tbox-baselines.docx', variants['b02-before14']);
}
