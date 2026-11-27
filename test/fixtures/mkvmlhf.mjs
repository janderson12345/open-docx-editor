// Builds test/fixtures/vml-hf-boxes.docx: a header and a footer whose lines are floating VML text boxes placed against
// the page (v:shape position:absolute, v:textbox, mso-position-*-relative:page, no wrapping, behind the text), as forms
// made from PDFs have them, all in one paragraph of a 14/240 line. Letter, margins 1500tw top, 280tw bottom, 600tw
// left and right, header and footer 720tw from the edges; Times New Roman 10pt body lines "Line 1".."Line 140".
// Header: "Supplement B" over "Form Instructions" in a box 342pt in, 7.3pt down (two lines), and an empty box.
// Footer: "FORM 10-100 Template draft of box." at 736.1pt (7pt Arial), "TEST 4/20.B" over "Page 1 of 3" at 527pt, and
// "Page N of 68" (a PAGE field) at 754.9pt, centred under the body.
// Word lays the boxes out of the flow: the footer is its one thin paragraph, so the body runs down to the bottom margin.
// usage: node test/fixtures/mkvmlhf.mjs [out dir] [--dml]  (default: here)
import '../setup.js';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const OUT = process.argv[2] && !process.argv[2].startsWith('--') ? path.resolve(process.argv[2]) : path.dirname(fileURLToPath(import.meta.url));
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
const TNR = (sz) => `<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`;
const ARIAL = (sz) => `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`;
const run = (t, rpr = '') => `<w:r><w:rPr>${rpr}</w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
const page = (rpr) => `<w:r><w:rPr>${rpr}</w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr>${rpr}</w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:rPr>${rpr}</w:rPr><w:fldChar w:fldCharType="separate"/></w:r>${run('1', rpr)}<w:r><w:rPr>${rpr}</w:rPr><w:fldChar w:fldCharType="end"/></w:r>`;
const bp = (content, before = 10, left = 20) => `<w:p><w:pPr><w:spacing w:before="${before}"/><w:ind w:left="${left}"/></w:pPr>${content}</w:p>`;
let id = 1000;
/** A floating text box placed against the page, no fill or line, no insets, behind the text. */
const box = (x, y, w, h, paras, type = false) => DML ? dbox(x, y, w, h, paras) : `<w:r><w:pict>${type ? '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' : ''}`
  + `<v:shape id="_x0000_s${id++}" type="#_x0000_t202" style="position:absolute;margin-left:${x}pt;margin-top:${y}pt;width:${w}pt;height:${h}pt;z-index:-${251658240 - id};mso-position-horizontal-relative:page;mso-position-vertical-relative:page" filled="f" stroked="f">`
  + `<v:textbox inset="0,0,0,0"><w:txbxContent>${paras}</w:txbxContent></v:textbox><w10:wrap anchorx="page" anchory="page"/></v:shape></w:pict></w:r>`;
// (With --dml, the same boxes as DrawingML text boxes, vml-hf-boxes-dml.docx: Word's print of both compared.)
const DML = process.argv.includes('--dml');
// (--top=, --header=, --bottom=, --footer=: the margins and header and footer distances in twips; --compat=: a compatibility mode in settings.xml;
// --name=: the file's name. For Word's prints of the variants.)
const arg = (k, d) => process.argv.find((a) => a.startsWith(`--${k}=`))?.slice(k.length + 3) ?? d;
const E = (pt) => Math.round(pt * 12700);
const dbox = (x, y, w, h, paras) => `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${id++}" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>`
  + `<wp:positionH relativeFrom="page"><wp:posOffset>${E(x)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${E(y)}</wp:posOffset></wp:positionV><wp:extent cx="${E(w)}" cy="${E(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="Text Box ${id}"/>`
  + `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${E(w)}" cy="${E(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>`
  + `<wps:txbx><w:txbxContent>${paras}</w:txbxContent></wps:txbx><wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
const thin = (content) => `<w:p><w:pPr><w:spacing w:line="14" w:lineRule="auto"/></w:pPr>${content}</w:p>`;

const header = thin(box(71, 7.3, 171, 15.3, bp(''), true) + box(342.05, 7.3, 199, 29.7, bp(run('Supplement B', TNR(24))) + bp(run('Form Instructions', TNR(24)), 12)));
const footer = thin(box(35, 736.1, 351.05, 10.95, bp(run('FORM 10-100 Template draft of box.', ARIAL(14)), 14), true)
  + box(527, 736.1, 50.05, 22.7, bp(run('TEST 4/20.B', ARIAL(14)), 14) + bp(run('Page 1 of 3', ARIAL(14)), 52, 161))
  + box(271.85, 754.9, 68.35, 15.3, bp(run('Page ', TNR(20)) + page(TNR(20)) + run(' of 68', TNR(20)))));

const files = unzipSync(blankDocxBytes());
const main = 'word/document.xml';
const lines = Array.from({ length: 140 }, (_, i) => `<w:p>${run(`Line ${i + 1}`)}</w:p>`).join('');
// (--table: the lines in one table row that may break across pages.)
const body = process.argv.includes('--table') ? `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>${lines}</w:tc></w:tr></w:tbl><w:p/>` : lines;
const sect = `<w:sectPr><w:headerReference w:type="default" r:id="rIdHd1"/><w:footerReference w:type="default" r:id="rIdFt1"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="${arg('top', 1500)}" w:right="600" w:bottom="${arg('bottom', 280)}" w:left="600" w:header="${arg('header', 720)}" w:footer="${arg('footer', 720)}" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>`;
let d = strFromU8(files[main]).replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${body}${sect}</w:body>`);
d = d.replace(/<w:document\b[^>]*>/, (m) => m.replace(/\s+xmlns:(w|r|v|o|w10|wp|a|wps)="[^"]*"/g, '').replace('<w:document', `<w:document ${W}`));
files[main] = strToU8(d);
files['word/header1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}>${header}</w:hdr>`);
files['word/footer1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${W}>${footer}</w:ftr>`);
const ct = '[Content_Types].xml';
files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'));
const rels = 'word/_rels/document.xml.rels';
files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdHd1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdFt1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>'));
const st = strFromU8(files['word/styles.xml']);
files['word/styles.xml'] = strToU8(st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, `$1<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr>${TNR(20)}</w:rPr>$2`));
const compat = arg('compat', null);
if (compat) {
  files['word/settings.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${Number(compat)}"/></w:compat></w:settings>`);
  files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'));
  files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdSt1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'));
}
fs.writeFileSync(path.join(OUT, arg('name', DML ? 'vml-hf-boxes-dml' : 'vml-hf-boxes') + '.docx'), zipSync(files));
