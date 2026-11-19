// Builds test/fixtures/float-room.docx (tables beside floats), header-float.docx (a header's floats) and sect-hang.docx
// (a section's closing paragraph holding only a picture), as Word lays them out (its PDFs and layout of documents made
// like these, features.mjs):
// - a logo placed against the page at its top (14pt down, 47.5pt into the left margin's text side), wrapped square, in
//   an empty first paragraph, and under it a table indented 153pt whose right edge is 4.5pt past the right margin: it
//   doesn't fit in the room beside the logo, so Word starts it under the logo (it had stayed beside it);
// - a logo placed from its paragraph, wrapped square, and a table with no indent that fits beside it: Word moves it over
//   to the logo's wrap distance and keeps it there;
// - a header holding a logo wrapped square beside its paragraph (60pt tall, reaching 26pt into the body's top margin)
//   and a floating table placed 10pt above the empty paragraph after it (tblpY -200): Word starts the body at its top
//   margin all the same (it had moved under the logo) and draws the table 10pt up (it had been at that paragraph), beside
//   the logo;
// - a section's closing paragraph (next page) holding only a picture wrapped square, its top 5pt above the footer's room
//   and its foot 5pt into it: Word keeps it on its page, as it does an empty one (it had gone on alone, a blank page
//   before the next section's).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

// A 2x1 PNG (a blue and a white pixel).
const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAEUlEQVR4nGNgYPj/n4GBgQEADvcB/0V0Pw0AAAAASUVORK5CYII=', 'base64'));
const PT = 12700; // EMU
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
let id = 0;
const logo = ({ x, y, w, h, hRel = 'margin', vRel = 'page', rel = 'rIdLogo' }) => {
  id++;
  return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">`
    + `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${hRel}"><wp:posOffset>${Math.round(x * PT)}</wp:posOffset></wp:positionH>`
    + `<wp:positionV relativeFrom="${vRel}"><wp:posOffset>${Math.round(y * PT)}</wp:posOffset></wp:positionV>`
    + `<wp:extent cx="${Math.round(w * PT)}" cy="${Math.round(h * PT)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>`
    + `<wp:docPr id="${id}" name="Logo ${id}"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`
    + `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>`
    + `<pic:blipFill><a:blip r:embed="${rel}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(w * PT)}" cy="${Math.round(h * PT)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`
    + '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
};
const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const para = (content = '', ppr = '') => `<w:p><w:pPr>${ppr}</w:pPr>${content}</w:p>`;
const small = '<w:spacing w:after="0"/><w:rPr><w:sz w:val="16"/></w:rPr>';
const table = (ind, width, label, float = '') => `<w:tbl><w:tblPr>${float}<w:tblW w:w="${width}" w:type="dxa"/><w:tblInd w:w="${ind}" w:type="dxa"/>`
  + '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>'
  + `<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="${width}"/></w:tblGrid>`
  + [1, 2].map((r) => `<w:tr><w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${para(run(`${label}, row ${r}`), '<w:spacing w:after="0"/>')}</w:tc></w:tr>`).join('')
  + '</w:tbl>';
const text = (n) => para(run(`Paragraph ${n}: the quick brown fox jumps over the lazy dog, and the five boxing wizards jump quickly.`));

const header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">`
  + para(logo({ x: 3.3, y: 1.7, w: 187, h: 60, vRel: 'paragraph', rel: 'rIdHLogo' }), '<w:spacing w:after="0"/>')
  + table(0, 2500, 'Header table', '<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="right" w:tblpY="-200"/><w:tblOverlap w:val="never"/>')
  + para('', '<w:spacing w:after="0"/>') + '</w:hdr>';
const body = [
  para(logo({ x: -42.5, y: 14, w: 80.6, h: 71.4 }), small),
  table(3060, 6390, 'Indented past the room'),
  ...[1, 2, 3, 4].map(text),
  para(logo({ x: -42.5, y: 0, w: 80.6, h: 71.4, vRel: 'paragraph' }), small),
  table(0, 4000, 'Fits beside'),
  ...[5, 6, 7, 8, 9, 10].map(text),
].join('');

function build(name, body, hdr, ftr) {
  const files = unzipSync(blankDocxBytes());
  files['word/media/logo.png'] = PNG;
  let rels = strFromU8(files['word/_rels/document.xml.rels']);
  rels = rels.replace('</Relationships>', `<Relationship Id="rIdLogo" Type="${R}/image" Target="media/logo.png"/>${hdr ? `<Relationship Id="rIdHdr" Type="${R}/header" Target="header1.xml"/>` : ''}${ftr ? `<Relationship Id="rIdFtr" Type="${R}/footer" Target="footer1.xml"/>` : ''}</Relationships>`);
  files['word/_rels/document.xml.rels'] = strToU8(rels);
  let ct = strFromU8(files['[Content_Types].xml']);
  if (!/Extension="png"/.test(ct)) ct = ct.replace('<Override', '<Default Extension="png" ContentType="image/png"/><Override');
  if (hdr) {
    files['word/header1.xml'] = strToU8(hdr);
    files['word/_rels/header1.xml.rels'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHLogo" Type="${R}/image" Target="media/logo.png"/></Relationships>`);
    ct = ct.replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>');
  }
  if (ftr) {
    files['word/footer1.xml'] = strToU8(ftr);
    ct = ct.replace('</Types>', '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>');
  }
  files['[Content_Types].xml'] = strToU8(ct);
  let doc = strFromU8(files['word/document.xml']);
  // (The body's own section properties, the last: a section break in the body keeps its own.)
  let sect = doc.slice(doc.lastIndexOf('<w:sectPr'));
  if (hdr) sect = sect.replace(/<w:sectPr\b([^>]*)>/, `<w:sectPr$1><w:headerReference w:type="default" r:id="rIdHdr"/>`);
  if (ftr) sect = sect.replace(/<w:sectPr\b([^>]*)>/, `<w:sectPr$1><w:footerReference w:type="default" r:id="rIdFtr"/>`);
  sect = sect.replace(/<w:pgMar\b[^>]*\/>/, `<w:pgMar w:top="1440" w:right="1440" w:bottom="${ftr ? 720 : 1440}" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>`);
  doc = doc.slice(0, doc.lastIndexOf('<w:sectPr')).replace(/<w:body>[\s\S]*$/, `<w:body>${body}`) + sect;
  if (!/xmlns:r=/.test(doc)) doc = doc.replace(/<w:document\b/, `<w:document xmlns:r="${R}"`);
  if (!/xmlns:wp=/.test(doc)) doc = doc.replace(/<w:document\b/, '<w:document xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"');
  files['word/document.xml'] = strToU8(doc);
  fs.writeFileSync(new URL(`./${name}`, import.meta.url), zipSync(files));
}
build('float-room.docx', body);
build('header-float.docx', [1, 2, 3, 4, 5, 6].map(text).join(''), header);

// (Exact lines: the page's text ends at 792 - 36 - 14 = 742pt, the 665pt line ends at 737pt, the closing paragraph's
// 10pt line runs to 747pt.)
const exact = (pt) => `<w:spacing w:after="0" w:line="${pt * 20}" w:lineRule="exact"/>`;
const footer = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="${W}">${para(run('Footer words'), exact(14))}</w:ftr>`;
const closing = '<w:sectPr><w:footerReference w:type="default" r:id="rIdFtr"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="720" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
build('sect-hang.docx', para(run('Filler'), exact(665)) + para(logo({ x: 0, y: 0, w: 100, h: 50, vRel: 'paragraph' }), exact(10) + closing) + para(run('Next section words')), null, footer);
