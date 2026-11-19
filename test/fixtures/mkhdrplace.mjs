// Builds test/fixtures/header-page-placed.docx: a first-page header only (w:titlePg, no default header), its header
// distance (1in) far below a small top margin (0.26in), and in it a picture placed against the page, 0.5in down, wrapped
// square, as the samples' header banners are. Word draws the picture at 0.5in, the first page's body under the header's
// line, and pages 2 on (no header) from the top margin. The picture had been drawn at the header's top, its paragraph
// pushed down under it with the body, and pages without a header still had their text from the header distance down
// (features.mjs).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

// A 2x1 PNG (a blue and a white pixel).
const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAEUlEQVR4nGNgYPj/n4GBgQEADvcB/0V0Pw0AAAAASUVORK5CYII=', 'base64'));
const EMU = 914400;
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const picture = `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">`
  + `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>${EMU}</wp:posOffset></wp:positionH>`
  + `<wp:positionV relativeFrom="page"><wp:posOffset>${EMU / 2}</wp:posOffset></wp:positionV>`
  + `<wp:extent cx="${4 * EMU}" cy="${EMU / 2}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>`
  + `<wp:docPr id="1" name="Banner"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`
  + `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="banner.png"/><pic:cNvPicPr/></pic:nvPicPr>`
  + `<pic:blipFill><a:blip r:embed="rIdBanner"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
  + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${4 * EMU}" cy="${EMU / 2}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`
  + '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
const header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>${picture}</w:p></w:hdr>`;
const para = (n) => `<w:p><w:r><w:t xml:space="preserve">Paragraph ${n}: the quick brown fox jumps over the lazy dog, and the five boxing wizards jump quickly over it again.</w:t></w:r></w:p>`;
const files = unzipSync(blankDocxBytes());
files['word/header1.xml'] = strToU8(header);
files['word/media/banner.png'] = PNG;
files['word/_rels/header1.xml.rels'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdBanner" Type="${R}/image" Target="media/banner.png"/></Relationships>`);
let rels = strFromU8(files['word/_rels/document.xml.rels']);
rels = rels.replace('</Relationships>', `<Relationship Id="rIdHdrFirst" Type="${R}/header" Target="header1.xml"/></Relationships>`);
files['word/_rels/document.xml.rels'] = strToU8(rels);
let ct = strFromU8(files['[Content_Types].xml']);
if (!/Extension="png"/.test(ct)) ct = ct.replace('<Override', '<Default Extension="png" ContentType="image/png"/><Override');
ct = ct.replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>');
files['[Content_Types].xml'] = strToU8(ct);
let doc = strFromU8(files['word/document.xml']);
const body = Array.from({ length: 60 }, (_, i) => para(i + 1)).join('');
doc = doc.replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`);
doc = doc.replace(/<w:sectPr\b([^>]*)>/, `<w:sectPr$1><w:headerReference w:type="first" r:id="rIdHdrFirst"/>`);
doc = doc.replace(/<w:pgMar\b[^>]*\/>/, '<w:pgMar w:top="374" w:right="1440" w:bottom="1440" w:left="1440" w:header="1440" w:footer="720" w:gutter="0"/>');
doc = doc.replace(/(<\/w:sectPr>)/, '<w:titlePg/>$1');
if (!/xmlns:r=/.test(doc)) doc = doc.replace(/<w:document\b/, `<w:document xmlns:r="${R}"`);
files['word/document.xml'] = strToU8(doc);
fs.writeFileSync(new URL('./header-page-placed.docx', import.meta.url), zipSync(files));
