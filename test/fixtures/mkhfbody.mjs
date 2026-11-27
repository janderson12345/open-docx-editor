// Builds test/fixtures/hf-pic-tabs.docx, hf-end-break.docx and hf-zero-rule.docx: headers whose height decides where the
// body starts, as Word lays them out (its PDFs of documents made like these, features.mjs). Calibri 11pt, the header
// 0.25in from the page's top, the top margin 0.5in, so the header's foot is the body's top:
// - two pictures with two tabs between them in one header paragraph: the line is as tall as the taller picture, the
//   tabs' text adds no descent under it (body at 110px in Word; it had been 4.4px lower);
// - a header paragraph closing with a line break over a 6pt mark: the empty line after the break is the mark's 6pt line
//   (body at 69.6px; it had been 9.8px higher, that line lost);
// - a header paragraph holding only a line shape of no height wrapped top and bottom, 13pt below its top: Word keeps no
//   text clear of it (body at 77.7px, three header lines; it had been 13.1px lower);
// - (hf-ptab-logo.docx, the header 0.3in down, the top margin 1in) a 12pt header line, a right positional tab, a logo
//   placed 0.3in from the page's top wrapped tight (82pt tall, 9pt side distances), spaces and another positional tab:
//   Word sets the spaces' line below the logo, 12px lower, and the tab's under it (body at 155.6px; it had started at
//   the top margin, 96px);
// - (hf-ptabs.docx, 1in margins) "Left words", a centred positional tab, "Centre words", a right one, "Right words":
//   Word's PDF has the centre words' middle at 408.2px and the right words' end at 720.3px, the margin (they had run
//   on after the left words: the tabs had no width).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

// A 2x1 PNG (a blue and a white pixel).
const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAEUlEQVR4nGNgYPj/n4GBgQEADvcB/0V0Pw0AAAAASUVORK5CYII=', 'base64'));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NSS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"`;
const inline = (cx, cy) => '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
  + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Logo"/><wp:cNvGraphicFramePr/>`
  + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>'
  + `<pic:blipFill><a:blip r:embed="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
  + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
const rule = (off) => '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
  + `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${off}</wp:posOffset></wp:positionV>`
  + '<wp:extent cx="5943600" cy="0"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/><wp:docPr id="2" name="Line 2"/><wp:cNvGraphicFramePr/>'
  + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5943600" cy="0"/></a:xfrm>'
  + '<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="6350"><a:solidFill><a:srgbClr val="2C358B"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
const txt = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const hp = (runs, ppr = '') => `<w:p><w:pPr><w:pStyle w:val="Header"/>${ppr}</w:pPr>${runs}</w:p>`;
const para = (n) => `<w:p><w:r><w:t xml:space="preserve">Paragraph ${n}: the quick brown fox jumps over the lazy dog.</w:t></w:r></w:p>`;
const STYLES = '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs></w:pPr></w:style>';

function build(name, header, { top = 720, hdr = 360 } = {}) {
  const files = unzipSync(blankDocxBytes());
  files['word/header1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NSS}>${header}</w:hdr>`);
  files['word/media/logo.png'] = PNG;
  files['word/_rels/header1.xml.rels'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLogo" Type="${R}/image" Target="media/logo.png"/></Relationships>`);
  files['word/_rels/document.xml.rels'] = strToU8(strFromU8(files['word/_rels/document.xml.rels']).replace('</Relationships>', `<Relationship Id="rIdHdr" Type="${R}/header" Target="header1.xml"/></Relationships>`));
  let ct = strFromU8(files['[Content_Types].xml']);
  if (!/Extension="png"/.test(ct)) ct = ct.replace('<Override', '<Default Extension="png" ContentType="image/png"/><Override');
  files['[Content_Types].xml'] = strToU8(ct.replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'));
  let st = strFromU8(files['word/styles.xml']);
  st = st.replace(/<w:docDefaults>[\s\S]*?<\/w:docDefaults>/, '').replace(/<w:style w:type="paragraph"[\s\S]*?<\/w:style>/g, '');
  files['word/styles.xml'] = strToU8(st.replace(/(<w:styles\b[^>]*>)/, `$1${STYLES}`));
  let doc = strFromU8(files['word/document.xml']);
  doc = doc.replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${[1, 2, 3, 4, 5].map(para).join('')}$1`);
  doc = doc.replace(/<w:sectPr\b([^>]*)>/, '<w:sectPr$1><w:headerReference w:type="default" r:id="rIdHdr"/>');
  doc = doc.replace(/<w:pgMar\b[^>]*\/>/, `<w:pgMar w:top="${top}" w:right="1440" w:bottom="1440" w:left="1440" w:header="${hdr}" w:footer="720" w:gutter="0"/>`);
  files['word/document.xml'] = strToU8(doc.replace(/<w:document\b[^>]*>/, `<w:document ${NSS}>`));
  fs.writeFileSync(new URL(`./${name}`, import.meta.url), zipSync(files));
}

const one = hp(txt('Header line one'));
const three = hp(txt('Header line three'));
build('hf-pic-tabs.docx', one + hp(inline(1075690, 479070) + '<w:r><w:tab/></w:r><w:r><w:tab/></w:r>' + inline(1475690, 429070)) + three);
build('hf-end-break.docx', one + hp(txt('Closing with a break') + '<w:r><w:br/></w:r>', '<w:rPr><w:sz w:val="12"/></w:rPr>'));
build('hf-zero-rule.docx', one + hp(rule(125095)) + three);
const logo = '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="3" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
  + '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>5486400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>274320</wp:posOffset></wp:positionV>'
  + '<wp:extent cx="932180" cy="749300"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>'
  + '<wp:docPr id="3" name="Logo 3"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="3" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>'
  + '<pic:blipFill><a:blip r:embed="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="932180" cy="749300"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
const ptab = '<w:r><w:ptab w:alignment="right" w:relativeTo="margin" w:leader="none"/></w:r>';
build('hf-ptab-logo.docx', hp('<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Test Header Organization Name</w:t></w:r>' + ptab + logo + txt('     ') + ptab), { top: 1440, hdr: 432 });
const ptabAt = (al) => `<w:r><w:ptab w:alignment="${al}" w:relativeTo="margin" w:leader="none"/></w:r>`;
build('hf-ptabs.docx', hp(txt('Left words') + ptabAt('center') + txt('Centre words') + ptabAt('right') + txt('Right words')), { top: 1440, hdr: 720 });
