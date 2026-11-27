// Builds test/fixtures/header-top-bottom.docx (compatibility mode 15): a header whose one paragraph, spaced 6pt before,
// its mark 1pt, holds nothing but a picture wrapped top and bottom, placed against the paragraph 15.5pt up and 59pt tall,
// so that it reaches past the top margin. Word's PDF of it: the picture from 20.5pt to 79.5pt down, the body's first line
// 86.7pt down (its space before again under the picture, then the mark's 1.2pt line); it had come right under the
// picture, 7.2pt high (features.mjs).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

// A 2x1 PNG (a blue and a white pixel).
const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAEUlEQVR4nGNgYPj/n4GBgQEADvcB/0V0Pw0AAAAASUVORK5CYII=', 'base64'));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const picture = '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
  + '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>-406400</wp:posOffset></wp:positionH>'
  + '<wp:positionV relativeFrom="paragraph"><wp:posOffset>-196850</wp:posOffset></wp:positionV>'
  + '<wp:extent cx="1243584" cy="749808"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/>'
  + '<wp:docPr id="1" name="Logo"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
  + '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>'
  + '<pic:blipFill><a:blip r:embed="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
  + '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1243584" cy="749808"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
  + '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
const mark = '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>';
const header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">`
  + `<w:p><w:pPr><w:pStyle w:val="Header"/><w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/>${mark}</w:pPr>${picture}</w:p></w:hdr>`;
const para = (n) => `<w:p><w:r><w:t xml:space="preserve">Paragraph ${n}: the quick brown fox jumps over the lazy dog.</w:t></w:r></w:p>`;
const files = unzipSync(blankDocxBytes());
files['word/header1.xml'] = strToU8(header);
files['word/media/logo.png'] = PNG;
files['word/_rels/header1.xml.rels'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLogo" Type="${R}/image" Target="media/logo.png"/></Relationships>`);
files['word/settings.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${W}"><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`);
let rels = strFromU8(files['word/_rels/document.xml.rels']);
rels = rels.replace('</Relationships>', `<Relationship Id="rIdHdr" Type="${R}/header" Target="header1.xml"/><Relationship Id="rIdSet" Type="${R}/settings" Target="settings.xml"/></Relationships>`);
files['word/_rels/document.xml.rels'] = strToU8(rels);
let ct = strFromU8(files['[Content_Types].xml']);
if (!/Extension="png"/.test(ct)) ct = ct.replace('<Override', '<Default Extension="png" ContentType="image/png"/><Override');
ct = ct.replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
  + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>');
files['[Content_Types].xml'] = strToU8(ct);
let doc = strFromU8(files['word/document.xml']);
doc = doc.replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${Array.from({ length: 30 }, (_, i) => para(i + 1)).join('')}$1`);
doc = doc.replace(/<w:sectPr\b([^>]*)>/, '<w:sectPr$1><w:headerReference w:type="default" r:id="rIdHdr"/>');
doc = doc.replace(/<w:pgMar\b[^>]*\/>/, '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>');
if (!/xmlns:r=/.test(doc)) doc = doc.replace(/<w:document\b/, `<w:document xmlns:r="${R}"`);
files['word/document.xml'] = strToU8(doc);
fs.writeFileSync(new URL('./header-top-bottom.docx', import.meta.url), zipSync(files));
