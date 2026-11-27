// Builds test/fixtures/cover-lines.docx, list-bullet-atleast.docx, behind-shapes-order.docx and picture-bilevel.docx: a cover page's lines,
// bullets under "at least" spacing and shapes behind the text as Word lays them (its own layout of the same documents
// made in Word, odx-test mkcover.ps1 and mkbulletatleast.ps1). Normal: Calibri 12pt, single, no spacing.
// cover-lines: "Opens" in a 48pt style, a bare line break (no size of its own) then two lines of 36pt text: Word's
//   lines 58.5, 44.25 and 43.5pt (the break's line in the style's 48pt; 195px in all). Then "Ends": a 20pt line and a
//   20pt "September 28, 2026" closing on a 36pt space: 48.75pt (65px), the space sizing nothing.
// list-bullet-atleast: 30 one-line items under a Symbol bullet, 10pt Calibri, at least 13pt: 13pt apart (17.33px), the
//   bullet's reach inside the set height (single, the bullet makes the line 12.75pt).
// behind-shapes-order: three coloured rectangles behind the text, their drawing order (relativeHeight) the reverse of the
//   document's: all drawn over the page, not under its white sheet. And a fourth, 600pt wide from the column: cut off at
//   the page's right edge, as Word shows it.
// picture-bilevel: a grey picture shown in black and white (a:biLevel), below.
import '../setup.js';
import fs from 'fs';
import zlib from 'zlib';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SINGLE = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
const CAL = (sz) => `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`;
const run = (t, rpr = '') => `<w:r><w:rPr>${rpr}</w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
const brk = (rpr = '') => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:br/></w:r>`;
const para = (content, ppr = '') => `<w:p><w:pPr>${ppr}</w:pPr>${content}</w:p>`;

function write(name, { body, styles = '', numbering = null }) {
  const files = unzipSync(blankDocxBytes());
  const main = 'word/document.xml';
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
  let st = strFromU8(files['word/styles.xml']);
  st = st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, `$1<w:pPr>${SINGLE}</w:pPr><w:rPr>${CAL(24)}</w:rPr>$2`);
  files['word/styles.xml'] = strToU8(st.replace('</w:styles>', `${styles}</w:styles>`));
  if (numbering) {
    files['word/numbering.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}>${numbering}</w:numbering>`);
    const ct = '[Content_Types].xml', rels = 'word/_rels/document.xml.rels';
    if (!strFromU8(files[ct]).includes('/word/numbering.xml')) files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'));
    if (!strFromU8(files[rels]).includes('numbering.xml')) files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'));
  }
  fs.writeFileSync(new URL(`./${name}`, import.meta.url), zipSync(files));
}

// cover-lines
write('cover-lines.docx', {
  styles: `<w:style w:type="paragraph" w:customStyle="1" w:styleId="BigTitle"><w:name w:val="Big Title"/><w:basedOn w:val="Normal"/><w:rPr>${CAL(96)}</w:rPr></w:style>`,
  body: para(run('Above'))
    + para(brk() + run('Opens title text', CAL(72)) + brk(CAL(72)) + run('Two', CAL(72)), `<w:pStyle w:val="BigTitle"/><w:rPr>${CAL(72)}</w:rPr>`)
    + para(run('Ends first line', CAL(40)) + brk(CAL(40)) + run('September 28, 2026', CAL(40)) + run(' ', CAL(72)), `<w:rPr>${CAL(72)}</w:rPr>`)
    + para(run('Below')),
});

// list-bullet-atleast
const lvl = `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>`;
write('list-bullet-atleast.docx', {
  numbering: `<w:abstractNum w:abstractNumId="0">${lvl}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`,
  body: Array.from({ length: 30 }, (_, i) => para(run(`Item line ${i + 1}`, CAL(20)), `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="0" w:after="0" w:line="260" w:lineRule="atLeast"/><w:rPr>${CAL(20)}</w:rPr>`)).join(''),
});

// behind-shapes-order
const EMU = 12700;
const rect = (n, rh, color, y, w = 200) => `<w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${rh}" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">`
  + `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${y * EMU}</wp:posOffset></wp:positionV>`
  + `<wp:extent cx="${w * EMU}" cy="${60 * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${n}" name="Rectangle ${n}"/><wp:cNvGraphicFramePr/>`
  + '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
  + `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w * EMU}" cy="${60 * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`
  + `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
write('behind-shapes-order.docx', {
  body: para(rect(1, 3, '23D2D7', 0) + run('First, on a cyan box')) + para('') + para('') + para('') + para('') + para('')
    + para(rect(2, 2, 'F4B183', 0) + run('Second, on an orange box')) + para('') + para('') + para('') + para('') + para('')
    + para(rect(3, 1, 'A9D18E', 0) + run('Third, on a green box')) + para('') + para('') + para('') + para('') + para('')
    + para(rect(4, 4, 'BDD7EE', 0, 600) + run('Fourth, on a blue box 600pt wide, past the right edge')),
});

// picture-bilevel: a 2x1 picture, its left pixel mid grey (#808080) and its right light grey (#E0E0E0), shown in black and
// white at 75% (a:biLevel thresh="75000"), 200x100pt in line: Word draws the left half black and the right half white.
const png = (rgb) => {
  const crc = (b) => zlib.crc32(b) >>> 0;
  const chunk = (type, data) => {
    const t = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const c = Buffer.alloc(4); c.writeUInt32BE(crc(t));
    return Buffer.concat([len, t, c]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(rgb.length, 0); ihdr.writeUInt32BE(1, 4); ihdr[8] = 8; ihdr[9] = 2; // 8-bit RGB
  const raw = Buffer.from([0, ...rgb.flat()]);
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
};
{
  const files = unzipSync(blankDocxBytes());
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const cx = 200 * EMU, cy = 100 * EMU;
  const pic = '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">'
    + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Picture 1"/><wp:cNvGraphicFramePr/>`
    + '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    + '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="grey.png"/><pic:cNvPicPr/></pic:nvPicPr>'
    + `<pic:blipFill><a:blip xmlns:r="${R}" r:embed="rIdGrey"><a:biLevel thresh="75000"/></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`
    + '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const main = 'word/document.xml';
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${para(pic)}${para(run('Under the picture'))}$1`));
  files['word/media/grey.png'] = new Uint8Array(png([[128, 128, 128], [224, 224, 224]]));
  const rels = 'word/_rels/document.xml.rels';
  files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', `<Relationship Id="rIdGrey" Type="${R}/image" Target="media/grey.png"/></Relationships>`));
  const ct = '[Content_Types].xml';
  if (!/Extension="png"/.test(strFromU8(files[ct]))) files[ct] = strToU8(strFromU8(files[ct]).replace('<Override', '<Default Extension="png" ContentType="image/png"/><Override'));
  fs.writeFileSync(new URL('./picture-bilevel.docx', import.meta.url), zipSync(files));
}
