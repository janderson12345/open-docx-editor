// Builds the test documents for the smaller things Word draws: text effects (outline, shadow, emboss, imprint),
// emphasis marks, run borders, a reflection and a glow (effects.docx); legal numbering and picture bullets (numbering-legal.docx); a
// section's vertical alignment (valign.docx); line numbers (line-numbers.docx); and text and picture watermarks
// (watermark-text.docx, watermark-picture.docx). Each was compared with Word's own PDF of it.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, zlibSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const VML = 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"';
const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;');
const run = (t, rpr = '') => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
const para = (runs, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${runs}</w:p>`;
const LONG = 'The quick brown fox jumps over the lazy dog, and then it runs on to the end of the line so that the paragraph wraps onto a second line. ';
const SECT = (extra = '') => `<w:sectPr>${extra.headers || ''}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>${extra.lnNum || ''}<w:cols w:space="720"/>${extra.vAlign ? `<w:vAlign w:val="${extra.vAlign}"/>` : ''}</w:sectPr>`;

/** A small PNG of one colour (for the picture bullet and the picture watermark). */
function png(w, h, [r, g, b]) {
  const crcTable = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
  const crc = (bytes) => { let c = 0xffffffff; for (const x of bytes) c = crcTable[(c ^ x) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const chunk = (type, data) => {
    const out = new Uint8Array(12 + data.length), dv = new DataView(out.buffer);
    dv.setUint32(0, data.length); out.set(strToU8(type), 4); out.set(data, 8);
    dv.setUint32(8 + data.length, crc(out.subarray(4, 8 + data.length)));
    return out;
  };
  const ihdr = new Uint8Array(13), hv = new DataView(ihdr.buffer);
  hv.setUint32(0, w); hv.setUint32(4, h); ihdr.set([8, 2, 0, 0, 0], 8);
  const raw = new Uint8Array(h * (1 + 3 * w));
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) raw.set([r, g, b], y * (1 + 3 * w) + 1 + 3 * x);
  const parts = [new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlibSync(raw)), chunk('IEND', new Uint8Array())];
  const out = new Uint8Array(parts.reduce((t, p) => t + p.length, 0));
  let at = 0; for (const p of parts) { out.set(p, at); at += p.length; }
  return out;
}

/** Writes `name`.docx: the blank document with this body, and extra parts { path: string | bytes } and relationships. */
function build(name, body, { parts = {}, rels = '', types = '' } = {}) {
  const files = unzipSync(blankDocxBytes());
  files['word/document.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W} ${VML}><w:body>${body}</w:body></w:document>`);
  if (rels) files['word/_rels/document.xml.rels'] = strToU8(strFromU8(files['word/_rels/document.xml.rels']).replace('</Relationships>', rels + '</Relationships>'));
  if (types) files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace('</Types>', types + '</Types>'));
  for (const [p, v] of Object.entries(parts)) files[p] = typeof v === 'string' ? strToU8(v) : v;
  fs.writeFileSync(new URL(`./${name}.docx`, import.meta.url), zipSync(files));
}

// Text effects, emphasis marks and run borders, each inside a paragraph long enough to wrap.
build('effects', [
  ...[['outline', '<w:outline/>'], ['shadow', '<w:shadow/>'], ['emboss', '<w:emboss/>'], ['imprint', '<w:imprint/>'],
    ['outline and shadow', '<w:outline/><w:shadow/>'], ['coloured outline', '<w:outline/><w:color w:val="C00000"/>'],
  ].map(([label, rpr]) => para(run(`Text with ${label}: `) + run('Effect Letters 123', `${rpr}<w:sz w:val="36"/>`) + run(' ' + LONG))),
  ...['dot', 'comma', 'circle', 'underDot'].map((em) => para(run(`Emphasis mark ${em}: `) + run('marked words here', `<w:em w:val="${em}"/>`) + run(' ' + LONG))),
  para(run('A run border: ') + run('bordered words', '<w:bdr w:val="single" w:sz="8" w:space="0" w:color="auto"/>') + run(' and then the line goes on. ' + LONG)),
  para(run('A run border with space: ') + run('spaced border', '<w:bdr w:val="single" w:sz="12" w:space="4" w:color="2F5496"/>') + run(' and then the line goes on. ' + LONG)),
  para(run('Two runs sharing one border: ') + run('first half ', '<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/>') + run('second half', '<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:b/>') + run(' after. ' + LONG)),
  // Word 2010's reflection (Text Effects > Reflection, its first: "Tight Reflection: Touching"), last so the
  // paragraphs above keep their places.
  para(run('Reflected words', '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="72"/><w14:reflection xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:blurRad="6350" w14:stA="55000" w14:stPos="0" w14:endA="300" w14:endPos="45500" w14:dist="0" w14:dir="5400000" w14:fadeDir="5400000" w14:sx="100000" w14:sy="-100000" w14:kx="0" w14:ky="0" w14:algn="bl"/>') + run(' and plain after')),
  // Word 2010's glow, as in effects-word/effects14b.docx's "Glow 10 at 36": 10pt of 00B050 at 40%.
  para(run('Glowing words', '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="72"/><w14:glow xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:rad="127000"><w14:srgbClr w14:val="00B050"><w14:alpha w14:val="60000"/></w14:srgbClr></w14:glow>') + run(' and plain after')),
  // An inner shadow (the WordArt gallery's Style 15: a w14:shadow scaled to nothing, half-transparent black, up and
  // left), after the reflection for the same reason.
  para(run('Pressed words', '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:color w:val="E8E8E8"/><w:sz w:val="72"/><w14:shadow xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:blurRad="63500" w14:dist="50800" w14:dir="13500000" w14:sx="0" w14:sy="0" w14:kx="0" w14:ky="0" w14:algn="none"><w14:srgbClr w14:val="000000"><w14:alpha w14:val="50000"/></w14:srgbClr></w14:shadow>') + run(' and plain after')),
  SECT(),
].join(''));

// Legal numbering: level 1 in Roman numerals; level 2 (isLgl) shows "1.1", level 3 (isLgl, letters) its own way.
// Picture bullets: a list whose bullet is a picture (w:lvlPicBulletId).
const lvl = (ilvl, fmt, text, extra = '') => `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/>${extra.includes('isLgl') ? '<w:isLgl/>' : ''}<w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/>${extra.includes('pic') ? '<w:lvlPicBulletId w:val="0"/>' : ''}<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (ilvl + 1)}" w:hanging="${ilvl ? 720 : 360}"/></w:pPr>${extra.includes('pic') ? '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>' : ''}</w:lvl>`;
const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W} ${VML}>`
  + '<w:numPicBullet w:numPicBulletId="0"><w:pict><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:9pt;height:9pt" o:bullet="t"><v:imagedata r:id="rIdPic1" o:title=""/></v:shape></w:pict></w:numPicBullet>'
  + `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>${lvl(0, 'upperRoman', '%1.')}${lvl(1, 'decimal', '%1.%2.', 'isLgl')}${lvl(2, 'lowerLetter', '%1.%2.%3', 'isLgl')}${lvl(3, 'decimal', '%1.%2.%3.%4')}</w:abstractNum>`
  + `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${lvl(0, 'bullet', '', 'pic')}</w:abstractNum>`
  + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>';
const li = (numId, ilvl, t) => para(run(t), `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`);
build('numbering-legal', [
  para(run('Legal numbering')),
  li(1, 0, 'First chapter'), li(1, 1, 'A section under it'), li(1, 2, 'A clause in letters'), li(1, 3, 'A level without isLgl'),
  li(1, 1, 'Another section'), li(1, 0, 'Second chapter'), li(1, 1, 'Its first section'), li(1, 2, 'Its clause'),
  para(run('Picture bullets')), li(2, 0, 'A picture bullet'), li(2, 0, 'Another, with a longer line of text that goes on far enough to wrap onto the next line under the first'),
  SECT(),
].join(''), {
  parts: {
    'word/numbering.xml': numbering,
    'word/_rels/numbering.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPic1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/bullet.png"/></Relationships>',
    'word/media/bullet.png': png(12, 12, [0x2f, 0x80, 0x40]),
  },
});

// A section's vertical alignment: centred, at the bottom, and spread over the page (both).
const sectionEnd = (vAlign) => para('', `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:vAlign w:val="${vAlign}"/></w:sectPr>`);
build('valign', [
  para(run('A centred section: these three paragraphs sit in the middle of the page.')), para(run(LONG)), para(run('The last of them.')), sectionEnd('center'),
  para(run('A section at the bottom of the page.')), para(run(LONG)), sectionEnd('bottom'),
  para(run('A section spread over the page (both): its paragraphs are spaced out to fill it.')), para(run(LONG)), para(run('Third.')), para(run('Fourth and last.')),
  SECT({ vAlign: 'both' }),
].join(''));

// The layout's cases that only the doc-generator corpus had: a one-section document in two columns centred on its page
// (pages.js verticalAlignCols); one in two columns before its endnotes, its columns ended level and the endnote under
// them (balanceEndCols); and sections starting on an odd and an even page (parityPage), each with a blank page before.
const colsSect = (extra = '') => `<w:sectPr>${extra}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:num="2" w:space="720"/></w:sectPr>`;
build('valign-columns', [
  para(run('Columns centred on the page: this text sits in the middle of it.')), para(run(LONG)), para(run(LONG)),
  colsSect().replace('</w:sectPr>', '<w:vAlign w:val="center"/></w:sectPr>'),
].join(''));
const EN_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const enSeps = '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>';
build('columns-endnotes', [
  ...Array.from({ length: 6 }, (_, i) => para(run(`Paragraph ${i + 1}. ${LONG}${LONG}`))),
  para(run('The last paragraph, with the endnote.') + '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="1"/></w:r>'),
  colsSect(),
].join(''), {
  parts: { 'word/endnotes.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes ${EN_W}>${enSeps}<w:endnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r>${run(' The endnote, under both columns.')}</w:p></w:endnote></w:endnotes>` },
  rels: '<Relationship Id="rIdEn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>',
  types: '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>',
});
const typed = (type) => `<w:sectPr><w:type w:val="${type}"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>`;
build('odd-even-breaks', [
  para(run('First section, on page 1.'), SECT()),
  para(run('Second section: it starts on an odd page, page 3, after a blank page 2.'), typed('oddPage')),
  para(run('Third section: it starts on an even page, page 4.')),
  typed('evenPage'),
].join(''));

// Line numbers: every line, restarting each page, in the first section; every second line, continuing, in the second;
// from w:start in the third (restarting there); a paragraph that suppresses them.
const lnSect = (lnNum) => para('', `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>${lnNum}<w:cols w:space="720"/></w:sectPr>`);
build('line-numbers', [
  ...Array.from({ length: 22 }, (_, i) => para(run(`${i + 1}. ${LONG}`), i === 3 ? '<w:suppressLineNumbers/>' : '')),
  lnSect('<w:lnNumType w:countBy="1" w:restart="newPage"/>'),
  ...Array.from({ length: 6 }, (_, i) => para(run(`Second section ${i + 1}. ${LONG}`))),
  lnSect('<w:lnNumType w:countBy="2" w:distance="567" w:restart="continuous"/>'),
  ...Array.from({ length: 3 }, (_, i) => para(run(`Third section ${i + 1}. ${LONG}`))),
  SECT({ lnNum: '<w:lnNumType w:countBy="1" w:start="4" w:restart="newSection"/>' }),
].join(''));

// Watermarks: Word's text watermark (a VML WordArt shape in the header, rotated) and its picture watermark (washed out).
const header = (inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W} ${VML}><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:pict>${inner}</w:pict></w:r></w:p></w:hdr>`;
const textMark = '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas><v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/><v:textpath on="t" fitshape="t"/><o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>'
  + '<v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s1025" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:412.4pt;height:247.45pt;rotation:315;z-index:-251658752;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="DRAFT"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape>';
const picMark = '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>'
  + '<v:shape id="WordPictureWatermark1" o:spid="_x0000_s1026" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:300pt;height:150pt;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f"><v:imagedata r:id="rIdWm1" o:title="" gain="19661f" blacklevel="22938f"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape>';
const HDR_REL = '<Relationship Id="rIdHdr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>';
const HDR_TYPE = '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>';
const marked = (title) => [para(run(title)), ...Array.from({ length: 12 }, () => para(run(LONG))), SECT({ headers: '<w:headerReference w:type="default" r:id="rIdHdr1"/>' })].join('');
build('watermark-text', marked('A document with a text watermark.'), { parts: { 'word/header1.xml': header(textMark) }, rels: HDR_REL, types: HDR_TYPE });
build('watermark-picture', marked('A document with a picture watermark.'), {
  parts: {
    'word/header1.xml': header(picMark),
    'word/_rels/header1.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdWm1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/mark.png"/></Relationships>',
    'word/media/mark.png': png(40, 20, [0x1f, 0x4e, 0x99]),
  },
  rels: HDR_REL, types: HDR_TYPE,
});

// One run mixing three alphabets Arial lacks in the standard PDF fonts (Greek, Hebrew, Cyrillic): PDF export draws each
// in the Noto font for it, not the second and third as "?".
build('mixed-scripts', [
  para(run('Greek αβγδ, Hebrew שלום and Cyrillic Жизнь in one run.', '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>')),
  SECT(),
].join(''));
