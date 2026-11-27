// Builds the test/fixtures/gaps-*.docx documents: gaps between body paragraphs as Word lays them out (its prints of the
// same documents, made here; checks in test/features.mjs). Each holds plain 12pt Times New Roman paragraphs round:
//   gaps-theme-bidi: runs in the theme's majorBidi font (Word: Times New Roman, whatever the theme's a:cs);
//   gaps-deleted-neighbour: a heading-like paragraph with contextual spacing, then an empty one of its style deleted whole
//     (tracked): the heading keeps its 12pt after over the next paragraph;
//   gaps-bullet-missing-font: 14pt bullets in Noto Sans Symbols, a font this machine lacks (Word: Calibri's line);
//   gaps-closing-tabs: a line break then only tabs, and 18 tabs after text that run over: each a line;
//   gaps-contextual-before: contextual spacing on the first of two paragraphs, 6pt after it, 6pt and 12pt before the next;
//   gaps-border-groups: bottom-bordered paragraphs, the second's indents its own, joined where the borders line up;
//   gaps-trailing-space-exact: a centred title on an exact line, a break, text ending in a space;
//   gaps-band-below-lines: a picture wrapped top and bottom 0.7in below a one-line paragraph: the next paragraph's first
//     line above it, its second under it.
// usage: node test/fixtures/mkgaps.mjs [dir]
import '../setup.js';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync, strToU8 } from 'fflate';

const out = path.resolve(process.argv[2] || path.dirname(fileURLToPath(import.meta.url)));
fs.mkdirSync(out, { recursive: true });
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DATE = 'w:author="Tester" w:date="2026-01-01T00:00:00Z"';

const theme = (cs) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements>`
  + '<a:clrScheme name="T"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="0E2841"/></a:dk2><a:lt2><a:srgbClr val="E8E8E8"/></a:lt2>'
  + '<a:accent1><a:srgbClr val="156082"/></a:accent1><a:accent2><a:srgbClr val="E97132"/></a:accent2><a:accent3><a:srgbClr val="196B24"/></a:accent3><a:accent4><a:srgbClr val="0F9ED5"/></a:accent4>'
  + '<a:accent5><a:srgbClr val="A02B93"/></a:accent5><a:accent6><a:srgbClr val="4EA72E"/></a:accent6><a:hlink><a:srgbClr val="467886"/></a:hlink><a:folHlink><a:srgbClr val="96607D"/></a:folHlink></a:clrScheme>'
  + `<a:fontScheme name="T"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface="${cs}"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface="${cs}"/></a:minorFont></a:fontScheme>`
  + '<a:fmtScheme name="T"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
  + '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
  + '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
  + '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>';

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>`
  + '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
  + '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Step"><w:name w:val="Step"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="0" w:after="240"/><w:contextualSpacing/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Spaced"><w:name w:val="Spaced"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/><w:contextualSpacing/></w:pPr></w:style>'
  + '<w:style w:type="paragraph" w:customStyle="1" w:styleId="BidiStyle"><w:name w:val="Bidi Style"/><w:basedOn w:val="Normal"/><w:rPr><w:rFonts w:asciiTheme="majorBidi" w:hAnsiTheme="majorBidi" w:cstheme="majorBidi"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Ruled"><w:name w:val="Ruled"/><w:basedOn w:val="Normal"/><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="18" w:color="808080"/></w:pBdr><w:ind w:left="1080" w:hanging="1080"/></w:pPr></w:style>'
  + '</w:styles>';

const numbering = (font, sz) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>`
  + `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="●"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="${font}" w:eastAsia="${font}" w:hAnsi="${font}" w:cs="${font}"/>${sz ? `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>` : ''}</w:rPr></w:lvl>`
  + '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';

const fontTable = (alt) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="${W}"><w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>`
  + (alt === 'bare' ? '<w:font w:name="Noto Sans Symbols"><w:charset w:val="00"/><w:family w:val="auto"/><w:pitch w:val="default"/></w:font>'
    : alt ? `<w:font w:name="Noto Sans Symbols">${alt === '-' ? '' : `<w:altName w:val="${alt}"/>`}<w:panose1 w:val="020B0604020202020204"/><w:charset w:val="00"/><w:family w:val="auto"/><w:pitch w:val="default"/></w:font>` : '') + '</w:fonts>';

const settings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${W}"><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;

function save(name, body, { cs = '', num = null, alt = null, base = 'Times New Roman', mode = 15, pic = false } = {}) {
  const files = {};
  if (pic) files['word/media/p.png'] = PNG;
  const parts = [['styles', 'styles.xml', 'styles+xml'], ['theme', 'theme/theme1.xml', null], ['settings', 'settings.xml', 'settings+xml'], ['fontTable', 'fontTable.xml', 'fontTable+xml']];
  if (num) parts.push(['numbering', 'numbering.xml', 'numbering+xml']);
  files['[Content_Types].xml'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    + '<Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + parts.map(([, p, t]) => `<Override PartName="/word/${p}" ContentType="${t ? `application/vnd.openxmlformats-officedocument.wordprocessingml.${t}` : 'application/vnd.openxmlformats-officedocument.theme+xml'}"/>`).join('') + '</Types>');
  files['_rels/.rels'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  files['word/_rels/document.xml.rels'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PR}">${parts.map(([k, p], i) => `<Relationship Id="rId${i + 1}" Type="${R}/${k}" Target="${p}"/>`).join('')}${pic ? `<Relationship Id="rIdP" Type="${R}/image" Target="media/p.png"/>` : ''}</Relationships>`);
  files['word/styles.xml'] = strToU8(STYLES.replace(/Times New Roman/g, base));
  files['word/theme/theme1.xml'] = strToU8(theme(cs));
  files['word/settings.xml'] = strToU8(mode ? settings.replace('w:val="15"', `w:val="${mode}"`) : settings.replace(/<w:compat>.*<\/w:compat>/, ''));
  files['word/fontTable.xml'] = strToU8(fontTable(alt));
  if (num) files['word/numbering.xml'] = strToU8(numbering(...num));
  files['word/document.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${body.join('')}`
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr></w:body></w:document>');
  fs.writeFileSync(path.join(out, `${name}.docx`), zipSync(files));
}

const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAEUlEQVR4nGNgYPj/n4GBgQEADvcB/0V0Pw0AAAAASUVORK5CYII=', 'base64'));
const LONG = 'the quick brown fox jumps over the lazy dog and keeps on running across the wide green field until the line wraps';
const r = (t, rpr = '') => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const p = (runs, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${runs}</w:p>`;
const plain = (n) => p(r(`Plain line ${n}: ${LONG}.`));
const tf = (k) => `<w:rFonts w:asciiTheme="${k}" w:hAnsiTheme="${k}" w:cstheme="${k}"/>`;

const tabs = (n, rpr = '') => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}${'<w:tab/>'.repeat(n)}</w:r>`;
const br = '<w:r><w:br/></w:r>';
const item = (n) => p(r(`Bullet item ${n}: ${LONG}.`), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="720" w:hanging="360"/>');
const delMark = `<w:rPr><w:del w:id="9" ${DATE}/></w:rPr>`;
const ctxA = '<w:spacing w:before="120" w:after="120"/><w:contextualSpacing/>';
const ruled = (t, ind = '') => p(r(t), `<w:pStyle w:val="Ruled"/>${ind}`);

save('gaps-theme-bidi', [plain(1), ...[2, 3].map((n) => p(r(`Themed line ${n}: ${LONG}.`, tf('majorBidi')))), plain(4)]);
save('gaps-deleted-neighbour', [plain(1), p(r('Step heading one'), '<w:pStyle w:val="Step"/>'), p('', `<w:pStyle w:val="Step"/>${delMark}`), plain(2), plain(3)]);
save('gaps-bullet-missing-font', [plain(1), item(1), item(2), plain(2)], { num: ['Noto Sans Symbols', 28], alt: 'Calibri' });
save('gaps-closing-tabs', [plain(1), p(r('Text before the break') + br + tabs(3)), p(r(`Text and then tabs to the end: ${LONG}`) + tabs(18)), plain(2)]);
save('gaps-contextual-before', [plain(1), p(r('First of the pair'), ctxA), p(r('Second, 6pt before'), '<w:spacing w:before="120"/>'), p(r('Third of them'), ctxA), p(r('Fourth, 12pt before'), '<w:spacing w:before="240"/>'), plain(2)]);
save('gaps-border-groups', [plain(1), ruled('First ruled paragraph'), ruled('Second ruled, own left indent', '<w:ind w:left="1440" w:hanging="1440"/>'), plain(2), ruled('Third ruled paragraph'), ruled('Fourth ruled, own hanging', '<w:ind w:left="1080" w:hanging="360"/>'), plain(3)]);
save('gaps-trailing-space-exact', [plain(1), p(br + r('Centred title text ', '<w:b/><w:sz w:val="28"/>'), '<w:spacing w:before="80" w:after="160" w:line="280" w:lineRule="exact"/><w:jc w:val="center"/>'), plain(2)]);
const tb = (yEmu) => `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>78105</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${yEmu}</wp:posOffset></wp:positionV><wp:extent cx="5577840" cy="405130"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/><wp:docPr id="5" name="Picture 5"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="5" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdP"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5577840" cy="405130"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
save('gaps-band-below-lines', [plain(1), p(tb(635635) + r('Anchor paragraph, one line')), plain(2), plain(3)], { pic: true });
console.log('written to', out);
