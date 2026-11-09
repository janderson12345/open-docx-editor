// A second set of focused .docx files for comparing our layout with Word's (features the first set doesn't isolate).
// Each file isolates one layout feature. Render them to PDF in Word (see
// docs/word-reference.md) and compare page by page.
// Run: node test/fixtures/make-layout-corpus2.mjs <out-dir>
import fs from 'node:fs';
import path from 'node:path';
import { zipSync, strToU8 } from 'fflate';
import zlib from 'node:zlib';

const out = process.argv[2] || 'test/fixtures/layout2';
fs.mkdirSync(out, { recursive: true });

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const LOREM = 'The quick brown fox jumps over the lazy dog while seven stored procedures were rewritten to remove a predicate that forced a full scan of an eleven million row table, and the reduction was measured over a five minute window.';

const rpr = ({ font, size, b, i, color } = {}) => {
  const k = [];
  if (font) k.push(`<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}" w:eastAsia="${font}"/>`);
  if (b) k.push('<w:b/>');
  if (i) k.push('<w:i/>');
  if (color) k.push(`<w:color w:val="${color}"/>`);
  if (size) k.push(`<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>`);
  return k.length ? `<w:rPr>${k.join('')}</w:rPr>` : '';
};
const run = (text, o) => `<w:r>${rpr(o)}${text.split('\t').map((t, n) => (n ? '<w:tab/>' : '') + (t ? `<w:t xml:space="preserve">${esc(t)}</w:t>` : '')).join('')}</w:r>`;
const para = (content, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${typeof content === 'string' ? run(content) : content.join('')}</w:p>`;
const P = (text, ppr, o) => para([run(text, o)], ppr);
const sp = (before, after, line, rule = 'auto') => `<w:spacing w:before="${before}" w:after="${after}"${line != null ? ` w:line="${line}" w:lineRule="${rule}"` : ''}/>`;
const H = (text) => P(text, '<w:pStyle w:val="Heading1"/>');

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>`;
const NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${W}>
<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${[0, 1, 2].map((l) => `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${['', 'o', ''][l]}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (l + 1)}" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="${['Symbol', 'Courier New', 'Wingdings'][l]}" w:hAnsi="${['Symbol', 'Courier New', 'Wingdings'][l]}" w:hint="default"/></w:rPr></w:lvl>`).join('')}</w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${[['decimal', '%1.'], ['lowerLetter', '%2.'], ['lowerRoman', '%3.']].map(([f, t], l) => `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${f}"/><w:lvlText w:val="${t}"/><w:lvlJc w:val="${l === 2 ? 'right' : 'left'}"/><w:pPr><w:ind w:left="${720 * (l + 1)}" w:hanging="${l === 2 ? 180 : 360}"/></w:pPr></w:lvl>`).join('')}</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;
const SETTINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings ${W}><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;

const SECT = (extra = '') => `<w:sectPr>${extra.replace(/<w:cols\b[^>]*\/>/, '')}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>${(/<w:cols\b[^>]*\/>/.exec(extra) || ['<w:cols w:space="720"/>'])[0]}</w:sectPr>`;

function docx(name, body, { headers = {}, sect = '', parts = [] } = {}) {
  const rels = ['<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
    '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'];
  const files = {};
  const overrides = [];
  let refs = '';
  for (const [kind, [type, xml]] of Object.entries(headers)) {
    const tag = kind.startsWith('header') ? 'hdr' : 'ftr';
    const part = `${kind}.xml`;
    files[`word/${part}`] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${tag} ${W}>${xml}</w:${tag}>`);
    rels.push(`<Relationship Id="rId_${kind}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${tag === 'hdr' ? 'header' : 'footer'}" Target="${part}"/>`);
    overrides.push(`<Override PartName="/word/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${tag === 'hdr' ? 'header' : 'footer'}+xml"/>`);
    refs += `<w:${tag === 'hdr' ? 'headerReference' : 'footerReference'} w:type="${type}" r:id="rId_${kind}"/>`;
  }
  // Extra parts: [path under word/, relationship type (last segment), content type, xml or bytes, rId].
  for (const [part, type, ctype, data, id] of parts) {
    files[`word/${part}`] = typeof data === 'string' ? strToU8(data) : data;
    rels.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${part}"/>`);
    if (ctype) overrides.push(`<Override PartName="/word/${part}" ContentType="${ctype}"/>`);
  }
  Object.assign(files, {
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>${overrides.join('')}</Types>`),
    '_rels/.rels': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/_rels/document.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`),
    'word/styles.xml': strToU8(STYLES),
    'word/numbering.xml': strToU8(NUMBERING),
    'word/settings.xml': strToU8(SETTINGS),
    'word/document.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body.join('')}${SECT(refs + sect)}</w:body></w:document>`),
  });
  fs.writeFileSync(path.join(out, `${name}.docx`), zipSync(files));
}

const LONG = Array.from({ length: 4 }, () => LOREM).join(' ');
const rule = (side, sz, space, val = 'single', color = 'auto') => `<w:${side} w:val="${val}" w:sz="${sz}" w:space="${space}" w:color="${color}"/>`;
const tc = (text, { w = 2340, span, vmerge, shd, ppr = '', o } = {}) =>
  `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${span ? `<w:gridSpan w:val="${span}"/>` : ''}${vmerge ? `<w:vMerge${vmerge === 'restart' ? ' w:val="restart"' : ''}/>` : ''}${shd ? `<w:shd w:val="clear" w:color="auto" w:fill="${shd}"/>` : ''}</w:tcPr>${P(text, ppr, o)}</w:tc>`;
const tr = (cells, trpr = '') => `<w:tr>${trpr ? `<w:trPr>${trpr}</w:trPr>` : ''}${cells.join('')}</w:tr>`;
const tbl = (rows, { grid = [2340, 2340, 2340, 2340], style = 'TableGrid', extra = '' } = {}) =>
  `<w:tbl><w:tblPr><w:tblStyle w:val="${style}"/><w:tblW w:w="0" w:type="auto"/>${extra}</w:tblPr><w:tblGrid>${grid.map((g) => `<w:gridCol w:w="${g}"/>`).join('')}</w:tblGrid>${rows.join('')}</w:tbl>`;

// 1. Fonts at several sizes and line spacings: single, 1.15, 1.5, double, exact 14pt, at least 18pt.
const F2 = ['Calibri', 'Arial', 'Times New Roman', 'Cambria', 'Georgia', 'Verdana', 'Segoe UI', 'Aptos'];
const SPACINGS = [[240, 'auto'], [276, 'auto'], [360, 'auto'], [480, 'auto'], [280, 'exact'], [360, 'atLeast']];
docx('c2-01-fonts-spacing', F2.flatMap((f) => [18, 21, 24, 28].flatMap((sz) => [
  P(`${f} ${sz / 2}pt`, sp(120, 60, 240), { b: true, font: f, size: sz }),
  ...SPACINGS.slice(0, 3).map(([l, r]) => P(LOREM, sp(0, 80, l, r), { font: f, size: sz })),
])).concat(F2.flatMap((f) => SPACINGS.slice(3).map(([l, r]) => P(LONG, sp(0, 120, l, r), { font: f })))));

// 2. Justified text in several fonts and sizes (Word squeezes spaces to fit a word; the browser only stretches).
docx('c2-02-justify', F2.flatMap((f) => [20, 22, 24].map((sz) => P(LONG, `<w:jc w:val="both"/>${sp(0, 160, 259)}`, { font: f, size: sz }))));

// 3. Paragraph borders and shading: boxes of several widths and spaces, a group sharing borders, bottom rules.
docx('c2-03-borders', [
  ...[2, 4, 6, 8, 12, 18, 24].map((sz) => P(`Box, sz ${sz}, space 4. ${LOREM}`, `<w:pBdr>${['top', 'left', 'bottom', 'right'].map((s) => rule(s, sz, 4)).join('')}</w:pBdr>${sp(0, 200)}`)),
  ...[0, 1, 6, 12].map((space) => P(`Bottom rule sz 8, space ${space}. ${LOREM}`, `<w:pBdr>${rule('bottom', 8, space)}</w:pBdr>${sp(0, 200)}`)),
  ...[1, 2, 3].map((k) => P(`Grouped paragraph ${k}. ${LOREM}`, `<w:pBdr>${['top', 'left', 'bottom', 'right'].map((s) => rule(s, 6, 1)).join('')}</w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="DDEEFF"/>${sp(0, 0)}`)),
  P('After the group.', sp(200, 200)),
  ...['double', 'dotted', 'dashed', 'thickThinSmallGap', 'thinThickSmallGap'].map((v) => P(`${v} rule under. ${LOREM}`, `<w:pBdr>${rule('bottom', 12, 2, v)}</w:pBdr>${sp(0, 240)}`)),
]);

// 4. Tables: merged cells, row heights (at least / exact), cell margins, a long table with a repeated header row.
const head = tr(['Item', 'Description', 'Qty', 'Notes'].map((t) => tc(t, { shd: 'D9D9D9', o: { b: true } })), '<w:tblHeader/>');
docx('c2-04-tables', [
  H('Merged cells'),
  tbl([
    tr([tc('Spans two columns', { w: 4680, span: 2 }), tc('C'), tc('D')]),
    tr([tc('Merged down', { vmerge: 'restart' }), tc(LOREM), tc('x'), tc('y')]),
    tr([tc('', { vmerge: 'cont' }), tc('Second row of the merge'), tc('x'), tc('y')]),
    tr([tc('Spans all four', { w: 9360, span: 4 })]),
  ]),
  P(''),
  H('Row heights'),
  tbl([
    tr([tc('At least 0.5in'), tc('short'), tc('x'), tc('y')], '<w:trHeight w:val="720"/>'),
    tr([tc('Exact 0.3in'), tc(LOREM), tc('x'), tc('y')], '<w:trHeight w:val="432" w:hRule="exact"/>'),
    tr([tc('Auto'), tc(LOREM), tc('x'), tc('y')]),
  ]),
  P(''),
  H('Cell margins'),
  tbl([tr([tc(LOREM), tc(LOREM), tc('x'), tc('y')]), tr([tc('a'), tc('b'), tc('c'), tc('d')])], { extra: '<w:tblCellMar><w:top w:w="115" w:type="dxa"/><w:left w:w="180" w:type="dxa"/><w:bottom w:w="115" w:type="dxa"/><w:right w:w="180" w:type="dxa"/></w:tblCellMar>' }),
  P(''),
  H('A long table with a repeated header row'),
  tbl([head, ...Array.from({ length: 60 }, (_, k) => tr([tc(String(k + 1)), tc(k % 3 ? 'A short description' : LOREM), tc(String((k * 7) % 13)), tc(k % 5 ? '' : 'Check')]))]),
]);

// 5. Keeping together near page ends: headings kept with the next paragraph, kept lines, a page break before.
docx('c2-05-keep', Array.from({ length: 14 }, (_, k) => [
  H(`Section ${k + 1}`),
  P(LONG, sp(0, 160, 259)),
  P(k % 2 ? LOREM : LONG, `${k % 3 === 0 ? '<w:keepLines/>' : ''}${sp(0, 160, 259)}`),
  ...(k === 9 ? [P('This paragraph starts a page.', '<w:pageBreakBefore/>')] : []),
]).flat());

// 6. Multilevel numbered and bulleted lists running over pages.
docx('c2-06-lists', Array.from({ length: 40 }, (_, k) => {
  const l = [0, 1, 2, 1, 0, 1][k % 6];
  const num = k % 12 < 8 ? 2 : 1;
  return P(k % 4 === 3 ? LOREM : `List item ${k + 1} at level ${l + 1}`, `<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${l}"/><w:numId w:val="${num}"/></w:numPr>`);
}));

// 7. Section breaks: a next-page section with narrower margins, a landscape one, then two columns.
const sectBreak = (inner) => `<w:p><w:pPr><w:sectPr>${inner}</w:sectPr></w:pPr></w:p>`;
const PORTRAIT = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
docx('c2-07-sections', [
  H('Portrait, 1in margins'), P(LONG), P(LONG),
  sectBreak(PORTRAIT),
  H('Narrow margins'), P(LONG), P(LONG),
  sectBreak('<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/>'),
  H('Landscape'), P(LONG), P(LONG),
  sectBreak('<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'),
  H('Back to portrait, two columns from here'), P(LOREM),
  sectBreak(PORTRAIT),
  P(LONG), P(LONG), P(LONG), P(LONG),
], { sect: '<w:type w:val="continuous"/><w:cols w:num="2" w:space="720"/>' });

// 8. Tabs: right tabs with dot leaders (a contents page), centre and decimal tabs, other leaders.
const tabs = (list) => `<w:tabs>${list.map(([val, pos, leader]) => `<w:tab w:val="${val}" w:pos="${pos}"${leader ? ` w:leader="${leader}"` : ''}/>`).join('')}</w:tabs>`;
docx('c2-08-tabs', [
  H('Contents'),
  ...Array.from({ length: 25 }, (_, k) => P(`${k + 1}. Chapter title number ${k + 1}${k % 3 ? '' : ' with a longer name'}\t${3 + k * 4}`, tabs([['right', 9350, 'dot']]))),
  H('Centre and decimal tabs'),
  ...['1.5', '12.25', '123.125', '0.5'].map((n) => P(`Left\tCentred text\t${n}`, tabs([['center', 4680], ['decimal', 8000]]))),
  ...['hyphen', 'underscore', 'middleDot'].map((l) => P(`Leader ${l}\t99`, tabs([['right', 9350, l]]))),
]);

// 9. Spacing: before and after together, contextual spacing in a list, widely spaced paragraphs over pages.
docx('c2-09-spacing', [
  ...[[0, 0], [240, 0], [0, 240], [240, 240], [480, 120], [120, 480]].flatMap(([b, a]) => [P(`before ${b} after ${a}. ${LOREM}`, sp(b, a, 259)), P(LOREM, sp(b, a, 259))]),
  ...Array.from({ length: 6 }, (_, k) => P(`Contextual item ${k + 1}`, `<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>${sp(120, 120)}`)),
  ...Array.from({ length: 14 }, () => P(LONG, sp(360, 360, 300))),
]);

// 10. Runs on a line: mixed sizes, superscript, bold and italic, character spacing.
docx('c2-10-runs', Array.from({ length: 12 }, (_, k) => para([
  run('Mixed ', {}), run('large ', { size: 32 }), run('and small ', { size: 16 }), run('bold ', { b: true }), run('italic ', { i: true }),
  '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>sup</w:t></w:r>', run(' text. '),
  `<w:r><w:rPr><w:spacing w:val="${(k % 4) * 10}"/></w:rPr><w:t xml:space="preserve">Spaced characters ${k} </w:t></w:r>`, run(LOREM),
], sp(0, 120, 259))));

// --- Corpus 3: footnotes, nested tables, inline pictures, autofit tables, hyperlinks and character formatting.
// A solid-colour PNG (w x h px), for inline pictures.
function png(w, h, [r, g, b]) {
  const crcT = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
  const crc = (buf) => { let c = 0xffffffff; for (const x of buf) c = crcT[(c ^ x) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const chunk = (type, data) => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type), data]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([len, td, c]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  const raw = Buffer.alloc((w * 3 + 1) * h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) raw.set([r, g, b], y * (w * 3 + 1) + 1 + x * 3);
  return new Uint8Array(Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]));
}
const pic = (rid, wpx, hpx, id) => `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${wpx * 9525}" cy="${hpx * 9525}"/><wp:docPr id="${id}" name="Picture ${id}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="p${id}.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${wpx * 9525}" cy="${hpx * 9525}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

// 11. Footnotes: references in body text, several per page, some long notes.
const fnote = (id, text) => `<w:footnote w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve"> ${esc(text)}</w:t></w:r></w:p></w:footnote>`;
const FOOTNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${W}><w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>${Array.from({ length: 12 }, (_, k) => fnote(k + 1, k % 4 === 3 ? LOREM : `Footnote number ${k + 1}.`)).join('')}</w:footnotes>`;
const fref = (id) => `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
docx('c3-01-footnotes', Array.from({ length: 12 }, (_, k) => para([run(`Paragraph ${k + 1}. ${LOREM}`), fref(k + 1), run(` ${LOREM}`)], sp(0, 160, 259))), {
  parts: [['footnotes.xml', 'footnotes', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml', FOOTNOTES, 'rIdFn']],
});

// 12. Nested tables, and a row taller than what is left of the page.
const inner = tbl([tr([tc('inner a', { w: 1100 }), tc('inner b', { w: 1100 })]), tr([tc(LOREM, { w: 1100 }), tc('x', { w: 1100 })])], { grid: [1100, 1100] });
docx('c3-02-nested-tables', [
  H('A table in a cell'),
  tbl([tr([tc('Outer one'), `<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr>${inner}${P('after the inner table')}</w:tc>`, tc('three'), tc('four')]), tr([tc('a'), tc('b'), tc('c'), tc('d')])]),
  P(''),
  H('A row taller than what is left of the page'),
  ...Array.from({ length: 8 }, () => P(LOREM)),
  tbl([tr([tc(Array.from({ length: 6 }, () => LOREM).join(' ')), tc(LONG), tc('x'), tc('y')]), tr([tc('next'), tc('row'), tc('x'), tc('y')])]),
]);

// 13. Inline pictures: inside lines of text, on a line of their own, in a table cell.
const IMGS = [[60, 40, [200, 40, 40]], [200, 120, [40, 120, 200]], [16, 16, [40, 160, 40]]];
docx('c3-03-pictures', [
  H('Pictures in the text'),
  para([run('A small picture '), pic('rIdImg2', 16, 16, 3), run(' sits in this line, which goes on. '), run(LOREM)], sp(0, 160, 259)),
  para([run('A taller one '), pic('rIdImg0', 60, 40, 1), run(' makes its line taller. '), run(LOREM)], sp(0, 160, 259)),
  para([pic('rIdImg1', 200, 120, 2)], sp(0, 160, 259)),
  P(LONG, sp(0, 160, 259)),
  tbl([tr([tc('In a cell:'), `<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr>${para([pic('rIdImg0', 60, 40, 4)])}</w:tc>`, tc('x'), tc('y')])]),
  ...Array.from({ length: 3 }, (_, k) => para([run(LOREM), pic('rIdImg0', 60, 40, 10 + k), run(LOREM)], sp(0, 160, 259))),
], { parts: IMGS.map(([w, h, c], k) => [`media/img${k}.png`, 'image', null, png(w, h, c), `rIdImg${k}`]) });

// 14. Autofit tables (widths from their content) and percent widths.
const atc = (text) => `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>${P(text)}</w:tc>`;
docx('c3-04-autofit', [
  H('Autofit to contents'),
  `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="autofit"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>${[['Name', 'A longer description of the item that should wrap if it has to', '12'], ['Short', 'x', '1234567'], ['Medium name', LOREM, '3']].map((r) => tr(r.map(atc))).join('')}</w:tbl>`,
  P(''),
  H('Percent widths'),
  `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="3120"/><w:gridCol w:w="6240"/></w:tblGrid>${tr([`<w:tc><w:tcPr><w:tcW w:w="1667" w:type="pct"/></w:tcPr>${P('One third')}</w:tc>`, `<w:tc><w:tcPr><w:tcW w:w="3333" w:type="pct"/></w:tcPr>${P(LOREM)}</w:tc>`])}</w:tbl>`,
]);

// 15. Character formatting across lines: underline, caps, small caps, highlight, strike, colour, a hyperlink.
const rr = (text, k) => `<w:r><w:rPr>${k}</w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
docx('c3-05-character', Array.from({ length: 10 }, (_, k) => para([
  rr('Underlined words, ', '<w:u w:val="single"/>'), rr('ALL CAPS STYLED, ', '<w:caps/>'), rr('Small Caps Here, ', '<w:smallCaps/>'),
  rr('highlighted, ', '<w:highlight w:val="yellow"/>'), rr('struck, ', '<w:strike/>'), rr('coloured ', '<w:color w:val="C00000"/>'),
  '<w:hyperlink w:anchor="top"><w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>a link</w:t></w:r></w:hyperlink>', run(`. ${LOREM} ${k}`),
], sp(0, 160, 259))));
// 16. Two columns in a middle section: continuous breaks in and out, so Word balances the columns and the text after
// them starts under the longer one.
const COLS = (num, type = 'continuous') => `<w:type w:val="${type}"/>${PORTRAIT}<w:cols w:num="${num}" w:space="720"/>`;
docx('c3-06-mid-columns', [
  H('One column first'), P(LOREM, sp(0, 160, 259)),
  sectBreak(`<w:type w:val="continuous"/>${PORTRAIT}<w:cols w:space="720"/>`),
  ...Array.from({ length: 5 }, (_, k) => P(k % 2 ? LONG : LOREM, sp(0, 160, 259))),
  sectBreak(COLS(2)),
  H('One column again'), P(LONG, sp(0, 160, 259)),
  sectBreak(`<w:type w:val="continuous"/>${PORTRAIT}<w:cols w:space="720"/>`),
  ...Array.from({ length: 9 }, (_, k) => P(k % 3 ? LOREM : LONG, sp(0, 160, 259))),
  sectBreak(COLS(2)),
  P('The end, in one column.', sp(0, 160, 259)),
], { sect: '<w:type w:val="continuous"/>' });

// 17. Three columns ending with a next-page break, then one column on the new page.
docx('c3-07-three-columns', [
  H('Three columns, then a new page'),
  sectBreak(`<w:type w:val="continuous"/>${PORTRAIT}<w:cols w:space="720"/>`),
  ...Array.from({ length: 12 }, (_, k) => P(k % 2 ? LOREM : LONG, sp(0, 160, 259))),
  sectBreak(COLS(3)),
  H('A new page in one column'), P(LONG, sp(0, 160, 259)), P(LOREM, sp(0, 160, 259)),
], { sect: '<w:type w:val="nextPage"/>' });

// 18. Tables inside columns: a column-wide table in a middle section and in the last one.
const ctbl = (n) => tbl([tr([tc('Name', { w: 2160 }), tc('Value', { w: 2160 })]), ...Array.from({ length: n }, (_, k) => tr([tc(`Row ${k + 1}`, { w: 2160 }), tc(k % 3 ? `${k * 17}` : 'A longer value that wraps in its cell', { w: 2160 })]))], { grid: [2160, 2160] });
docx('c3-08-column-tables', [
  H('Tables in columns'),
  sectBreak(`<w:type w:val="continuous"/>${PORTRAIT}<w:cols w:space="720"/>`),
  P(LOREM, sp(0, 160, 259)), ctbl(6), P(LONG, sp(0, 160, 259)), ctbl(4), P(LOREM, sp(0, 160, 259)),
  sectBreak(COLS(2)),
  H('One column between'), P(LOREM, sp(0, 160, 259)),
  sectBreak(`<w:type w:val="continuous"/>${PORTRAIT}<w:cols w:space="720"/>`),
  P(LONG, sp(0, 160, 259)), ctbl(8), P(LONG, sp(0, 160, 259)), P(LOREM, sp(0, 160, 259)), ctbl(5), P(LOREM, sp(0, 160, 259)),
], { sect: '<w:type w:val="continuous"/><w:cols w:num="2" w:space="720"/>' });
// 19. A table with a header row crossing into the next column: the header row is repeated there, as in Word.
const hrow = tr([tc('Item', { w: 2160, shd: 'D9D9D9', o: { b: true } }), tc('Amount', { w: 2160, shd: 'D9D9D9', o: { b: true } })], '<w:tblHeader/>');
docx('c3-09-column-header-rows', [
  H('A header row in columns'),
  sectBreak(`<w:type w:val="continuous"/>${PORTRAIT}<w:cols w:space="720"/>`),
  P(LONG, sp(0, 160, 259)), P(LOREM, sp(0, 160, 259)),
  tbl([hrow, ...Array.from({ length: 28 }, (_, k) => tr([tc(`Line ${k + 1}`, { w: 2160 }), tc(k % 4 ? `${(k + 3) * 41}` : 'A value long enough to wrap', { w: 2160 })]))], { grid: [2160, 2160] }),
  P(LOREM, sp(0, 160, 259)),
], { sect: '<w:type w:val="continuous"/><w:cols w:num="2" w:space="720"/>' });
console.log('wrote', fs.readdirSync(out).length, 'files to', out);
