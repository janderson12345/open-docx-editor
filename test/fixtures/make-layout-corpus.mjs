// Builds small, focused .docx files for comparing our page layout with Word's.
// Each file isolates one layout feature. Render them to PDF in Word (see
// docs/word-reference.md) and compare page by page.
// Run: node test/fixtures/make-layout-corpus.mjs <out-dir>
import fs from 'node:fs';
import path from 'node:path';
import { zipSync, strToU8 } from 'fflate';

const out = process.argv[2] || 'test/fixtures/layout';
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

const SECT = (extra = '') => `<w:sectPr>${extra}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>`;

function docx(name, body, { headers = {}, sect = '' } = {}) {
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
  Object.assign(files, {
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>${overrides.join('')}</Types>`),
    '_rels/.rels': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/_rels/document.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`),
    'word/styles.xml': strToU8(STYLES),
    'word/numbering.xml': strToU8(NUMBERING),
    'word/settings.xml': strToU8(SETTINGS),
    'word/document.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body.join('')}${SECT(refs + sect)}</w:body></w:document>`),
  });
  fs.writeFileSync(path.join(out, `${name}.docx`), zipSync(files));
}

// 1. Fonts: Word's single line height and line breaks differ per font.
const FONTS = ['Calibri', 'Calibri Light', 'Cambria', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Tahoma', 'Segoe UI', 'Aptos', 'Consolas', 'Courier New'];
docx('01-fonts', FONTS.flatMap((f) => [P(f, sp(120, 0, 240), { b: true, font: f }), P(LOREM, sp(0, 0, 240), { font: f }), P(LOREM, sp(0, 120, 240), { font: f, size: 18 })]));

// 2. Line spacing rules and space before/after between paragraphs.
docx('02-spacing', [
  ...[[240, 'auto'], [276, 'auto'], [360, 'auto'], [480, 'auto'], [240, 'exact'], [360, 'exact'], [120, 'atLeast'], [480, 'atLeast']].map(([l, r]) => P(`line=${l} ${r}. ${LOREM}`, sp(0, 0, l, r))),
  P('after=240 then before=240 (are they added or collapsed?)', sp(0, 240, 240)),
  P('second paragraph, before=240', sp(240, 0, 240)),
  P('after=120', sp(0, 120, 240)),
  P('before=360', sp(360, 0, 240)),
  P('Same-style paragraphs with contextual spacing:', sp(0, 0, 240)),
  ...[1, 2, 3].map((n) => P(`List paragraph ${n} (after 160 ignored between them)`, '<w:pStyle w:val="ListParagraph"/>')),
  P('Mixed sizes on one line:', sp(240, 0, 240)),
  para([run('small 8pt ', { size: 16 }), run('normal 11pt ', {}), run('big 20pt ', { size: 40 }), run('normal again, wrapping onto more lines: ' + LOREM, {})], sp(0, 0, 240)),
  P('Empty paragraphs at 11pt, 20pt, 8pt follow:', sp(240, 0, 240)),
  para([], sp(0, 0, 240)),
  para([], sp(0, 0, 240) + '<w:rPr><w:sz w:val="40"/></w:rPr>'),
  para([], sp(0, 0, 240) + '<w:rPr><w:sz w:val="16"/></w:rPr>'),
  P('end', sp(0, 0, 240)),
]);

// 3. Paragraphs crossing page breaks: widow/orphan control, keepLines, keepNext, pageBreakBefore.
const long = Array.from({ length: 6 }, () => LOREM).join(' ');
docx('03-paragraph-breaks', [
  ...Array.from({ length: 7 }, (_, n) => P(`Filler ${n + 1}. ${LOREM} ${LOREM}`)),
  P(`Widow control on (default). ${long}`),
  ...Array.from({ length: 4 }, (_, n) => P(`Filler ${n + 8}. ${LOREM} ${LOREM}`)),
  P(`Widow control off. ${long}`, '<w:widowControl w:val="0"/>'),
  ...Array.from({ length: 3 }, (_, n) => P(`Filler ${n + 12}. ${LOREM}`)),
  P(`Keep lines together. ${long}`, '<w:keepLines/>'),
  H('Heading keeps with next'),
  P(`Body after heading. ${LOREM}`),
  P('Page break before this paragraph.', '<w:pageBreakBefore/>'),
  para([run('Run-level page break follows'), '<w:r><w:br w:type="page"/></w:r>', run('after the break')]),
]);

// 4. Tables: border widths, cell margins, splitting with repeated header, cantSplit, vertical merge, alignment.
const tc = (text, { w = 2340, extra = '', o } = {}) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${extra}</w:tcPr>${(Array.isArray(text) ? text : [text]).map((t) => P(t, '', o)).join('')}</w:tc>`;
const tr = (cells, trPr = '') => `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells.join('')}</w:tr>`;
const tbl = (rows, tblPr = '<w:tblStyle w:val="TableGrid"/>', cols = [2340, 2340, 2340, 2340]) => `<w:tbl><w:tblPr>${tblPr}<w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${cols.map((c) => `<w:gridCol w:w="${c}"/>`).join('')}</w:tblGrid>${rows.join('')}</w:tbl>`;
const bordered = (sz) => `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="${sz}" w:space="0" w:color="000000"/>`).join('')}</w:tblBorders>`;
docx('04-tables', [
  P('Table Grid (0.5pt), default cell margins:'),
  tbl([tr(['A1', 'B1', 'C1', 'D1'].map((t) => tc(t))), tr(['A2 with more text that wraps inside the cell', 'B2', 'C2', 'D2'].map((t) => tc(t)))]),
  ...[2, 12, 24].map((sz) => [P(`Borders sz=${sz} (${sz / 8}pt):`, sp(240, 0, 240)), tbl([1, 2, 3].map((r) => tr([1, 2, 3, 4].map((c) => tc(`r${r}c${c}`)))), bordered(sz))]).flat(),
  P('No borders, cell margins 0.1in all round, shading, vertical alignment:', sp(240, 0, 240)),
  tbl([tr([tc('top'), tc('center', { extra: '<w:vAlign w:val="center"/>' }), tc('bottom', { extra: '<w:vAlign w:val="bottom"/>' }), tc(['tall', 'cell', 'with', 'lines'], { extra: '<w:shd w:val="clear" w:color="auto" w:fill="DDEEFF"/>' })])],
    '<w:tblCellMar><w:top w:w="144" w:type="dxa"/><w:left w:w="144" w:type="dxa"/><w:bottom w:w="144" w:type="dxa"/><w:right w:w="144" w:type="dxa"/></w:tblCellMar>'),
  P('Vertically merged cell:', sp(240, 0, 240)),
  tbl([tr([tc('merged', { extra: '<w:vMerge w:val="restart"/>' }), tc('b1'), tc('c1'), tc('d1')]), tr([tc('', { extra: '<w:vMerge/>' }), tc('b2'), tc('c2'), tc('d2')])]),
  P('A 70-row table with a repeating header row; row 30 cannot split:', sp(240, 0, 240)),
  tbl([tr(['#', 'Name', 'Value', 'Notes'].map((t) => tc(t, { o: { b: true } })), '<w:tblHeader/>'),
    ...Array.from({ length: 70 }, (_, n) => tr([tc(String(n + 1)), tc(`Item ${n + 1}`), tc(String((n + 1) * 137)), tc(n % 7 === 3 ? ['two', 'lines'] : n === 29 ? ['cant', 'split', 'this', 'row', 'of', 'six'] : 'ok')], n === 29 ? '<w:cantSplit/>' : ''))]),
  P('After the table.'),
]);

// 5. Lists: bullets, numbers, nesting, restart.
const li = (text, numId, ilvl) => P(text, `<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`);
docx('05-lists', [
  P('Bullets:'), li('First bullet', 1, 0), li('Second bullet, long enough to wrap onto a second line so the hanging indent shows. ' + LOREM, 1, 0), li('Nested bullet', 1, 1), li('Third level', 1, 2), li('Back to top', 1, 0),
  P('Numbers:'), li('One', 2, 0), li('Two', 2, 0), li('Two a', 2, 1), li('Two b ' + LOREM, 2, 1), li('Two b i', 2, 2), li('Two b ii', 2, 2), li('Three', 2, 0),
  ...Array.from({ length: 9 }, (_, n) => li(`Item ${n + 4} (two-digit numbers shift the tab)`, 2, 0)),
]);

// 6. Tabs: custom stops, alignment, leaders, default stops.
docx('06-tabs', [
  P('Default stops:\tone\ttwo\tthree'),
  P('Left 2in\tRight 6in', '<w:tabs><w:tab w:val="left" w:pos="2880"/><w:tab w:val="right" w:pos="8640"/></w:tabs>'),
  P('Chapter one\t12', '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs>'),
  P('Center\tcentered at 3.25in\tend', '<w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs>'),
  ...['1.5', '123.25', '12345.125'].map((v) => P(`Decimal\t${v}`, '<w:tabs><w:tab w:val="decimal" w:pos="4320"/></w:tabs>')),
  P('Hanging indent\twith a tab after the hang. ' + LOREM, '<w:ind w:left="1440" w:hanging="1440"/>'),
]);

// 7. Indents, alignment, borders, shading.
docx('07-paragraph-format', [
  P('Left aligned. ' + LOREM),
  P('Centered. ' + LOREM, '<w:jc w:val="center"/>'),
  P('Right aligned. ' + LOREM, '<w:jc w:val="right"/>'),
  P('Justified. ' + LOREM + ' ' + LOREM, '<w:jc w:val="both"/>'),
  P('Indented 1in both sides, first line 0.5in. ' + LOREM, '<w:ind w:left="1440" w:right="1440" w:firstLine="720"/>'),
  P('Hanging 0.5in. ' + LOREM, '<w:ind w:left="720" w:hanging="720"/>'),
  P('Box border, 4pt space, shaded. ' + LOREM, '<w:pBdr><w:top w:val="single" w:sz="8" w:space="4" w:color="1F4E79"/><w:left w:val="single" w:sz="8" w:space="4" w:color="1F4E79"/><w:bottom w:val="single" w:sz="8" w:space="4" w:color="1F4E79"/><w:right w:val="single" w:sz="8" w:space="4" w:color="1F4E79"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="EAF1FB"/>'),
  P('Bottom rule only, like a heading underline.', '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="2E74B5"/></w:pBdr>', { size: 32 }),
  P('Left bar callout. ' + LOREM, '<w:pBdr><w:left w:val="single" w:sz="24" w:space="8" w:color="C55A11"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="FBE5D6"/><w:ind w:left="200"/>'),
  P('Two bordered paragraphs in a row share one box:', sp(240, 0, 240)),
  ...[1, 2].map((n) => P(`Boxed ${n}. ${LOREM}`, '<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="4" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="4" w:color="auto"/></w:pBdr>')),
]);

// 8. Headers and footers: page X of Y, different first page.
const fld = (instr) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
docx('08-headers-footers', [
  ...Array.from({ length: 24 }, (_, n) => P(`Paragraph ${n + 1}. ${LOREM} ${LOREM}`)),
], {
  sect: '<w:titlePg/>',
  headers: {
    header1: ['default', P('Default header, right aligned', '<w:jc w:val="right"/>')],
    header2: ['first', P('FIRST PAGE HEADER', '<w:jc w:val="center"/>', { b: true })],
    footer1: ['default', para([run('Page '), fld('PAGE'), run(' of '), fld('NUMPAGES')], '<w:jc w:val="center"/>')],
  },
});

const list = fs.readdirSync(out).filter((f) => f.endsWith('.docx'));
console.log(`${list.length} files in ${out}: ${list.join(', ')}`);
