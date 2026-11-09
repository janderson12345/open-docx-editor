// A minimal, valid blank .docx, generated in code (no Microsoft template used).
import { zipSync, strToU8 } from 'fflate';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const HDR = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

function style(type, id, name, { basedOn, next, rPr = '', pPr = '', def, priority, link } = {}) {
  return `<w:style w:type="${type}"${def ? ' w:default="1"' : ''} w:styleId="${id}"><w:name w:val="${name}"/>${basedOn ? `<w:basedOn w:val="${basedOn}"/>` : ''}${next ? `<w:next w:val="${next}"/>` : ''}${link ? `<w:link w:val="${link}"/>` : ''}${priority != null ? `<w:uiPriority w:val="${priority}"/>` : ''}<w:qFormat/>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}</w:style>`;
}

const heading = (n, sz, color = '2F5496') =>
  style('paragraph', `Heading${n}`, `heading ${n}`, { basedOn: 'Normal', next: 'Normal', priority: 9, pPr: `<w:keepNext/><w:keepLines/><w:spacing w:before="${n === 1 ? 360 : 160}" w:after="80"/><w:outlineLvl w:val="${n - 1}"/>`, rPr: `<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:color w:val="${color}"/><w:sz w:val="${sz}"/>` });

export const STYLES_XML = `${HDR}<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>${[
  style('paragraph', 'Normal', 'Normal', { def: true, priority: 0 }),
  style('paragraph', 'Title', 'Title', { basedOn: 'Normal', next: 'Normal', priority: 10, pPr: '<w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/>', rPr: '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/>' }),
  style('paragraph', 'Subtitle', 'Subtitle', { basedOn: 'Normal', next: 'Normal', priority: 11, rPr: '<w:color w:val="595959"/><w:spacing w:val="15"/><w:sz w:val="28"/>' }),
  heading(1, 32),
  heading(2, 26),
  heading(3, 24, '1F3763'),
  style('paragraph', 'Quote', 'Quote', { basedOn: 'Normal', next: 'Normal', priority: 29, pPr: '<w:spacing w:before="200"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/>', rPr: '<w:i/><w:color w:val="404040"/>' }),
  style('paragraph', 'ListParagraph', 'List Paragraph', { basedOn: 'Normal', priority: 34, pPr: '<w:ind w:left="720"/><w:contextualSpacing/>' }),
  style('character', 'Hyperlink', 'Hyperlink', { priority: 99, rPr: '<w:color w:val="0563C1"/><w:u w:val="single"/>' }),
  `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>`,
].join('')}</w:styles>`;

function lvl(i, fmt, text, font) {
  const left = 720 * (i + 1);
  return `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr>${font ? `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr>` : ''}</w:lvl>`;
}

const BULLETS = ['', 'o', ''];
const BULLET_FONTS = ['Symbol', 'Courier New', 'Wingdings'];
const NUM_FMTS = ['decimal', 'lowerLetter', 'lowerRoman'];

export function bulletAbstract(id) {
  return `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${Array.from({ length: 9 }, (_, i) => lvl(i, 'bullet', BULLETS[i % 3], BULLET_FONTS[i % 3])).join('')}</w:abstractNum>`;
}
export function decimalAbstract(id) {
  return `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${Array.from({ length: 9 }, (_, i) => lvl(i, NUM_FMTS[i % 3], `%${i + 1}.`)).join('')}</w:abstractNum>`;
}

export const NUMBERING_XML = `${HDR}<w:numbering ${W}>${bulletAbstract(0)}${decimalAbstract(1)}<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;

const SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>';

export function blankDocxBytes() {
  const files = {
    '[Content_Types].xml': `${HDR}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`,
    '_rels/.rels': `${HDR}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `${HDR}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`,
    'word/document.xml': `${HDR}<w:document ${W} ${R}><w:body><w:p/>${SECT}</w:body></w:document>`,
    'word/styles.xml': STYLES_XML,
    'word/numbering.xml': NUMBERING_XML,
  };
  const out = {};
  for (const [k, v] of Object.entries(files)) out[k] = strToU8(v);
  return zipSync(out);
}

/**
 * List definitions offered in the Bullets / Numbering / Multilevel galleries.
 * Each is 9 levels of [numFmt, lvlText, font?, left, hanging].
 */
const fill = (first, rest) => Array.from({ length: 9 }, (_, i) => (i === 0 ? [...first, 720, 360] : [...rest(i), 720 * (i + 1), 360]));
const bulletRest = (i) => ['bullet', BULLETS[i % 3], BULLET_FONTS[i % 3]];
const numberRest = (i) => [NUM_FMTS[i % 3], `%${i + 1}.`];
export const LIST_LIBRARY = {
  bullets: [
    { label: '•', levels: fill(['bullet', '', 'Symbol'], bulletRest) },
    { label: '◦', levels: fill(['bullet', 'o', 'Courier New'], bulletRest) },
    { label: '▪', levels: fill(['bullet', '', 'Wingdings'], bulletRest) },
    { label: '❖', levels: fill(['bullet', '', 'Wingdings'], bulletRest) },
    { label: '➢', levels: fill(['bullet', '', 'Wingdings'], bulletRest) },
    { label: '✓', levels: fill(['bullet', '', 'Wingdings'], bulletRest) },
    { label: '–', levels: fill(['bullet', '–', 'Calibri'], bulletRest) },
  ],
  numbering: [
    { label: '1.', levels: fill(['decimal', '%1.'], numberRest) },
    { label: '1)', levels: fill(['decimal', '%1)'], numberRest) },
    { label: 'I.', levels: fill(['upperRoman', '%1.'], numberRest) },
    { label: 'A.', levels: fill(['upperLetter', '%1.'], numberRest) },
    { label: 'a)', levels: fill(['lowerLetter', '%1)'], numberRest) },
    { label: 'a.', levels: fill(['lowerLetter', '%1.'], numberRest) },
    { label: 'i.', levels: fill(['lowerRoman', '%1.'], numberRest) },
  ],
  multilevel: [
    { label: '1. a. i.', levels: fill(['decimal', '%1.'], numberRest) },
    { label: '1. 1.1. 1.1.1.', levels: Array.from({ length: 9 }, (_, i) => ['decimal', Array.from({ length: i + 1 }, (_, j) => `%${j + 1}.`).join(''), null, 360 + 432 * i, 360 + 72 * i]) },
    { label: '• ◦ ▪', levels: fill(['bullet', '', 'Symbol'], bulletRest) },
    { label: 'I. A. 1. a)', levels: Array.from({ length: 9 }, (_, i) => [['upperRoman', 'upperLetter', 'decimal', 'lowerLetter', 'lowerRoman'][i % 5], `%${i + 1}${i % 5 === 3 ? ')' : '.'}`, null, 720 * (i + 1), 360]) },
  ],
};

export function abstractFor(id, levels) {
  return `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="${levels.every((l) => l[1] === levels[0][1]) ? 'singleLevel' : 'hybridMultilevel'}"/>${levels.map(([fmt, text, font, left, hanging], i) =>
    `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="${hanging}"/></w:pPr>${font ? `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr>` : ''}</w:lvl>`).join('')}</w:abstractNum>`;
}
