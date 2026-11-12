// The small OOXML toolkit the generator builds documents with: elements whose children are written in the schema's
// order (Word calls a file corrupt when they aren't), runs and paragraphs, a seeded random source, and the package
// (styles, numbering, settings, notes, comments, headers and footers, pictures) around a body.
import { zipSync, zlibSync, strToU8 } from 'fflate';

export const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
  'xmlns:w10="urn:schemas-microsoft-com:office:word"',
].join(' ');

export const esc = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

// ---- elements in schema order ----

const ORDER = {
  rPr: 'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange',
  pPr: 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange',
  sectPr: 'headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange',
  settings: 'writeProtection view zoom removePersonalInformation removeDateAndTime doNotDisplayPageBoundaries displayBackgroundShape printPostScriptOverText printFractionalCharacterWidth printFormsData embedTrueTypeFonts embedSystemFonts saveSubsetFonts saveFormsData mirrorMargins alignBordersAndEdges bordersDoNotSurroundHeader bordersDoNotSurroundFooter gutterAtTop hideSpellingErrors hideGrammaticalErrors activeWritingStyle proofState formsDesign attachedTemplate linkStyles stylePaneFormatFilter stylePaneSortMethod documentType mailMerge revisionView trackRevisions doNotTrackMoves doNotTrackFormatting documentProtection autoFormatOverride styleLockTheme styleLockQFSet defaultTabStop autoHyphenation consecutiveHyphenLimit hyphenationZone doNotHyphenateCaps showEnvelope summaryLength clickAndTypeStyle defaultTableStyle evenAndOddHeaders bookFoldRevPrinting bookFoldPrinting bookFoldPrintingSheets drawingGridHorizontalSpacing drawingGridVerticalSpacing displayHorizontalDrawingGridEvery displayVerticalDrawingGridEvery doNotUseMarginsForDrawingGridOrigin drawingGridHorizontalOrigin drawingGridVerticalOrigin doNotShadeFormData noPunctuationKerning characterSpacingControl printTwoOnOne strictFirstAndLastChars noLineBreaksAfter noLineBreaksBefore savePreviewPicture doNotValidateAgainstSchema saveInvalidXml ignoreMixedContent alwaysShowPlaceholderText doNotDemarcateInvalidXml saveXmlDataOnly useXSLTWhenSaving saveThroughXslt showXMLTags alwaysMergeEmptyNamespace updateFields hdrShapeDefaults footnotePr endnotePr compat docVars rsids mathPr attachedSchema themeFontLang clrSchemeMapping',
};
const RANK = Object.fromEntries(Object.entries(ORDER).map(([k, v]) => [k, Object.fromEntries(v.split(' ').map((t, i) => [t, i]))]));
// Elements that may appear more than once (the rest merge: two features setting w:spacing share one element).
const REPEATABLE = new Set(['headerReference', 'footerReference']);

/** One child element: `tag` without the w: prefix, its attributes (w: added unless the name has a prefix), inner XML. */
export const el = (tag, attrs = {}, inner = '') => ({ tag, attrs: { ...attrs }, inner });

/** A property element (rPr, pPr, sectPr, settings) from a list of children: merged where the same element is set twice, in schema order. */
export function props(kind, children, { wrap = true } = {}) {
  const merged = [];
  for (const c of children.filter(Boolean)) {
    const same = !REPEATABLE.has(c.tag) && merged.find((m) => m.tag === c.tag);
    if (same) { Object.assign(same.attrs, c.attrs); same.inner += c.inner; } else merged.push({ tag: c.tag, attrs: { ...c.attrs }, inner: c.inner });
  }
  const rank = RANK[kind];
  for (const m of merged) if (!(m.tag in rank)) throw new Error(`${kind}: unknown child ${m.tag}`);
  merged.sort((a, b) => rank[a.tag] - rank[b.tag]);
  const xml = merged.map(({ tag, attrs, inner }) => {
    const at = Object.entries(attrs).filter(([, v]) => v != null).map(([k, v]) => ` ${k.includes(':') ? k : 'w:' + k}="${esc(v)}"`).join('');
    return inner ? `<w:${tag}${at}>${inner}</w:${tag}>` : `<w:${tag}${at}/>`;
  }).join('');
  if (!wrap) return xml;
  return xml ? `<w:${kind}>${xml}</w:${kind}>` : '';
}

export const run = (text, rpr = [], { del = false } = {}) => `<w:r>${props('rPr', rpr)}${String(text).split('\t').map((t, n) => (n ? '<w:tab/>' : '') + (t ? `<w:${del ? 'delText' : 't'} xml:space="preserve">${esc(t)}</w:${del ? 'delText' : 't'}>` : '')).join('')}</w:r>`;
export const para = (inner, ppr = []) => `<w:p>${props('pPr', ppr)}${inner}</w:p>`;
/** A container's content must end with a paragraph (a table cell, a note, a header, a text box). */
export const endsWithPara = (blocks) => (blocks.length && blocks[blocks.length - 1].startsWith('<w:p>') || blocks.length && blocks[blocks.length - 1].startsWith('<w:p ') ? blocks : [...blocks, '<w:p/>']);

// ---- seeded randomness: mulberry32 ----

export function rng(seed) {
  let a = seed >>> 0;
  const next = () => { a = (a + 0x6d2b79f5) >>> 0; let t = a; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
  const int = (lo, hi) => lo + Math.floor(next() * (hi - lo + 1));
  const pick = (arr) => arr[Math.floor(next() * arr.length)];
  return { next, int, pick };
}
export const hash = (s) => { let h = 2166136261; for (const c of s) h = Math.imul(h ^ c.charCodeAt(0), 16777619); return h >>> 0; };

const WORDS = ('ability account action active address advice agency amount answer approach area argument article aspect attempt audience balance basis behavior benefit board border branch budget building camera capital career carbon center chapter choice circle climate coffee college column comment company concept contract corner council country course credit culture current custom danger debate decade degree demand design detail device dinner doctor drawing driver economy edition effect effort energy engine estate evening event evidence example factor family farmer feature figure finance flight floor flower focus forest format frame friend future garden goal ground growth habit health history holiday horse hotel house image impact income index injury island issue journey kitchen language leader lesson letter level library limit market master matter meaning measure media member memory method minute model moment money morning mother motion nation nature network notice number object office option orange order output owner paper parent party pattern period person phone picture piece place planet player pocket policy portion power practice present pressure price problem process product program project quality question radio range reason record region report request research result river sample school science season second section sector series service signal silver simple sister skill society source space speech spirit spring square stage standard station status story strategy street structure student studio style subject summer supply surface system table target teacher theory thing thought ticket topic tower track trade travel truth union unit value version village vision volume water weather weekend window winter worker world writer yellow').split(' ');

export function textSource(seed) {
  const r = rng(seed);
  const words = (n) => Array.from({ length: n }, () => r.pick(WORDS)).join(' ');
  const sentence = (n = r.int(12, 22)) => { const w = words(n); return w[0].toUpperCase() + w.slice(1) + '.'; };
  return { ...r, words, sentence, sentences: (k) => Array.from({ length: k }, () => sentence()).join(' ') };
}

// ---- a small PNG of one colour ----

export function png(w, h, [r, g, b]) {
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

// ---- the package ----

const T = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const fontEl = (f) => `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/>`;

export function stylesXml({ defaultFont = 'Calibri', defaultSize = 22 } = {}) {
  const ps = (id, name, ppr = '', rpr = '', extra = '') => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/>${extra}<w:qFormat/>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}</w:style>`;
  const cs = (id, name, rpr) => `<w:style w:type="character" w:styleId="${id}"><w:name w:val="${name}"/><w:rPr>${rpr}</w:rPr></w:style>`;
  const B = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>`
    + `<w:docDefaults><w:rPrDefault><w:rPr>${fontEl(defaultFont)}<w:sz w:val="${defaultSize}"/><w:szCs w:val="${defaultSize}"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
    + ps('Heading1', 'heading 1', '<w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="80"/><w:outlineLvl w:val="0"/>', `${fontEl('Calibri Light')}<w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/>`, '<w:next w:val="Normal"/>')
    + ps('Heading2', 'heading 2', '<w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="1"/>', `${fontEl('Calibri Light')}<w:color w:val="2F5496"/><w:sz w:val="26"/><w:szCs w:val="26"/>`, '<w:next w:val="Normal"/>')
    + ps('Title', 'Title', '<w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/>', `${fontEl('Calibri Light')}<w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/>`, '<w:next w:val="Normal"/>')
    + ps('Quote', 'Quote', '<w:spacing w:before="200"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/>', '<w:i/><w:iCs/><w:color w:val="404040"/>', '<w:next w:val="Normal"/>')
    + ps('ListParagraph', 'List Paragraph', '<w:ind w:left="720"/><w:contextualSpacing/>')
    + ps('Header', 'header', '<w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')
    + ps('Footer', 'footer', '<w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')
    + ps('FootnoteText', 'footnote text', '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>')
    + ps('EndnoteText', 'endnote text', '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>')
    + ps('CommentText', 'annotation text', '<w:spacing w:line="240" w:lineRule="auto"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>')
    + cs('FootnoteReference', 'footnote reference', '<w:vertAlign w:val="superscript"/>')
    + cs('EndnoteReference', 'endnote reference', '<w:vertAlign w:val="superscript"/>')
    + cs('CommentReference', 'annotation reference', '<w:sz w:val="16"/><w:szCs w:val="16"/>')
    + cs('Hyperlink', 'Hyperlink', '<w:color w:val="0563C1"/><w:u w:val="single"/>')
    + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
    + `<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(B).join('')}</w:tblBorders></w:tblPr></w:style>`
    + `<w:style w:type="table" w:styleId="TestBanded"><w:name w:val="Test Banded"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblBorders>${['top', 'bottom', 'insideH'].map(B).join('')}</w:tblBorders></w:tblPr>`
    + '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF"/></w:rPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr>'
    + '<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr></w:style>'
    + '</w:styles>';
}

export function numberingXml() {
  const lvl = (ilvl, fmt, text, { font, hanging = 360, isLgl } = {}) => `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/>${isLgl ? '<w:isLgl/>' : ''}<w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (ilvl + 1)}" w:hanging="${hanging}"/></w:pPr>${font ? `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr>` : ''}</w:lvl>`;
  const levels = (f) => Array.from({ length: 9 }, (_, i) => f(i)).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${NS}>`
    + `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${levels((i) => lvl(i, 'bullet', ['', 'o', ''][i % 3], { font: ['Symbol', 'Courier New', 'Wingdings'][i % 3] }))}</w:abstractNum>`
    + `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${levels((i) => lvl(i, ['decimal', 'lowerLetter', 'lowerRoman'][i % 3], `%${i + 1}.`))}</w:abstractNum>`
    + `<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="multilevel"/>${levels((i) => lvl(i, 'decimal', Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.') + '.', { hanging: 360 + 216 * i }))}</w:abstractNum>`
    + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>'
    + '</w:numbering>';
}

const NOTE_SEP = (tag) => `<w:${tag} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:${tag}>`
  + `<w:${tag} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:${tag}>`;

/**
 * Zips a document. `doc`: { body, settings: [el], styles: {defaultFont, defaultSize}, headers: {default|first|even: blocks},
 * footers: same, footnotes: [{id, blocks}], endnotes: same, comments: [{id, blocks}], media: {name: bytes} }.
 * Relationships are written as placeholders while the content is made ({{REL:type:target}}), because the same content
 * can land in the body, a header or a note: each part gets its own relationships for the placeholders it holds.
 */
export function pack(doc) {
  const files = {};
  const types = [];
  const rels = { 'word/document.xml': [] };
  const addPart = (path, xml, type) => { files[path] = xml; types.push(`<Override PartName="/${path}" ContentType="${type}"/>`); };
  const docRel = (type, target) => { const id = `rIdP${rels['word/document.xml'].length + 1}`; rels['word/document.xml'].push(`<Relationship Id="${id}" Type="${T}/${type}" Target="${target}"/>`); return id; };

  docRel('styles', 'styles.xml'); addPart('word/styles.xml', stylesXml(doc.styles), `${CT}.styles+xml`);
  docRel('numbering', 'numbering.xml'); addPart('word/numbering.xml', numberingXml(), `${CT}.numbering+xml`);
  const settings = props('settings', [
    el('defaultTabStop', { val: 720 }), el('characterSpacingControl', { val: 'doNotCompress' }),
    el('compat', {}, `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${doc.compat || 15}"/>`),
    ...(doc.settings || []),
  ], { wrap: false });
  docRel('settings', 'settings.xml'); addPart('word/settings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${NS}>${settings}</w:settings>`, `${CT}.settings+xml`);

  const notes = (kind, list) => {
    if (!list?.length) return;
    const tag = kind.slice(0, -1);
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${kind} ${NS}>${NOTE_SEP(tag)}${list.map((n) => `<w:${tag} w:id="${n.id}">${endsWithPara(n.blocks).join('')}</w:${tag}>`).join('')}</w:${kind}>`;
    docRel(kind, `${kind}.xml`); addPart(`word/${kind}.xml`, xml, `${CT}.${kind}+xml`);
  };
  notes('footnotes', doc.footnotes); notes('endnotes', doc.endnotes);
  if (doc.comments?.length) {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${NS}>${doc.comments.map((c) => `<w:comment w:id="${c.id}" w:author="Tester" w:date="2026-01-01T00:00:00Z" w:initials="T">${endsWithPara(c.blocks).join('')}</w:comment>`).join('')}</w:comments>`;
    docRel('comments', 'comments.xml'); addPart('word/comments.xml', xml, `${CT}.comments+xml`);
  }
  // Headers and footers: the section properties refer to them by placeholders {{HDR:default}} etc.
  const refs = {};
  for (const [kind, tag, ct] of [['headers', 'hdr', 'header'], ['footers', 'ftr', 'footer']]) {
    for (const [which, blocks] of Object.entries(doc[kind] || {})) {
      const name = `${ct}-${which}.xml`;
      refs[`${ct}:${which}`] = docRel(ct, name);
      addPart(`word/${name}`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${tag} ${NS}>${endsWithPara(blocks).join('')}</w:${tag}>`, `${CT}.${ct}+xml`);
    }
  }
  let body = doc.body.replace(/\{\{REF:(header|footer):(\w+)\}\}/g, (_, k, w) => refs[`${k}:${w}`]);
  files['word/document.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`;
  types.push(`<Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/>`);

  // Placeholders for pictures and hyperlinks, per part.
  for (const path of Object.keys(files)) {
    if (!path.startsWith('word/') || !path.endsWith('.xml')) continue;
    const list = rels[path] || (rels[path] = []);
    const seen = new Map();
    files[path] = files[path].replace(/\{\{REL:(image|hyperlink):([^}]+)\}\}/g, (_, type, target) => {
      const key = type + target;
      if (!seen.has(key)) {
        const id = `rIdX${list.length + 1}`;
        list.push(`<Relationship Id="${id}" Type="${T}/${type}" Target="${esc(target)}"${type === 'hyperlink' ? ' TargetMode="External"' : ''}/>`);
        seen.set(key, id);
      }
      return seen.get(key);
    });
  }
  for (const [name, bytes] of Object.entries(doc.media || {})) files[`word/media/${name}`] = bytes;
  for (const [path, list] of Object.entries(rels)) {
    if (!list.length) continue;
    const dir = path.slice(0, path.lastIndexOf('/')), base = path.slice(path.lastIndexOf('/') + 1);
    files[`${dir}/_rels/${base}.rels`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list.join('')}</Relationships>`;
  }
  files['_rels/.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${T}/officeDocument" Target="word/document.xml"/></Relationships>`;
  files['[Content_Types].xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>${types.join('')}</Types>`;
  return zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, typeof v === 'string' ? strToU8(v) : v])), { level: 6 });
}
