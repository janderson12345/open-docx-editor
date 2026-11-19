// Builds two fixtures of vertical spacing between plain paragraphs (features.mjs), each checked against Word:
// - spaces-lines.docx: 12pt Times New Roman lines at 1.08 (w:line="259"), between "X" lines, holding nothing but spaces:
//   an 18pt and a 12pt space, an 18pt space, two, an 18pt lowered space, all over a 12pt mark; and a 12pt space over an
//   18pt mark. Word makes each the mark's line (15pt apart, as the X lines; 21.75pt for the 18pt mark), not its
//   spaces': runs of only spaces don't size a line. (An 18pt space had made its line 22.35pt.)
// - columns-empty-section.docx: two lines, then a continuous section in two columns holding only its empty
//   section-break paragraph, then two lines in one column again. Word gives that empty paragraph its line, as in one
//   column (the text after it 29.25pt under the line before, two lines; it had come up one, onto the empty paragraph's).
// - shadow-lines.docx: 15pt Times New Roman lines, single, between plain ones: with Word's preset outer shadow (w14:shadow,
//   4pt blur, 3pt at 45 degrees) on the run, from a paragraph style, and on an empty paragraph of that style; then an 8pt
//   blur on 10pt text. Word's PDFs: each shadowed line 2.7pt taller (3.6px), the 10pt one 4pt (5.33px); they had none.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const F = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>';
const run = (sz, t, extra = '') => `<w:r><w:rPr>${F}<w:sz w:val="${sz}"/>${extra}</w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
const para = (mark, runs) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="259" w:lineRule="auto"/><w:rPr>${F}<w:sz w:val="${mark}"/></w:rPr></w:pPr>${runs}</w:p>`;
const sub = '<w:vertAlign w:val="subscript"/>';
const page = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
const sect = (type, cols) => `<w:sectPr>${type ? '<w:type w:val="continuous"/>' : ''}${page}<w:cols w:space="720"${cols ? ` w:num="${cols}"` : ''}/></w:sectPr>`;

function make(name, body, last) {
  const files = unzipSync(blankDocxBytes());
  let doc = strFromU8(files['word/document.xml']);
  doc = doc.replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${body}${last}</w:body>`);
  files['word/document.xml'] = strToU8(doc);
  fs.writeFileSync(new URL(`./${name}`, import.meta.url), zipSync(files));
}

const X = para(24, run(24, 'X'));
const blanks = [
  para(24, run(36, ' ') + run(24, ' ')),
  para(24, run(36, ' ')),
  para(24, run(36, '  ')),
  para(24, run(36, ' ', sub)),
  para(36, run(24, ' ')),
];
make('spaces-lines.docx', X + blanks.map((b) => b + X).join(''), sect(false, 0));

const line = (t) => para(24, run(24, t));
const body = line('First line.')
  + `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="259" w:lineRule="auto"/>${sect(false, 0)}</w:pPr>${run(24, 'Last line of the first section.')}</w:p>`
  + `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="259" w:lineRule="auto"/>${sect(true, 2)}</w:pPr></w:p>`
  + line('After the empty columns section.') + line('One more line.');
make('columns-empty-section.docx', body, sect(true, 0));

const shadow = (blur, dist, dir) => `<w14:shadow w14:blurRad="${blur}" w14:dist="${dist}" w14:dir="${dir}" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl"><w14:srgbClr w14:val="000000"><w14:alpha w14:val="60000"/></w14:srgbClr></w14:shadow>`;
const preset = shadow(50800, 38100, 2700000);
const single = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
const sp = (sz, t, pStyle = '', rx = '') => `<w:p><w:pPr>${pStyle}${single}<w:rPr>${F}<w:sz w:val="${sz}"/></w:rPr></w:pPr>${t == null ? '' : `<w:r><w:rPr>${F}<w:sz w:val="${sz}"/>${rx}</w:rPr><w:t>${t}</w:t></w:r>`}</w:p>`;
const styled = '<w:pStyle w:val="Shadowed"/>';
const shBody = [
  sp(30, 'Plain'), sp(30, 'Plain'), sp(30, 'Shadow on the run', '', preset), sp(30, 'Plain'),
  sp(30, 'Shadow from the style', styled), sp(30, 'Plain'), sp(30, null, styled), sp(30, 'Plain'),
  sp(20, 'Small plain'), sp(20, 'Small with a wide blur', '', shadow(101600, 0, 0)), sp(20, 'Small plain'),
].join('');
{
  const files = unzipSync(blankDocxBytes());
  let doc = strFromU8(files['word/document.xml']);
  doc = doc.replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${shBody}${sect(false, 0)}</w:body>`);
  if (!/xmlns:w14=/.test(doc)) doc = doc.replace(/<w:document\b/, '<w:document xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"');
  files['word/document.xml'] = strToU8(doc);
  let styles = strFromU8(files['word/styles.xml']);
  if (!/xmlns:w14=/.test(styles)) styles = styles.replace(/<w:styles\b/, '<w:styles xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"');
  styles = styles.replace('</w:styles>', `<w:style w:type="paragraph" w:customStyle="1" w:styleId="Shadowed"><w:name w:val="Shadowed"/><w:basedOn w:val="Normal"/><w:rPr>${preset}</w:rPr></w:style></w:styles>`);
  files['word/styles.xml'] = strToU8(styles);
  fs.writeFileSync(new URL('./shadow-lines.docx', import.meta.url), zipSync(files));
}
