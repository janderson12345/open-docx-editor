// Builds test/fixtures/rowspan-mixed.docx, rowspan-inside.docx, vml-inline-group.docx, cols-break-landscape.docx and
// cols-footer.docx: a form's merged cells, framed text boxes and columns as Word lays them out (its prints of the same
// documents made by this script). Arial 8pt single, no cell margins.
// rowspan-mixed: a cell merged down rows of at least 291, 296, 296, 296, 296 and 225tw beside one-line cells, its six
//   paragraphs (spaced 31, 81, 1, 8 and 5tw before, two at 249/240 lines, one empty in 8.5pt) fitting the rows' own
//   height: Word keeps the rows at their minimum ("After table" 132.7px under "Before"); the rows' content alone is
//   less than the merged cell's, and its last row had been given the difference.
// rowspan-inside: the table's lines 1.5pt, the side cells' own 0.75pt, a cell merged down six rows of at least 296tw
//   whose lower parts set a top line of nil: each row is the 1.5pt inside line over its minimum in Word ("Side 6" 108px
//   under "Side 1"), the line under the part above counting though the part's own top is nil.
// vml-inline-group: an inline VML group 400 x 120pt of two text boxes, a title at its top and three lines lower down:
//   the line under it ("After group") 174.7px under "Before" in Word.
// cols-break-landscape: a continuous section of two unequal columns part way down a portrait page, a paragraph starting
//   with a column break, then one column again; the last section landscape. Word: the second column's lines level
//   with the first's; "After columns" 16.6px under the columns' last lines.
// cols-footer: a page of 600tw bottom margin and a one-line footer 405tw up; a continuous section of two columns and one
//   of one column part way down it, both 1440tw bottom margin and footer 720tw up, then a page section. Word fills the
//   page to its own bottom margin: 62 of the 80 lines after the columns on page 1.
// usage: node test/fixtures/mkrowspan.mjs [out dir]  (default: here)
import '../setup.js';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const OUT = process.argv[2] ? path.resolve(process.argv[2]) : path.dirname(fileURLToPath(import.meta.url));
const SINGLE = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
const ARIAL = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="16"/><w:szCs w:val="16"/>';
const NS = 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"';

/** A document from the blank one: its body, its last section's properties, a footer's paragraphs. */
function write(name, body, { last = null, footer = null } = {}) {
  const files = unzipSync(blankDocxBytes());
  const main = 'word/document.xml';
  let d = strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`);
  if (last) d = d.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/, last);
  if (!/xmlns:v=/.test(d)) d = d.replace('<w:document ', `<w:document ${NS} `);
  if (footer) {
    files['word/footer1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${footer}</w:ftr>`);
    files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'));
    const rels = 'word/_rels/document.xml.rels';
    files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdFt1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>'));
    if (!/xmlns:r=/.test(d)) d = d.replace('<w:document ', '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
  }
  files[main] = strToU8(d);
  const st = strFromU8(files['word/styles.xml']);
  files['word/styles.xml'] = strToU8(st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, `$1<w:pPr>${SINGLE}</w:pPr><w:rPr>${ARIAL}</w:rPr>$2`));
  fs.writeFileSync(path.join(OUT, name), zipSync(files));
}
const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const lines = (ts) => ts.map((t, i) => (i ? '<w:r><w:br/></w:r>' : '') + run(t)).join('');
const para = (content, ppr = '') => `<w:p><w:pPr>${ppr}</w:pPr>${content}</w:p>`;
const many = (k, n) => Array.from({ length: n }, (_, i) => para(run(`${k} ${i + 1}`))).join('');
const ln = (side, v) => (v == null ? '' : v === 'nil' ? `<w:${side} w:val="nil"/>` : `<w:${side} w:val="single" w:sz="${v}" w:space="0" w:color="auto"/>`);

/** A cell merged down over rows of the given minimum heights beside one-line cells. */
function table(heights, merged, { inside = 6, sideTop = null, contTop = null } = {}) {
  const B = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="${inside}" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
  const rows = heights.map((h, i) => {
    const mb = i ? ln('top', contTop) : '';
    const left = `<w:tc><w:tcPr><w:tcW w:w="7000" w:type="dxa"/><w:vMerge${i ? '' : ' w:val="restart"'}/>${mb ? `<w:tcBorders>${mb}</w:tcBorders>` : ''}</w:tcPr>${i ? para('') : merged}</w:tc>`;
    const side = `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${sideTop != null ? `<w:tcBorders>${ln('top', sideTop)}${ln('bottom', sideTop)}</w:tcBorders>` : ''}</w:tcPr>${para(run(`Side ${i + 1}`))}</w:tc>`;
    return `<w:tr>${h ? `<w:trPr><w:trHeight w:val="${h}"/></w:trPr>` : ''}${left}${side}</w:tr>`;
  }).join('');
  return `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>${B}<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>`
    + `<w:tblGrid><w:gridCol w:w="7000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows}</w:tbl>`;
}
const around = (x) => para(run('Before')) + x + para(run('After table')) + para(run('After two'));

write('rowspan-mixed.docx', around(table([291, 296, 296, 296, 296, 225], [
  para(run('Distribution'), '<w:spacing w:before="31"/>'),
  para(lines(['First office (1D)', 'Second office (1D)']), '<w:spacing w:before="81" w:line="249" w:lineRule="auto"/>'),
  para(run('Third office (1D)'), '<w:spacing w:before="1"/>'),
  para(run('Fourth office (1D)'), '<w:spacing w:before="8"/>'),
  para('', '<w:spacing w:before="5"/><w:rPr><w:sz w:val="17"/></w:rPr>'),
  para(lines(['Delivered by mail, in the suite of software named in the current standard,', 'to be sent by e-mail.']), '<w:spacing w:line="249" w:lineRule="auto"/>'),
].join(''))));
write('rowspan-inside.docx', around(table([296, 296, 296, 296, 296, 296], ['Merged line 1', 'Merged line 2', 'Merged line 3'].map((t) => para(run(t))).join(''), { inside: 12, sideTop: 6, contTop: 'nil' })));

// An inline VML group of two text boxes (group coordinates 8000 x 2400 over 400 x 120pt).
const box = (id, top, h, body) => `<v:shape id="${id}" type="#_x0000_t202" style="position:absolute;left:0;top:${top};width:8000;height:${h}" filled="f" strokecolor="black" strokeweight="1pt"><v:textbox inset="0,0,0,0"><w:txbxContent>${body}</w:txbxContent></v:textbox></v:shape>`;
const group = '<w:r><w:pict><v:group id="g1" style="width:400pt;height:120pt;mso-position-horizontal-relative:char;mso-position-vertical-relative:line" coordsize="8000,2400">'
  + '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
  + box('b2', 600, 1800, para(run('Body line one')) + para(run('Body line two')) + para(run('Body line three')))
  + box('b1', 0, 500, para(run('Group title')))
  + '<w10:anchorlock/></v:group></w:pict></w:r>';
write('vml-inline-group.docx', para(run('Before')) + para(group) + para(run('After group')) + para(run('After two')));

// Columns.
const MAR = (top, bottom, footer = 720) => `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="${top}" w:right="580" w:bottom="${bottom}" w:left="600" w:header="720" w:footer="${footer}" w:gutter="0"/>`;
const one = (top, bottom, { cont = false, footer, ref } = {}) => `<w:sectPr>${ref ? '<w:footerReference w:type="default" r:id="rIdFt1"/>' : ''}${cont ? '<w:type w:val="continuous"/>' : ''}${MAR(top, bottom, footer)}<w:cols w:space="720"/></w:sectPr>`;
const two = (top, bottom, footer = 720) => `<w:sectPr><w:type w:val="continuous"/>${MAR(top, bottom, footer)}<w:cols w:num="2" w:space="720" w:equalWidth="0"><w:col w:w="5210" w:space="40"/><w:col w:w="5810"/></w:cols></w:sectPr>`;
const brk = (s) => `<w:p><w:pPr>${s}</w:pPr></w:p>`;
const LAND = '<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1020" w:right="640" w:bottom="720" w:left="640" w:header="166" w:footer="522" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>';
const colB = para('<w:r><w:br w:type="column"/></w:r>' + run('Right head')) + many('Right item', 3);
write('cols-break-landscape.docx', many('Intro line', 12) + brk(one(720, 600)) + many('Left item', 4) + colB + brk(two(1500, 280))
  + para(run('After columns'), '<w:spacing w:before="62"/>') + many('After line', 3) + brk(one(1500, 280, { cont: true })) + para(run('Landscape page')), { last: LAND });
const FOOT = para(run('Footer line'));
write('cols-footer.docx', many('Intro line', 12) + brk(one(720, 600, { footer: 405, ref: true })) + many('Left item', 3) + colB + brk(two(1500, 1440, 720))
  + many('Line', 80) + brk(one(1500, 1440, { cont: true, footer: 720 })) + para(run('Last page')), { last: one(720, 600, { footer: 405 }), footer: FOOT });
