// The features the generator combines, and how a set of them becomes one document.
//
// Every document has the same shape: a title line naming its features, four content paragraphs, then plain
// paragraphs after them (so the layout after the features is measured too). Each content paragraph is a lead-in, a
// "featured" run, anything inserted after that run, and a sentence or two. A feature changes one part of that shape:
//   run      the featured run's properties (bold, a font, a size...)
//   inline   something inserted in each content paragraph (a tab, a field, a footnote reference, a picture...)
//   para     the content paragraphs' properties (alignment, indents, spacing, a style...)
//   block    a structure the content paragraphs are put in (a list, a table, a drop cap)
//   where    where the content goes (a text box, a frame, a footnote, a header or footer)
//   section  the section's page (size, margins, columns, borders, headers and footers) or a section break around the content
//   doc      the whole document (settings, default font, length)
// Features in the same group are alternatives (one table kind, one alignment...), never combined. A feature marked
// bodyOnly can't go where a `where` feature puts the content (a page break in a header, a footnote in a text box).
import { el, props, run, para, endsWithPara, textSource } from './ooxml.mjs';
import { png } from './ooxml.mjs';

const DATE = '2026-01-01T00:00:00Z';
const fld = (instr, shown) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${shown}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const rp = (tag, attrs, inner) => (ctx) => ctx.rpr.push(el(tag, attrs, inner));
const pp = (tag, attrs, inner) => (ctx) => ctx.ppr.push(el(tag, attrs, inner));
const box4 = (sz, space, color = 'auto') => ['top', 'left', 'bottom', 'right'].map((s) => `<w:${s} w:val="single" w:sz="${sz}" w:space="${space}" w:color="${color}"/>`).join('');
const EMU = 914400; // per inch

function inlinePicture(ctx, cx, cy, file) {
  const id = ctx.id();
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>${pictureGraphic(cx, cy, file)}</wp:inline></w:drawing></w:r>`;
}
const pictureGraphic = (cx, cy, file) => `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="${file}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{{REL:image:media/${file}}}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
function anchored(ctx, { cx, cy, wrap, posH, posV, graphic, behind = 0 }) {
  const id = ctx.id();
  return `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${251658240 + id}" behindDoc="${behind}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>${posH}${posV}<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrap}<wp:docPr id="${id}" name="Shape ${id}"/><wp:cNvGraphicFramePr/>${graphic}</wp:anchor>`;
}
const usePicture = (ctx, name, colour) => { ctx.media[name] ||= png(24, 16, colour); return name; };

function table({ style = 'TableGrid', widths, rows, tblPr = '', look = '04A0', layout }) {
  const grid = `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
  const pr = `<w:tblPr><w:tblStyle w:val="${style}"/>${tblPr.includes('tblpPr') ? tblPr : ''}<w:tblW w:w="${layout === 'auto' ? 0 : widths.reduce((a, b) => a + b, 0)}" w:type="${layout === 'auto' ? 'auto' : 'dxa'}"/>${layout === 'fixed' ? '<w:tblLayout w:type="fixed"/>' : ''}<w:tblLook w:val="${look}" w:firstRow="${look === '04A0' ? 1 : 0}" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>`;
  const tr = (row) => `<w:tr>${row.trPr ? `<w:trPr>${row.trPr}</w:trPr>` : ''}${row.cells.map((c) => `<w:tc><w:tcPr><w:tcW w:w="${layout === 'auto' ? 0 : c.w}" w:type="${layout === 'auto' ? 'auto' : 'dxa'}"/>${c.tcPr || ''}</w:tcPr>${endsWithPara(c.blocks).join('')}</w:tc>`).join('')}</w:tr>`;
  return `<w:tbl>${pr}${grid}${rows.map(tr).join('')}</w:tbl>`;
}
const cell = (w, blocks, tcPr = '') => ({ w, blocks, tcPr });
const short = (ctx) => para(run(ctx.t.words(2)));

const VML_WATERMARK = '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas><v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/><v:textpath on="t" fitshape="t"/><o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>'
  + '<v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s1025" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:412.4pt;height:247.45pt;rotation:315;z-index:-251658752;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="DRAFT"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r></w:p>';

const F = [];
const def = (id, kind, group, desc, apply, extra = {}) => F.push({ id, kind, group: group || id, desc, apply, ...extra });

// ---- run ----
def('bold', 'run', '', 'bold', (c) => { rp('b')(c); rp('bCs')(c); });
def('italic', 'run', '', 'italic', (c) => { rp('i')(c); rp('iCs')(c); });
def('underline', 'run', 'u', 'single underline', rp('u', { val: 'single' }));
def('underline-double', 'run', 'u', 'double underline', rp('u', { val: 'double' }));
def('underline-words', 'run', 'u', 'underline words only', rp('u', { val: 'words' }));
def('strike', 'run', 'strike', 'strikethrough', rp('strike'));
def('dstrike', 'run', 'strike', 'double strikethrough', rp('dstrike'));
def('caps', 'run', 'case', 'all caps', rp('caps'));
def('smallcaps', 'run', 'case', 'small caps', rp('smallCaps'));
def('superscript', 'run', 'vert', 'superscript', rp('vertAlign', { val: 'superscript' }));
def('subscript', 'run', 'vert', 'subscript', rp('vertAlign', { val: 'subscript' }));
def('size-large', 'run', 'size', '18pt text', (c) => { rp('sz', { val: 36 })(c); rp('szCs', { val: 36 })(c); });
def('size-small', 'run', 'size', '8pt text', (c) => { rp('sz', { val: 16 })(c); rp('szCs', { val: 16 })(c); });
for (const [id, font] of [['font-times', 'Times New Roman'], ['font-courier', 'Courier New'], ['font-cambria', 'Cambria'], ['font-georgia', 'Georgia'], ['font-arial', 'Arial']]) {
  def(id, 'run', 'font', `${font} font`, rp('rFonts', { ascii: font, hAnsi: font, cs: font }));
}
def('color', 'run', '', 'coloured text', rp('color', { val: 'C00000' }));
def('highlight', 'run', '', 'highlighted', rp('highlight', { val: 'yellow' }));
def('run-shading', 'run', '', 'shaded run', rp('shd', { val: 'clear', color: 'auto', fill: 'D9D9D9' }));
def('char-expanded', 'run', 'charspace', 'character spacing expanded 2pt', rp('spacing', { val: 40 }));
def('char-condensed', 'run', 'charspace', 'character spacing condensed 0.75pt', rp('spacing', { val: -15 }));
def('char-scale', 'run', '', 'characters scaled to 150%', rp('w', { val: 150 }));
def('raised', 'run', '', 'raised 3pt', rp('position', { val: 6 }));
def('run-border', 'run', '', 'run border', rp('bdr', { val: 'single', sz: 4, space: 0, color: 'auto' }));
def('outline', 'run', 'effect', 'outline effect', rp('outline'));
def('shadow', 'run', 'effect', 'shadow effect', rp('shadow'));
def('emboss', 'run', 'effect', 'emboss effect', rp('emboss'));

// ---- inline ----
def('tab', 'inline', '', 'tab characters at the default stops', (c) => c.inlines.push(() => '<w:r><w:tab/></w:r>' + run(c.t.words(2)) + '<w:r><w:tab/></w:r>' + run(c.t.words(1))));
def('line-break', 'inline', '', 'a manual line break', (c) => c.inlines.push(() => '<w:r><w:br/></w:r>' + run(c.t.words(3))));
def('page-break', 'inline', 'break', 'a page break in the second paragraph', (c) => c.inlines.push((i) => (i === 1 ? '<w:r><w:br w:type="page"/></w:r>' + run(c.t.words(3)) : '')), { bodyOnly: true });
def('column-break', 'inline', 'break', 'a column break in the third paragraph', (c) => c.inlines.push((i) => (i === 2 ? '<w:r><w:br w:type="column"/></w:r>' + run(c.t.words(3)) : '')), { bodyOnly: true });
def('field-page', 'inline', '', 'a PAGE field', (c) => c.inlines.push(() => run(' page ') + fld('PAGE', '1')));
def('field-numpages', 'inline', '', 'a NUMPAGES field', (c) => c.inlines.push(() => run(' of ') + fld('NUMPAGES', '1')));
def('symbol', 'inline', '', 'a Wingdings symbol', (c) => c.inlines.push(() => '<w:r><w:sym w:font="Wingdings" w:char="F04A"/></w:r>' + run(' ' + c.t.words(1))));
def('special-chars', 'inline', '', 'no-break space, no-break hyphen and soft hyphen', (c) => c.inlines.push(() => run(` ${c.t.words(1)}\u00a0${c.t.words(1)} `) + '<w:r><w:t>well</w:t></w:r><w:r><w:noBreakHyphen/></w:r><w:r><w:t xml:space="preserve">known </w:t></w:r><w:r><w:t>extra</w:t></w:r><w:r><w:softHyphen/></w:r><w:r><w:t xml:space="preserve">ordinary </w:t></w:r>'));
def('footnote-ref', 'inline', 'note-ref', 'a footnote in each paragraph', (c) => c.inlines.push(() => {
  const id = c.note('footnotes', [para('<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' + run(' ' + c.t.sentence()), [el('pStyle', { val: 'FootnoteText' })])]);
  return `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
}), { bodyOnly: true });
def('endnote-ref', 'inline', 'note-ref', 'an endnote in each paragraph', (c) => c.inlines.push(() => {
  const id = c.note('endnotes', [para('<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>' + run(' ' + c.t.sentence()), [el('pStyle', { val: 'EndnoteText' })])]);
  return `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;
}), { bodyOnly: true });
def('comment', 'inline', '', 'a comment on the featured run', (c) => c.runWraps.push({ order: 30, fn: (xml) => {
  const id = c.comments.length;
  c.comments.push({ id, blocks: [para('<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>' + run(c.t.sentence()), [el('pStyle', { val: 'CommentText' })])] });
  return `<w:commentRangeStart w:id="${id}"/>${xml}<w:commentRangeEnd w:id="${id}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
} }), { bodyOnly: true });
def('hyperlink', 'inline', '', 'the featured run is a hyperlink', (c) => { c.rpr.push(el('rStyle', { val: 'Hyperlink' })); c.runWraps.push({ order: 20, fn: (xml) => `<w:hyperlink r:id="{{REL:hyperlink:https://example.com/page}}" w:history="1">${xml}</w:hyperlink>` }); });
def('content-control', 'inline', '', 'the featured run in a content control', (c) => c.runWraps.push({ order: 25, fn: (xml) => `<w:sdt><w:sdtPr><w:alias w:val="Field"/><w:id w:val="${c.id()}"/></w:sdtPr><w:sdtContent>${xml}</w:sdtContent></w:sdt>` }));
def('tracked-insert', 'inline', 'track', 'the featured run is a tracked insertion', (c) => c.runWraps.push({ order: 10, fn: (xml) => `<w:ins w:id="${c.id()}" w:author="Tester" w:date="${DATE}">${xml}</w:ins>` }));
def('tracked-delete', 'inline', 'track', 'the featured run is a tracked deletion', (c) => { c.featDel = true; c.runWraps.push({ order: 10, fn: (xml) => `<w:del w:id="${c.id()}" w:author="Tester" w:date="${DATE}">${xml}</w:del>` }); });
def('tracked-format', 'inline', '', 'the featured run has a tracked formatting change', (c) => c.rprAt.push(() => [el('rPrChange', { id: c.id(), author: 'Tester', date: DATE }, '<w:rPr/>')]));
def('image-inline', 'inline', '', 'an inline picture', (c) => c.inlines.push(() => inlinePicture(c, 0.5 * EMU, 0.33 * EMU, usePicture(c, 'blue.png', [0x2f, 0x55, 0x97]))));
def('image-float', 'inline', 'float', 'a picture floating at the right, text wrapped square', (c) => c.inlines.push((i) => (i !== 0 ? '' : `<w:r><w:drawing>${anchored(c, {
  cx: 1.5 * EMU, cy: 1 * EMU, wrap: '<wp:wrapSquare wrapText="bothSides"/>',
  posH: '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>', posV: '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
  graphic: pictureGraphic(1.5 * EMU, 1 * EMU, usePicture(c, 'green.png', [0x2f, 0x80, 0x40])),
})}</w:drawing></w:r>`)), { bodyOnly: true });
def('image-topbottom', 'inline', 'float', 'a picture with text above and below it', (c) => c.inlines.push((i) => (i !== 1 ? '' : `<w:r><w:drawing>${anchored(c, {
  cx: 2 * EMU, cy: 0.75 * EMU, wrap: '<wp:wrapTopAndBottom/>',
  posH: '<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>', posV: '<wp:positionV relativeFrom="paragraph"><wp:posOffset>91440</wp:posOffset></wp:positionV>',
  graphic: pictureGraphic(2 * EMU, 0.75 * EMU, usePicture(c, 'red.png', [0xc0, 0x40, 0x30])),
})}</w:drawing></w:r>`)), { bodyOnly: true });

// ---- para ----
for (const [id, v] of [['align-center', 'center'], ['align-right', 'right'], ['justify', 'both'], ['distribute', 'distribute']]) def(id, 'para', 'align', `paragraphs aligned ${v}`, pp('jc', { val: v }));
def('indent-sides', 'para', 'indent', 'left and right indents of 0.5in', pp('ind', { left: 720, right: 720 }));
def('indent-first', 'para', 'indent', 'first-line indent of 0.5in', pp('ind', { firstLine: 720 }));
def('indent-hanging', 'para', 'indent', 'hanging indent of 0.5in', pp('ind', { left: 720, hanging: 720 }));
def('space-before-after', 'para', 'pspace', '12pt before and after', pp('spacing', { before: 240, after: 240 }));
def('space-none', 'para', 'pspace', 'no space before or after', pp('spacing', { before: 0, after: 0 }));
def('line-exact', 'para', 'line', 'exactly 15pt line spacing', pp('spacing', { line: 300, lineRule: 'exact' }));
def('line-double', 'para', 'line', 'double line spacing', pp('spacing', { line: 480, lineRule: 'auto' }));
def('line-1.5', 'para', 'line', '1.5 line spacing', pp('spacing', { line: 360, lineRule: 'auto' }));
def('line-atleast', 'para', 'line', 'at least 20pt line spacing', pp('spacing', { line: 400, lineRule: 'atLeast' }));
def('para-border', 'para', '', 'a box border round each paragraph', pp('pBdr', {}, box4(6, 4)));
def('para-shading', 'para', '', 'shaded paragraphs', pp('shd', { val: 'clear', color: 'auto', fill: 'F2F2F2' }));
def('tab-leader', 'para', 'tabs', 'a right tab stop with a dot leader', (c) => { pp('tabs', {}, '<w:tab w:val="right" w:leader="dot" w:pos="8640"/>')(c); c.inlines.push(() => '<w:r><w:tab/></w:r>' + run('42')); });
def('tab-center-decimal', 'para', 'tabs', 'centre and decimal tab stops', (c) => { pp('tabs', {}, '<w:tab w:val="center" w:pos="3600"/><w:tab w:val="decimal" w:pos="6480"/>')(c); c.inlines.push(() => '<w:r><w:tab/></w:r>' + run('middle') + '<w:r><w:tab/></w:r>' + run('1234.56')); });
for (const [id, style] of [['heading1', 'Heading1'], ['heading2', 'Heading2'], ['title', 'Title'], ['quote', 'Quote']]) def(id, 'para', 'pstyle', `paragraphs in the ${style} style`, pp('pStyle', { val: style }));
def('page-break-before', 'para', '', 'the third paragraph starts a page', (c) => c.pprAt.push((i) => (i === 2 ? [el('pageBreakBefore')] : [])), { bodyOnly: true });
def('keep-next', 'para', '', 'keep with next', pp('keepNext'));
def('keep-lines', 'para', '', 'keep lines together', pp('keepLines'));
def('widow-off', 'para', '', 'widow and orphan control off', pp('widowControl', { val: 0 }));
def('contextual-spacing', 'para', '', "don't add space between paragraphs of the same style", pp('contextualSpacing'));

// ---- block ----
const list = (numId, levels) => (c) => { c.list = true; c.pprAt.push((i) => [el('numPr', {}, `<w:ilvl w:val="${levels[i % levels.length]}"/><w:numId w:val="${numId}"/>`)]); };
def('list-bullet', 'block', 'list', 'a bulleted list', list(1, [0, 0, 1, 0]));
def('list-number', 'block', 'list', 'a numbered list', list(2, [0, 1, 1, 0]));
def('list-multilevel', 'block', 'list', 'a multilevel 1.1.1 list', list(3, [0, 1, 2, 1]));
const wrap = (order, fn) => (c) => c.wraps.push({ order, fn: (b) => fn(c, b) });
def('table-grid', 'block', 'table', 'a 2x2 table with borders', wrap(30, (c, b) => [table({ widths: [4680, 4680], rows: [{ cells: [cell(4680, [b[0], b[1]]), cell(4680, [b[2]])] }, { cells: [cell(4680, [b[3]]), cell(4680, [short(c)])] }] })]));
def('table-merged', 'block', 'table', 'a table with merged cells', wrap(30, (c, b) => [table({ widths: [3120, 3120, 3120], rows: [
  { cells: [cell(6240, [b[0]], '<w:gridSpan w:val="2"/><w:vAlign w:val="center"/>'), cell(3120, [b[1]], '<w:vMerge w:val="restart"/><w:vAlign w:val="bottom"/>')] },
  { cells: [cell(3120, [b[2]]), cell(3120, [b[3]]), cell(3120, ['<w:p/>'], '<w:vMerge/>')] },
] })]));
def('table-nested', 'block', 'table', 'a table inside a table cell', wrap(30, (c, b) => [table({ widths: [5400, 3960], rows: [{ cells: [
  cell(5400, [b[0], table({ widths: [2500, 2500], rows: [{ cells: [cell(2500, [b[1]]), cell(2500, [b[2]])] }] }), '<w:p/>']), cell(3960, [b[3]]),
] }] })]));
def('table-banded', 'block', 'table', 'a table style with a header row and banded rows', wrap(30, (c, b) => [table({ style: 'TestBanded', widths: [5000, 2180, 2180], rows: [
  { cells: ['Column A', 'Column B', 'Column C'].map((t, k) => cell([5000, 2180, 2180][k], [para(run(t))])) },
  ...b.map((x) => ({ cells: [cell(5000, [x]), cell(2180, [short(c)]), cell(2180, [short(c)])] })),
] })]));
def('table-header-repeat', 'block', 'table', 'a long table whose header row repeats on each page', wrap(30, (c, b) => [table({ widths: [6360, 3000], rows: [
  { trPr: '<w:tblHeader/>', cells: [cell(6360, [para(run('Repeated header'))]), cell(3000, [para(run('Second header'))])] },
  ...b.map((x) => ({ cells: [cell(6360, [x]), cell(3000, [short(c)])] })),
  ...Array.from({ length: 40 }, (_, k) => ({ trPr: k % 3 ? '' : '<w:cantSplit/>', cells: [cell(6360, [para(run(`Row ${k + 1} ${c.t.words(6)}`))]), cell(3000, [short(c)])] })),
] })]));
def('table-float', 'block', 'table', 'a floating table at the right with text wrapping round it', wrap(30, (c, b) => [table({
  widths: [3600], tblPr: '<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="right" w:tblpY="1"/>',
  rows: [{ cells: [cell(3600, [b[0]])] }, { cells: [cell(3600, [b[1]])] }],
}), b[2], b[3]]));
def('table-fixed', 'block', 'table', 'a fixed-layout table with a narrow first column', wrap(30, (c, b) => [table({ layout: 'fixed', widths: [1800, 7560], rows: b.map((x) => ({ cells: [cell(1800, [short(c)]), cell(7560, [x])] })) })]));
def('table-autofit', 'block', 'table', 'an autofit table sized by its content', wrap(30, (c, b) => [table({ layout: 'auto', widths: [3120, 3120, 3120], rows: [
  { cells: [cell(0, [para(run(c.t.words(1)))]), cell(0, [b[0]]), cell(0, [para(run(c.t.words(3)))])] },
  { cells: [cell(0, [b[1]]), cell(0, [para(run(c.t.words(1)))]), cell(0, [b[2]])] },
  { cells: [cell(0, [para(run('x'))]), cell(0, [b[3]]), cell(0, [para(run(c.t.words(2)))])] },
] })]));
def('drop-cap', 'block', 'dropcap', 'a three-line drop cap before the first paragraph', wrap(65, (c, b) => [
  para('<w:r><w:rPr><w:position w:val="-9"/><w:sz w:val="111"/></w:rPr><w:t>D</w:t></w:r>', [el('keepNext'), el('framePr', { dropCap: 'drop', lines: 3, wrap: 'around', vAnchor: 'text', hAnchor: 'text' }), el('spacing', { after: 0, line: 827, lineRule: 'exact' }), el('textAlignment', { val: 'baseline' })]),
  ...b,
]), { bodyOnly: true, notWith: ['table'] });

// ---- where ----
def('textbox', 'where', 'where', 'the content in a floating text box', wrap(60, (c, b) => [para(`<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>${anchored(c, {
  cx: 3.4 * EMU, cy: 4.2 * EMU, wrap: '<wp:wrapSquare wrapText="bothSides"/>',
  posH: '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>', posV: '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
  graphic: `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${3.4 * EMU}" cy="${4.2 * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent>${endsWithPara(b).join('')}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:spAutoFit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>`,
})}</w:drawing></mc:Choice></mc:AlternateContent></w:r>` + run(c.t.sentences(3)))]));
def('frame', 'where', 'where', 'the content paragraphs in a frame at the right', pp('framePr', { w: 4320, hSpace: 180, wrap: 'around', vAnchor: 'text', hAnchor: 'margin', xAlign: 'right', y: 1 }), { notWith: ['table', 'dropcap'] });
def('footnote-content', 'where', 'where', 'the content in a footnote', wrap(60, (c, b) => {
  const id = c.note('footnotes', [para('<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' + run(' The footnote holding the content.'), [el('pStyle', { val: 'FootnoteText' })]), ...b]);
  return [para(run('This paragraph has the footnote that holds the content.') + `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`)];
}));
def('endnote-content', 'where', 'where', 'the content in an endnote', wrap(60, (c, b) => {
  const id = c.note('endnotes', [para('<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>' + run(' The endnote holding the content.'), [el('pStyle', { val: 'EndnoteText' })]), ...b]);
  return [para(run('This paragraph has the endnote that holds the content.') + `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`)];
}));
def('header-content', 'where', 'where', 'the content in the page header', wrap(60, (c, b) => { (c.headers.default ||= []).push(...b); return []; }), { notWith: ['section-wrap'] });
def('footer-content', 'where', 'where', 'the content in the page footer', wrap(60, (c, b) => { (c.footers.default ||= []).push(...b); return []; }), { notWith: ['section-wrap'] });

// ---- section ----
const sp = (tag, attrs, inner) => (c) => c.sect.push(el(tag, attrs, inner));
def('landscape', 'section', 'page', 'landscape Letter pages', sp('pgSz', { w: 15840, h: 12240, orient: 'landscape' }));
def('a4', 'section', 'page', 'A4 pages', sp('pgSz', { w: 11906, h: 16838 }));
def('legal', 'section', 'page', 'Legal pages', sp('pgSz', { w: 12240, h: 20160 }));
def('margins-narrow', 'section', 'margins', '0.5in margins', sp('pgMar', { top: 720, right: 720, bottom: 720, left: 720 }));
def('margins-wide', 'section', 'margins', '1.5in side margins', sp('pgMar', { right: 2160, left: 2160 }));
def('margins-tall', 'section', 'margins', '2in top and bottom margins, header and footer 1in', sp('pgMar', { top: 2880, bottom: 2880, header: 1440, footer: 1440 }));
def('gutter', 'section', '', 'a 0.5in gutter', sp('pgMar', { gutter: 720 }));
def('mirror-margins', 'doc', '', 'mirrored margins', (c) => { c.settings.push(el('mirrorMargins')); sp('pgMar', { left: 1800, right: 1080 })(c); });
def('columns-2', 'section', 'cols', 'two columns', (c) => { c.cols = el('cols', { num: 2, space: 720 }); });
def('columns-3-unequal', 'section', 'cols', 'three unequal columns', (c) => { c.cols = el('cols', { num: 3, space: 360, equalWidth: 0 }, '<w:col w:w="2000" w:space="360"/><w:col w:w="3000" w:space="360"/><w:col w:w="3640"/>'); });
def('columns-2-line', 'section', 'cols', 'two columns with a line between', (c) => { c.cols = el('cols', { num: 2, sep: 1, space: 720 }); });
def('page-borders', 'section', '', 'a page border', sp('pgBorders', { offsetFrom: 'page' }, box4(12, 24)));
def('line-numbers', 'section', '', 'line numbers', sp('lnNumType', { countBy: 1, restart: 'continuous' }));
def('valign-center', 'section', '', 'text centred on the page vertically', sp('vAlign', { val: 'center' }));
def('doc-grid', 'section', '', 'a document grid of 18pt lines', sp('docGrid', { type: 'lines', linePitch: 360 }));
const hdr = (text) => para(run(text) + '<w:r><w:tab/></w:r><w:r><w:tab/></w:r>' + run('Right side'), [el('pStyle', { val: 'Header' })]);
const ftrPage = (text) => para(run(text + ' ') + fld('PAGE', '1'), [el('pStyle', { val: 'Footer' }), el('jc', { val: 'center' })]);
def('header-footer', 'section', '', 'a header and a page-number footer', (c) => { (c.headers.default ||= []).push(hdr('Document header')); (c.footers.default ||= []).push(ftrPage('Page')); });
def('title-page', 'section', '', 'a different first-page header and footer', (c) => { c.sect.push(el('titlePg')); c.headers.first = [hdr('First page header')]; c.footers.first = [ftrPage('First page')]; });
def('even-odd', 'doc', '', 'different even-page headers and footers', (c) => { c.settings.push(el('evenAndOddHeaders')); c.headers.even = [hdr('Even page header')]; c.footers.even = [ftrPage('Even page')]; c.headers.default ||= [hdr('Odd page header')]; });
def('page-number-roman', 'section', '', 'page numbers in roman numerals from iii', (c) => { c.sect.push(el('pgNumType', { fmt: 'lowerRoman', start: 3 })); c.needPageFooter = true; });
def('watermark', 'section', '', 'a text watermark', (c) => { (c.headers.default ||= []).unshift(VML_WATERMARK); });
def('columns-midpage', 'section', 'section-wrap', 'the content in a two-column continuous section between one-column text', wrap(70, (c, b) => [
  para(run('Before the columns: ' + c.t.sentences(2))), para('', [el('sectPr', {}, '{{SECT:a}}')]), ...b, para('', [el('sectPr', {}, '{{SECT:b}}')]),
]), { then: (c) => { c.specs.b = { cols: el('cols', { num: 2, space: 720 }), type: 'continuous' }; c.specs.a = {}; c.specs.final = { type: 'continuous' }; } });
def('landscape-section', 'section', 'section-wrap', 'the content in a landscape section between portrait pages', wrap(70, (c, b) => [
  para(run('Before the landscape section: ' + c.t.sentences(2))), para('', [el('sectPr', {}, '{{SECT:a}}')]), ...b, para('', [el('sectPr', {}, '{{SECT:b}}')]),
]), { then: (c) => { c.specs.a = {}; c.specs.b = { rotate: true, type: 'nextPage' }; c.specs.final = { type: 'nextPage' }; } });

// ---- doc ----
def('compat-2010', 'doc', 'compat', 'Word 2010 compatibility mode', (c) => { c.compat = 14; });
def('compat-2007', 'doc', 'compat', 'Word 2007 compatibility mode', (c) => { c.compat = 12; });
def('auto-hyphenation', 'doc', '', 'automatic hyphenation', (c) => c.settings.push(el('autoHyphenation')));
def('tab-default-half', 'doc', '', 'default tab stops every 0.25in', (c) => c.settings.push(el('defaultTabStop', { val: 360 })));
def('default-times-12', 'doc', 'defaults', 'Times New Roman 12pt as the default font', (c) => { c.styles = { defaultFont: 'Times New Roman', defaultSize: 24 }; });
def('default-arial-10', 'doc', 'defaults', 'Arial 10pt as the default font', (c) => { c.styles = { defaultFont: 'Arial', defaultSize: 20 }; });
def('length-long', 'doc', '', 'many paragraphs after the content (several pages)', (c) => { c.after = 24; });

export const FEATURES = F;
export const BY_ID = Object.fromEntries(F.map((f) => [f.id, f]));

/** Can these features go in one document? (No two from a group; nothing body-only where a `where` feature puts the content.) */
export function compatible(ids) {
  const fs = ids.map((id) => BY_ID[id]);
  const groups = new Set();
  for (const f of fs) { if (groups.has(f.group)) return false; groups.add(f.group); }
  const where = fs.some((f) => f.kind === 'where');
  for (const f of fs) {
    if (where && f.bodyOnly) return false;
    for (const n of f.notWith || []) if (fs.some((g) => g !== f && (g.id === n || g.group === n))) return false;
  }
  return true;
}

/** The parts of one document made of these features: what `pack` zips. */
export function compose(ids, { seed, title }) {
  const t = textSource(seed);
  let next = 1;
  const c = {
    t, rpr: [], rprAt: [], ppr: [], pprAt: [], runWraps: [], inlines: [], wraps: [], sect: [], settings: [], styles: {},
    headers: {}, footers: {}, footnotes: [], endnotes: [], comments: [], media: {}, specs: {}, after: 2, cols: el('cols', { space: 720 }),
    id: () => next++,
    note: (kind, blocks) => { const id = c[kind].length + 1; c[kind].push({ id, blocks }); return id; },
  };
  const fs = ids.map((id) => BY_ID[id]);
  for (const f of fs) { f.apply(c); f.then?.(c); }

  const contentPara = (i) => {
    const text = t.words(4);
    let feat = run(text, [...c.rpr, ...c.rprAt.flatMap((g) => g(i))], { del: c.featDel });
    for (const w of [...c.runWraps].sort((a, b) => a.order - b.order)) feat = w.fn(feat, i);
    const inner = run(`Paragraph ${i + 1}, ${t.words(2)}: `) + feat + c.inlines.map((g) => g(i)).join('') + run(' ' + t.sentences(2));
    const ppr = [...c.ppr, ...c.pprAt.flatMap((g) => g(i))];
    if (c.list && !ppr.some((p) => p.tag === 'pStyle')) ppr.unshift(el('pStyle', { val: 'ListParagraph' }));
    return para(inner, ppr);
  };
  let blocks = [0, 1, 2, 3].map(contentPara);
  for (const w of [...c.wraps].sort((a, b) => a.order - b.order)) blocks = w.fn(blocks);
  if (c.needPageFooter && !c.footers.default) c.footers.default = [ftrPage('Page')];

  const sectPr = (spec = {}) => {
    const kids = [
      ...Object.keys(c.headers).map((w) => el('headerReference', { type: w, 'r:id': `{{REF:header:${w}}}` })),
      ...Object.keys(c.footers).map((w) => el('footerReference', { type: w, 'r:id': `{{REF:footer:${w}}}` })),
      spec.type && el('type', { val: spec.type }),
      el('pgSz', { w: 12240, h: 15840 }), el('pgMar', { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 }),
      ...c.sect, spec.cols || c.cols,
    ];
    let xml = props('sectPr', kids, { wrap: false });
    if (spec.rotate) xml = xml.replace(/<w:pgSz w:w="(\d+)" w:h="(\d+)"(?: w:orient="\w+")?\/>/, (m, w, h) => `<w:pgSz w:w="${h}" w:h="${w}"${+h > +w ? ' w:orient="landscape"' : ''}/>`);
    return xml;
  };

  const after = Array.from({ length: c.after }, (_, k) => para(run(`After ${k + 1}: ` + t.sentences(3))));
  let body = [para(run(title)), ...blocks, ...after].join('') + `<w:sectPr>${sectPr(c.specs.final)}</w:sectPr>`;
  body = body.replace(/\{\{SECT:(\w+)\}\}/g, (_, k) => sectPr(c.specs[k]));
  const settings = [...c.settings];
  return { body, settings, styles: c.styles, compat: c.compat, headers: c.headers, footers: c.footers, footnotes: c.footnotes, endnotes: c.endnotes, comments: c.comments, media: c.media };
}
