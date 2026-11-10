// Builds test/fixtures/rtl-hebrew.docx and rtl-arabic.docx: right-to-left paragraphs (w:bidi) in each alignment, with
// indents and a first-line indent (which mirror), English and numbers inside right-to-left text, a tab stop, a
// right-to-left table (w:bidiVisual), and a left-to-right paragraph for contrast; long enough to run onto a second page.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;');
// (In the schema's order: w:rtl after the sizes.)
const rpr = (rtl) => `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/>${rtl ? '<w:rtl/>' : ''}</w:rPr>`;
const run = (t, rtl = true) => `<w:r>${rpr(rtl)}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
const tab = () => `<w:r>${rpr(true)}<w:tab/></w:r>`;
const para = (runs, { bidi = true, jc = null, ind = null, tabs = null, style = null } = {}) =>
  `<w:p><w:pPr>${style ? `<w:pStyle w:val="${style}"/>` : ''}${tabs || ''}${bidi ? '<w:bidi/>' : ''}${ind ? `<w:ind ${ind}/>` : ''}${jc ? `<w:jc w:val="${jc}"/>` : ''}</w:pPr>${runs}</w:p>`; // (schema order: tabs, bidi, ind, jc)
const cell = (t, w) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${para(run(t))}</w:tc>`;
const table = (rows) => `<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('')}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${rows.map((r) => `<w:tr>${r.map((t) => cell(t, 3000)).join('')}</w:tr>`).join('')}</w:tbl>`;

function build(name, t) {
  const body = [
    para(run(t.title), { style: 'Heading1' }),
    para(run(t.long)),
    para(run(t.long), { jc: 'both' }),
    para(run(t.short), { jc: 'center' }),
    para(run(t.short), { jc: 'left' }),
    para(run(t.short), { jc: 'right' }),
    para(run(t.long), { ind: 'w:left="1440" w:right="720"' }),
    para(run(t.long), { ind: 'w:firstLine="720"' }),
    para(run(t.long), { ind: 'w:left="720" w:hanging="720"' }),
    para(run(t.mixedA) + run('Microsoft Word 2016', false) + run(t.mixedB) + run(' 12,345.67 ') + run(t.short)),
    para(run(t.label) + tab() + run('42'), { tabs: '<w:tabs><w:tab w:val="left" w:pos="4320"/></w:tabs>' }),
    table([[t.c1, t.c2, t.c3], [t.long.slice(0, 60), '1,024', t.short]]),
    para(run('A left-to-right paragraph in English, for contrast: it starts at the left margin and wraps to the right. '.repeat(2), false), { bidi: false }),
    ...Array.from({ length: 8 }, (_, i) => para(run(`${i + 1}. ${t.long}`), { jc: i % 2 ? 'both' : null })),
  ].join('');
  const files = unzipSync(blankDocxBytes());
  const doc = strFromU8(files['word/document.xml']);
  files['word/document.xml'] = strToU8(doc.replace(/<w:body>/, `<w:body>${body}`));
  fs.writeFileSync(new URL(`./${name}.docx`, import.meta.url), zipSync(files));
}

build('rtl-hebrew', {
  title: 'דוח רבעוני',
  short: 'שורה קצרה בעברית.',
  long: 'זהו משפט ארוך בעברית שנועד לבדוק כיצד הטקסט נשבר לשורות כאשר הפסקה נכתבת מימין לשמאל, עם מילים רבות ככל האפשר כדי שהשורה תתמלא ותעבור לשורה הבאה. ',
  mixedA: 'הגרסה ',
  mixedB: ' שוחררה בשנה שעברה, והמחיר הוא',
  label: 'סכום כולל',
  c1: 'שם', c2: 'כמות', c3: 'הערות',
});
build('rtl-arabic', {
  title: 'التقرير الربعي',
  short: 'سطر قصير باللغة العربية.',
  long: 'هذه جملة طويلة باللغة العربية تهدف إلى اختبار كيفية انقسام النص إلى أسطر عندما تكتب الفقرة من اليمين إلى اليسار، مع أكبر عدد ممكن من الكلمات حتى يمتلئ السطر وينتقل إلى السطر التالي. ',
  mixedA: 'الإصدار ',
  mixedB: ' صدر في العام الماضي، والسعر هو',
  label: 'المجموع الكلي',
  c1: 'الاسم', c2: 'الكمية', c3: 'ملاحظات',
});
