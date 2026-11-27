// Builds test/fixtures/char-scale.docx: text of a character scale (w:w, 110%) in one run and in runs of one to five
// letters, as Word sets their widths (its prints of the same runs, made here and printed by Word: 474.55, 477.53 and
// 482.97px for the three lines of 11pt Times New Roman).
// With a folder (node test/fixtures/mkcharscale.mjs <dir>) it writes every variant tried in Word there instead.
import '../setup.js';
import fs from 'fs';
import path from 'path';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const out = process.argv[2] ? path.resolve(process.argv[2]) : null;
if (out) fs.mkdirSync(out, { recursive: true });

function write(name, body) {
  const files = unzipSync(blankDocxBytes());
  const main = 'word/document.xml';
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
  let st = strFromU8(files['word/styles.xml']);
  st = st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, '$1<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr>$2');
  files['word/styles.xml'] = strToU8(st);
  fs.writeFileSync(out ? path.join(out, name) : new URL(`./${name}`, import.meta.url), zipSync(files));
}

const TEXT = 'Name and Sample ID of explanation that represented units';
const p = (font, hp, w) => `<w:p><w:r><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/>${w !== 100 ? `<w:w w:val="${w}"/>` : ''}<w:sz w:val="${hp}"/></w:rPr><w:t xml:space="preserve">${font.split(' ')[0]} ${hp} ${w}: ${TEXT}</w:t></w:r></w:p>`;
const all = (font) => [20, 21, 22, 23, 24].flatMap((hp) => [100, 90, 105, 110, 150].map((w) => p(font, hp, w))).join('');
// The same text in runs of one to five letters, alternating colours (as a converted PDF's form has it), each scaled,
// and a space with its own character spacing.
const pieces = (hp, w, sp = 0) => {
  const t = `Split ${hp} ${w} ${sp}: ${TEXT}`; const out = []; let i = 0, k = 0;
  while (i < t.length) { const n = 1 + (k * 7) % 5; out.push(t.slice(i, i + n)); i += n; k++; }
  return `<w:p>${out.map((x, j) => `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:color w:val="${['2F2838', '010101', '211828'][j % 3]}"/>${sp && x === ' ' ? `<w:spacing w:val="${sp}"/>` : ''}${w !== 100 ? `<w:w w:val="${w}"/>` : ''}<w:sz w:val="${hp}"/></w:rPr><w:t xml:space="preserve">${x}</w:t></w:r>`).join('')}</w:p>`;
};
const split = [21, 22, 23].flatMap((hp) => [p('Times New Roman', hp, 110), pieces(hp, 110), pieces(hp, 110, -7), pieces(hp, 100)]).join('');
// Forty of one character a line, at 100% and 110%: Word's advance for each.
const CH = 'aceimnorstuyNSACDI ';
const reps = (hp) => [...CH].flatMap((c) => [100, 110].map((w) => `<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>${w !== 100 ? `<w:w w:val="${w}"/>` : ''}<w:sz w:val="${hp}"/></w:rPr><w:t xml:space="preserve">|${c.repeat(40)}|</w:t></w:r></w:p>`)).join('');
const variants = { 'w01-times': all('Times New Roman'), 'w02-arial': all('Arial'), 'w03-calibri': all('Calibri'), 'w04-split': split, 'w05-reps22': reps(22), 'w06-reps20': reps(20),
  // Bold Arial 10pt, a line of a box that Word wraps a word earlier than its advances would (1113 page 60's codes box).
  'w07-arialbold': ['A -  FIRST SAMPLE ITEM C', 'A -  FIRST SAMPLE ITEM', 'C -  LINE NUMBER SEQUENCE D -'].map((t) => `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`).join('') };
if (out) {
  for (const [n, b] of Object.entries(variants)) write(`${n}.docx`, b);
  console.log(Object.keys(variants).length, 'documents in', out);
} else {
  write('char-scale.docx', p('Times New Roman', 22, 110) + pieces(22, 110) + pieces(22, 110, -7));
}
