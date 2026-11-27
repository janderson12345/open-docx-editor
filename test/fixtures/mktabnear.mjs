// Builds test/fixtures/tab-near-stop.docx: a hanging paragraph "Label<tab><tab>Text" whose own tab stop lies 0 to 10 twips
// past the hanging indent, as Word sets it (its prints of the same paragraphs, made here and printed by Word): the
// second tab goes to that stop however near, not to the next default one.
// With a folder (node test/fixtures/mktabnear.mjs <dir>) it writes every variant tried in Word there instead.
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
  st = st.replace(/(<w:style w:type="paragraph" w:default="1" w:styleId="Normal">[\s\S]*?)(<\/w:style>)/, '$1<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/></w:rPr>$2');
  files['word/styles.xml'] = strToU8(st);
  fs.writeFileSync(out ? path.join(out, name) : new URL(`./${name}`, import.meta.url), zipSync(files));
}

const r = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const tab = '<w:r><w:tab/></w:r>';
// (1113 page 3: left 3020, hanging 2828, a stop at 3021.)
const para = (stop, label) => `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="${stop}"/></w:tabs><w:ind w:left="3020" w:right="681" w:hanging="2828"/></w:pPr>${r(label)}${tab}${tab}${r(`Text after two tabs, stop ${stop}`)}</w:p>`;
const body = [3020, 3021, 3025, 3030, 3040].map((s, k) => para(s, `Label${k}`)).join('');
if (out) {
  write('n01-near.docx', body);
  console.log('1 document in', out);
} else {
  write('tab-near-stop.docx', body);
}
