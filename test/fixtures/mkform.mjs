// Builds test/fixtures/form.docx: Word-style content controls (text, checkbox, dropdown, date, block).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';
const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const ph = (t) => `<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>${t}</w:t></w:r>`;
const sdt = (tag, alias, pr, content, plc = true) => `<w:sdt><w:sdtPr><w:alias w:val="${alias}"/><w:tag w:val="${tag}"/><w:id w:val="${Math.floor(Math.random() * 1e8)}"/>${plc ? '<w:showingPlcHdr/>' : ''}${pr}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const P = (...x) => `<w:p>${x.join('')}</w:p>`;
const R = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const body = [
  P(R('Customer name: '), sdt('CustomerName', 'Customer name', '<w:text/>', ph('Click or tap here to enter text.'))),
  P(R('Agree to terms: '), sdt('Agree', 'Agree', `<w14:checkbox ${W14}><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>`, '<w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/></w:rPr><w:t>☐</w:t></w:r>', false)),
  P(R('Color: '), sdt('Color', 'Color', '<w:dropDownList><w:listItem w:displayText="Red" w:value="R"/><w:listItem w:displayText="Green" w:value="G"/><w:listItem w:displayText="Blue" w:value="B"/></w:dropDownList>', ph('Choose an item.'))),
  P(R('Date: '), sdt('Due', 'Due date', '<w:date><w:dateFormat w:val="M/d/yyyy"/><w:lid w:val="en-US"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date>', ph('Click or tap to enter a date.'))),
  sdt('Notes', 'Notes', '', P(ph('Click or tap here to enter text.'))),
  P(R('End of form.')),
].join('');
const files = unzipSync(blankDocxBytes());
const doc = strFromU8(files['word/document.xml']);
files['word/document.xml'] = strToU8(doc.replace(/<w:body>/, `<w:body>${body}`));
fs.writeFileSync(new URL('./form.docx', import.meta.url), zipSync(files));
