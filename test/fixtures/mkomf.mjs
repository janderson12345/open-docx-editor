// Builds test/fixtures/omf-fonts.docx: text in Microsoft fonts the editor draws in open-metric-fonts stand-ins where the
// computer lacks them (Grandview, Seaford, Aptos SemiBold), for test/csp.mjs (loaded from the host's own copy).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const P = (font, t) => `<w:p><w:r><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const body = [P('Grandview', 'Grandview: the quick brown fox jumps over the lazy dog.'), P('Seaford', 'Seaford: the quick brown fox.'), P('Aptos SemiBold', 'Aptos SemiBold: the quick brown fox.')].join('');
const files = unzipSync(blankDocxBytes());
files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/<w:body>/, `<w:body>${body}`));
fs.writeFileSync(new URL('./omf-fonts.docx', import.meta.url), zipSync(files));
