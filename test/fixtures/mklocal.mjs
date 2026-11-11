// Builds test/fixtures/localised-styles.docx: a document as a localised (Spanish) Word saves it. Its built-in styles keep
// Word's English names (w:name "heading 1", "toc 1", "TOC Heading") but have the language's own ids ("Ttulo1", "TDC1",
// "TtuloTDC"). Headings in Ttulo1 and Ttulo2, and a few paragraphs.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const R = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const P = (t, style) => `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}${R(t)}</w:p>`;
const files = unzipSync(blankDocxBytes());
let styles = strFromU8(files['word/styles.xml']);
for (const n of [1, 2, 3]) {
  styles = styles.replace(new RegExp(`w:styleId="Heading${n}"`, 'g'), `w:styleId="Ttulo${n}"`).replace(new RegExp(`w:val="Heading${n}"`, 'g'), `w:val="Ttulo${n}"`);
}
styles = styles.replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="TtuloTDC"><w:name w:val="TOC Heading"/><w:basedOn w:val="Ttulo1"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>'
  + [1, 2, 3].map((n) => `<w:style w:type="paragraph" w:styleId="TDC${n}"><w:name w:val="toc ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="100"/>${n > 1 ? `<w:ind w:left="${(n - 1) * 220}"/>` : ''}</w:pPr></w:style>`).join('') + '</w:styles>');
files['word/styles.xml'] = strToU8(styles);
const body = [P('Informe', null), P('Introducción', 'Ttulo1'), P('Texto de la introducción.'), P('Métodos', 'Ttulo1'), P('Detalle', 'Ttulo2'), P('Texto de los métodos.')].join('');
files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
fs.writeFileSync(new URL('./localised-styles.docx', import.meta.url), zipSync(files));
