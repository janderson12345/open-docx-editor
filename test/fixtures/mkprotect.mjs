// Builds the Restrict Editing fixtures, each as Word writes it:
// - protect-readonly.docx: protected read-only, with a part everyone may edit and one only tester@example.com may;
// - protect-sections.docx: a protected form whose first section is left unprotected (w:formProt w:val="false");
// - protect-format.docx: formatting limited to the styles not locked (Heading 2 locked), editing otherwise free;
// - protect-final.docx: marked as final (docProps/custom.xml _MarkAsFinal).
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const para = (inner, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${inner}</w:p>`;
const MARGINS = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
const SETTINGS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';

function build(name, body, protection, { styles, custom } = {}) {
  const files = unzipSync(blankDocxBytes());
  files['word/document.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}<w:sectPr>${MARGINS}</w:sectPr></w:body></w:document>`);
  files['word/settings.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W}>${protection}<w:defaultTabStop w:val="720"/></w:settings>`);
  files['word/_rels/document.xml.rels'] = strToU8(strFromU8(files['word/_rels/document.xml.rels']).replace('</Relationships>', `<Relationship Id="rIdSet" Type="${SETTINGS_REL}" Target="settings.xml"/></Relationships>`));
  let types = strFromU8(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>');
  if (styles) files['word/styles.xml'] = strToU8(styles(strFromU8(files['word/styles.xml'])));
  if (custom) {
    files['docProps/custom.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="_MarkAsFinal"><vt:bool>true</vt:bool></property></Properties>`);
    files['_rels/.rels'] = strToU8(strFromU8(files['_rels/.rels']).replace('</Relationships>', '<Relationship Id="rIdCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>'));
    types = types.replace('</Types>', '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>');
  }
  files['[Content_Types].xml'] = strToU8(types);
  fs.writeFileSync(new URL(`./${name}`, import.meta.url), zipSync(files));
}

build('protect-readonly.docx', [
  para(run('This paragraph is locked.')),
  para(`${run('Before. ')}<w:permStart w:id="1" w:edGrp="everyone"/>${run('Everyone may edit this.')}<w:permEnd w:id="1"/>${run(' After.')}`),
  para(`<w:permStart w:id="2" w:ed="tester@example.com"/>${run('Only the tester may edit this.')}<w:permEnd w:id="2"/>`),
  para(run('This paragraph is locked too.')),
].join(''), '<w:documentProtection w:edit="readOnly" w:enforcement="1"/>');

const textField = (tag, value) => `<w:sdt><w:sdtPr><w:alias w:val="${tag}"/><w:tag w:val="${tag}"/><w:text/></w:sdtPr><w:sdtContent>${run(value)}</w:sdtContent></w:sdt>`;
build('protect-sections.docx', [
  para(run('This section is not protected: it can be edited freely.')),
  para('', `<w:sectPr>${MARGINS}<w:type w:val="continuous"/><w:formProt w:val="false"/></w:sectPr>`),
  para(run('Name: ') + textField('Name', 'Your name')),
  para(run('This text of the form is locked.')),
].join(''), '<w:documentProtection w:edit="forms" w:enforcement="1"/>');

build('protect-format.docx', [
  para(run('A heading in an allowed style'), '<w:pStyle w:val="Heading1"/>'),
  para(run('Some text that may be edited, but not given direct formatting.')),
].join(''), '<w:documentProtection w:formatting="1" w:enforcement="1"/>', {
  // (A built-in style is locked by its w:locked and its w:latentStyles entry both, as Word writes it.)
  styles: (xml) => xml.replace('<w:style ', '<w:latentStyles w:defLockedState="0" w:count="1"><w:lsdException w:name="heading 2" w:locked="1" w:uiPriority="9" w:qFormat="1"/></w:latentStyles><w:style ')
    .replace(/(<w:style\b[^>]*w:styleId="Heading2"[^>]*>[\s\S]*?<w:qFormat\/>)/, '$1<w:locked/>'),
});

build('protect-final.docx', para(run('A document marked as final.')), '', { custom: true });
