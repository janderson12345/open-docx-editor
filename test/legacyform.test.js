import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { EditorState } from 'prosemirror-state';
import { DocxSession } from '../src/ooxml/docx.js';
import { blankDocxBytes } from '../src/ooxml/template.js';
import { listFields, fieldValues, setFieldValues } from '../src/editor/fields.js';
import { schema } from '../src/editor/schema.js';

// Legacy form fields (Developer > Legacy Tools), as Word saves them. Addr's result runs over two paragraphs, as
// Word writes a multi-line value.
const RPR = '<w:rPr><w:sz w:val="20"/></w:rPr>';
const fld = (ff, instr, result = '') => `<w:r>${RPR}<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="${ff[0]}"/><w:enabled/><w:calcOnExit w:val="0"/>${ff[1]}</w:ffData></w:fldChar></w:r>`
  + `<w:r>${RPR}<w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r>${RPR}<w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r>${RPR}<w:fldChar w:fldCharType="end"/></w:r>`;
const t = (s) => `<w:r>${RPR}<w:t xml:space="preserve">${s}</w:t></w:r>`;
const BODY = `<w:p><w:r><w:t>Name: </w:t></w:r>${fld(['Name', '<w:textInput><w:maxLength w:val="10"/></w:textInput>'], 'FORMTEXT', t('Ann'))}</w:p>`
  + `<w:p>${fld(['Agree', '<w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox>'], 'FORMCHECKBOX')}<w:r><w:t> Agree</w:t></w:r></w:p>`
  + `<w:p>${fld(['Pick', '<w:ddList><w:listEntry w:val="One"/><w:listEntry w:val="Two &amp; more"/></w:ddList>'], 'FORMDROPDOWN')}</w:p>`
  + `<w:p>${fld(['Addr', '<w:textInput/>'], 'FORMTEXT', t('Line 1') + '</w:p><w:p>' + t('Line 2'))}</w:p>`
  + `<w:p>${fld(['Blank', '<w:textInput/>'], 'FORMTEXT', t('     '))}</w:p>`;

function open(body, settings) {
  const files = unzipSync(blankDocxBytes());
  const main = new DocxSession(blankDocxBytes()).pkg.mainPath;
  files[main] = strToU8(strFromU8(files[main]).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`));
  if (settings) {
    // The blank template has no settings part: add one.
    files['word/settings.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${settings}</w:settings>`);
    const rels = 'word/_rels/document.xml.rels';
    files[rels] = strToU8(strFromU8(files[rels]).replace('</Relationships>', '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'));
    const ct = '[Content_Types].xml';
    files[ct] = strToU8(strFromU8(files[ct]).replace('</Types>', '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'));
  }
  return new DocxSession(zipSync(files));
}

test('legacy form fields are listed with their values, and check boxes and drop-downs are drawn', () => {
  const s = open(BODY);
  assert.deepEqual(fieldValues(s.doc), { Name: 'Ann', Agree: false, Pick: 'One', Addr: 'Line 1\nLine 2', Blank: '' });
  assert.deepEqual(listFields(s.doc).find((f) => f.tag === 'Pick').items.map((i) => i.text), ['One', 'Two & more']);
  const shown = [];
  // (A check box or drop-down is drawn from its begin fldChar's ffData.)
  s.doc.descendants((n) => { if (n.attrs.kind === 'fldChar') { const d = schema.nodes.opaque_inline.spec.toDOM(n); if (d[2]) shown.push(d[2]); } });
  assert.deepEqual(shown, ['☐', 'One']);
});

test('legacy form fields are filled in and saved as Word expects', () => {
  const s = open(BODY);
  const view = { state: EditorState.create({ doc: s.doc }), dispatch(tr) { this.state = this.state.apply(tr); } };
  assert.deepEqual(setFieldValues(view, { Name: 'Bartholome', Agree: true, Pick: 'Two & more', Addr: 'A\nB', Blank: 'x' }), []);
  assert.deepEqual(fieldValues(view.state.doc), { Name: 'Bartholome', Agree: true, Pick: 'Two & more', Addr: 'A\nB', Blank: 'x' });
  const out = strFromU8(unzipSync(s.save(view.state.doc))[s.pkg.mainPath]);
  assert.match(out, /<w:checkBox><w:sizeAuto\/><w:default w:val="0"\/><w:checked w:val="1"\/><\/w:checkBox>/);
  assert.match(out, /<w:ddList><w:result w:val="1"\/><w:listEntry/);
  // The new result keeps the field's run formatting.
  assert.match(out, /fldCharType="separate"\/><\/w:r><w:r><w:rPr><w:sz w:val="20"\/><\/w:rPr><w:t[^>]*>Bartholome<\/w:t>/);
  const again = open(out.replace(/^[\s\S]*<w:body>|<w:sectPr[\s\S]*$/g, ''));
  assert.deepEqual(fieldValues(again.doc), fieldValues(view.state.doc));
  // Addr's two lines are two paragraphs, the field ending in the second.
  assert.match(out, /<w:t[^>]*>A<\/w:t><\/w:r><\/w:p><w:p>(?:(?!<\/w:p>).)*<w:t[^>]*>B<\/w:t><\/w:r><w:r>(?:(?!<\/w:r>).)*fldCharType="end"/);
});

test('a one-line value written into a two-paragraph text field joins it into one paragraph', () => {
  const s = open(BODY);
  const view = { state: EditorState.create({ doc: s.doc }), dispatch(tr) { this.state = this.state.apply(tr); } };
  const paras = view.state.doc.childCount;
  setFieldValues(view, { Addr: 'One line' });
  assert.equal(view.state.doc.childCount, paras - 1);
  assert.equal(fieldValues(view.state.doc).Addr, 'One line');
  setFieldValues(view, { Addr: 'x\ny\nz' });
  assert.equal(view.state.doc.childCount, paras + 1);
  assert.equal(fieldValues(view.state.doc).Addr, 'x\ny\nz');
  assert.equal(fieldValues(view.state.doc).Blank, '');
});

test('a document protected for filling in forms says so', () => {
  assert.equal(open(BODY, '<w:documentProtection w:edit="forms" w:enforcement="1"/>').formsProtected, true);
  assert.equal(open(BODY, '<w:documentProtection w:edit="readOnly" w:enforcement="1"/>').formsProtected, false);
  assert.equal(open(BODY).formsProtected, false);
});
