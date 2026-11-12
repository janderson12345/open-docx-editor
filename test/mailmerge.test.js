import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { mailMerge, mergeFieldNames, formatMergeValue } from '../src/index.js';

const bytes = new Uint8Array(fs.readFileSync(new URL('./fixtures/mailmerge.docx', import.meta.url)));
const records = [
  { FirstName: 'Ada', 'Last Name': 'Lovelace', City: 'London', Title: 'Countess', Amount: '1234.5', Due: '2026-12-01', Account: 'A-1' },
  { firstname: 'Alan', LastName: 'Turing', City: 'Wilmslow', Title: '', Amount: 7, Due: '2027-01-15', Account: 'B-2' },
];
const docXml = (b) => strFromU8(unzipSync(b)['word/document.xml']);
const text = (xml) => (xml.match(/<w:t(?: [^>]*)?>[^<]*<\/w:t>/g) || []).map((t) => t.replace(/<[^>]+>/g, '')).join('');

test('merge field names, each once, in order', () => {
  assert.deepEqual(mergeFieldNames(bytes), ['FirstName', 'Last Name', 'City', 'Title', 'Amount', 'Due', 'Account']);
});

test('field switches as Word formats them', () => {
  assert.equal(formatMergeValue('lovelace', ' MERGEFIELD X \\* Upper '), 'LOVELACE');
  assert.equal(formatMergeValue('ada lovelace', ' MERGEFIELD X \\* Caps '), 'Ada Lovelace');
  assert.equal(formatMergeValue('1234.5', ' MERGEFIELD X \\# "$#,##0.00" '), '$1,234.50');
  assert.equal(formatMergeValue('7', ' MERGEFIELD X \\# "0.0" '), '7.0');
  assert.equal(formatMergeValue('2026-12-01T00:00', ' MERGEFIELD X \\@ "MMMM d, yyyy" '), 'December 1, 2026');
  assert.equal(formatMergeValue('Sir', ' MERGEFIELD X \\b "Dear " \\f "," '), 'Dear Sir,');
  assert.equal(formatMergeValue('', ' MERGEFIELD X \\b "Dear " \\f "," '), '');
});

test('one document of all records, each on a new page, fields as plain text in their formatting', () => {
  const { bytes: out, fields, missing } = mailMerge(bytes, records);
  assert.equal(fields.length, 7);
  assert.deepEqual(missing, []);
  const xml = docXml(out);
  assert.doesNotMatch(xml, /MERGEFIELD|fldChar|fldSimple|«/);
  const t = text(xml);
  assert.match(t, /Ada LOVELACE/);
  assert.match(t, /London, record 1/);
  assert.match(t, /Dear Countess, end of greeting/);
  assert.match(t, /Balance due: \$1,234\.50 by December 1, 2026\./);
  assert.match(t, /Account: A-1/);
  // Field names matched ignoring case and spaces; an empty value drops its \b and \f text.
  assert.match(t, /Alan TURING/);
  assert.match(t, /Wilmslow, record 2/);
  assert.match(t, /Wilmslow, record 2 end of greeting/);
  assert.match(t, /Balance due: \$7\.00 by January 15, 2027\./);
  // The bold field stays bold, the simple field's italic run stays italic.
  assert.match(xml, /<w:r><w:rPr><w:b\/><\/w:rPr><w:t[^>]*>LOVELACE<\/w:t>/);
  assert.match(xml, /<w:r><w:rPr><w:i\/><\/w:rPr><w:t[^>]*>London<\/w:t>/);
  // Record 1 ends with a next-page section break; the document's own last section stays last.
  assert.equal((xml.match(/<w:sectPr\b/g) || []).length, 2);
  assert.match(xml, /Yours sincerely\.<\/w:t><\/w:r><\/w:p>|<w:type w:val="nextPage"\/>/);
  assert.equal((xml.match(/<w:type w:val="nextPage"\/>/g) || []).length, 1);
});

test('one document per record', () => {
  const { files } = mailMerge(bytes, records, { combine: false });
  assert.equal(files.length, 2);
  assert.match(text(docXml(files[1])), /Alan TURING/);
  assert.doesNotMatch(text(docXml(files[1])), /Ada/);
});
