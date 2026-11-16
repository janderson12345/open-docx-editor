import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { mailMerge, mergeFieldNames, formatMergeValue, ifCompare } from '../src/index.js';
import { mergeRecord, mergeCursor } from '../src/merge.js';
import { mailMergePlugin, previewStory } from '../src/editor/mailmerge.js';
import { DocxSession } from '../src/ooxml/docx.js';
import { EditorState } from 'prosemirror-state';

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

// IF fields and merge fields in a header and footer (test/fixtures/mkmergeif.mjs), against Word's own merge of the same
// document with these records (Mailings > Finish & Merge through Word, a CSV data source): each line's text per record.
const ifBytes = new Uint8Array(fs.readFileSync(new URL('./fixtures/mergeif.docx', import.meta.url)));
const people = [
  { Name: 'Ada', Amount: '150', City: 'New York' },
  { Name: 'alan', Amount: '50', City: 'Boston' },
  { Name: 'Bea', Amount: '100', City: 'New York' },
  { Name: 'Zed', Amount: '', City: 'London' },
  { Name: 'Ida', Amount: '45', City: '' },
];
const LABELS = 'gt lt ge le ne eqcase star qmark starcase strlt strltcase quoted bare numstr text nested nofalse upper rec'.split(' ');
const WORD = [
  'big|not under|yes|gt|has amount|ne|A-name|q|no a-star|before B|before b|NY|NY|not hundred|Dear Ada, thanks|big|only big|BIG|not second',
  'small|under|no|le|has amount|ne|other|nq|a-star|before B|before b|not NY|not NY|not hundred|Hello alan|mid||SMALL|second',
  'small|not under|yes|gt|has amount|ne|other|nq|no a-star|from B|from b|NY|NY|hundred|Hello Bea|mid||SMALL|not second',
  'small|under|no|le|no amount|ne|other|nq|no a-star|from B|from b|not NY|not NY|not hundred|Hello Zed|tiny||SMALL|not second',
  'small|under|no|le|has amount|ne|other|q|no a-star|from B|from b|not NY|not NY|not hundred|Hello Ida|mid||SMALL|not second',
];
const WORD_HF = [['Statement for Ada', 'City: New York / big'], ['Statement for alan', 'City: Boston / small'], ['Statement for Bea', 'City: New York / small'],
  ['Statement for Zed', 'City: London / small'], ['Statement for Ida', 'City:  / small']];
const paras = (xml) => xml.split('</w:p>').map(text).filter(Boolean);

test('IF fields per record, as Word merges them', () => {
  const { bytes: out } = mailMerge(ifBytes, people);
  const xml = docXml(out);
  assert.doesNotMatch(xml, /fldChar|instrText|fldSimple/);
  const lines = paras(xml);
  assert.equal(lines.length, LABELS.length * people.length);
  people.forEach((_, r) => WORD[r].split('|').forEach((want, k) => assert.equal(lines[r * LABELS.length + k], `${LABELS[k]}: ${want}`, `record ${r + 1}, ${LABELS[k]}`)));
});

test('the comparisons Word\'s IF makes', () => {
  assert.equal(ifCompare('1,234.5', '>', '999'), true);
  assert.equal(ifCompare('100', '=', '100.0'), true);
  assert.equal(ifCompare('Ada', '=', 'ada'), false);
  assert.equal(ifCompare('Ada', '<>', 'A*'), false);
  assert.equal(ifCompare('alan', '<', 'B'), true);
  assert.equal(ifCompare('a.b', '=', 'a?b'), true);
  assert.equal(ifCompare('axb', '=', 'a.b'), false);
});

test('headers and footers filled for each record, each record its own', () => {
  const { bytes: out, fields } = mailMerge(ifBytes, people);
  assert.deepEqual(fields, ['Amount', 'Name', 'City']);
  const files = unzipSync(out);
  const rels = strFromU8(files['word/_rels/document.xml.rels']);
  const part = (id) => strFromU8(files[`word/${new RegExp(`Id="${id}"[^>]*Target="([^"]+)"`).exec(rels)[1]}`]);
  const sects = docXml(out).match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g);
  assert.equal(sects.length, people.length);
  sects.forEach((s, r) => {
    const h = /headerReference w:type="default" r:id="([^"]+)"/.exec(s)[1];
    const f = /footerReference w:type="default" r:id="([^"]+)"/.exec(s)[1];
    assert.deepEqual([text(part(h)), text(part(f))], WORD_HF[r], `record ${r + 1}`);
    assert.doesNotMatch(part(h) + part(f), /fldChar|instrText/);
  });
  assert.match(strFromU8(files['[Content_Types].xml']), /\/word\/header5\.xml/);
  // One document per record: its own header filled in place.
  const one = mailMerge(ifBytes, people, { combine: false }).files[2];
  assert.equal(text(strFromU8(unzipSync(one)['word/header1.xml'])), 'Statement for Bea');
});

// NEXT, NEXTIF and SKIPIF (test/fixtures/mkmergenext.mjs), against Word's own merge of the same documents with these
// records (Finish & Merge through Word, a CSV data source) and its Preview Results at each record: each document's
// lines, its header (the record it ended on), MERGEREC the record's number, MERGESEQ the merged document's.
const group = [['Ada', 'x', '5'], ['Bob', 'x', '20'], ['Cy', 'y', '30'], ['Di', 'x', '2'], ['Ed', 'y', '40'], ['Flo', 'x', '50'], ['Gus', 'y', '60']]
  .map(([Name, Group, Amount]) => ({ Name, Group, Amount }));
const fixture = (name) => new Uint8Array(fs.readFileSync(new URL(`./fixtures/${name}.docx`, import.meta.url)));
const norm = (s) => s.replace(/\s+/g, ' ').trim();
const WORD_NEXT = {
  mergenext: [
    ['a: 1 1 Ada|n1: .|b: 2 1 Bob|n2: .|c: 3 1 Cy|if: big', 'H: Cy'],
    ['a: 4 2 Di|n1: .|b: 5 2 Ed|n2: .|c: 6 2 Flo|if: big', 'H: Flo'],
    // (NEXT past the last record: every field after it empty, MERGEREC and MERGESEQ too.)
    ['a: 7 3 Gus|n1: .|b: 7 3 Gus|n2: .|c:|if: small', 'H:'],
  ],
  // (Ada and Di left out by the first SKIPIF, Ed by the second after its NEXTIF stayed on it.)
  mergeskip: [['a: 2 1 Bob|s1: .|b: 2 1 Bob|n: .|c: 3 1 Cy|s2: .|d: 3 1 Cy', 'H: Cy'], ['a: 6 2 Flo|s1: .|b: 6 2 Flo|n: .|c: 7 2 Gus|s2: .|d: 7 2 Gus', 'H: Gus']],
  // (Ada's document, moved on to Bob, left out: the next starts after Bob.)
  mergeskip2: [['a: 3 1 Cy|n: .|b: 3 1 Cy|s: .', 'H: Cy'], ['a: 4 2 Di|n: .|b: 5 2 Ed|s: .', 'H: Ed'], ['a: 6 3 Flo|n: .|b: 7 3 Gus|s: .', 'H: Gus']],
};

test('NEXT, NEXTIF and SKIPIF as Word merges them', () => {
  for (const [name, want] of Object.entries(WORD_NEXT)) {
    const { bytes: out } = mailMerge(fixture(name), group);
    const files = unzipSync(out);
    const xml = docXml(out);
    assert.doesNotMatch(xml, /fldChar|instrText|fldSimple/, name);
    const rels = strFromU8(files['word/_rels/document.xml.rels']);
    const part = (id) => strFromU8(files[`word/${new RegExp(`Id="${id}"[^>]*Target="([^"]+)"`).exec(rels)[1]}`]);
    // Each document's paragraphs, up to and with the one holding its section break.
    const docs = [[]];
    for (const p of xml.split('</w:p>')) {
      if (norm(text(p))) docs[docs.length - 1].push(norm(text(p)));
      if (/<w:sectPr\b/.test(p)) docs.push([]);
    }
    docs.splice(-1, 1);
    for (let i = 0; i < docs.length; i++) docs[i] = docs[i].join('|');
    const heads = [...xml.matchAll(/headerReference w:type="default" r:id="([^"]+)"/g)].map((m) => norm(text(part(m[1]))));
    assert.deepEqual(docs.map((d, i) => [d, heads[i]]), want, name);
    // One document per merged document.
    assert.equal(mailMerge(fixture(name), group, { combine: false }).files.length, want.length, name);
  }
});

const WORD_PREVIEW = {
  mergenext: [
    ['a: 1 Ada|n1: .|b: 2 Bob|n2: .|c: 3 Cy|if: big', 'H: Cy'], ['a: 2 Bob|n1: .|b: 3 Cy|n2: .|c: 4 Di|if: small', 'H: Di'],
    ['a: 3 Cy|n1: .|b: 3 Cy|n2: .|c: 4 Di|if: small', 'H: Di'], ['a: 4 Di|n1: .|b: 5 Ed|n2: .|c: 6 Flo|if: big', 'H: Flo'],
    ['a: 5 Ed|n1: .|b: 5 Ed|n2: .|c: 6 Flo|if: big', 'H: Flo'], ['a: 6 Flo|n1: .|b: 7 Gus|n2: .|c:|if: small', 'H:'],
    ['a: 7 Gus|n1: .|b: 7 Gus|n2: .|c:|if: small', 'H:'],
  ],
  // (SKIPIF has no effect in the preview: Ada's record shown.)
  mergeskip: [
    ['a: 1 Ada|s1: .|b: 1 Ada|n: .|c: 2 Bob|s2: .|d: 2 Bob', 'H: Bob'], ['a: 2 Bob|s1: .|b: 2 Bob|n: .|c: 3 Cy|s2: .|d: 3 Cy', 'H: Cy'],
    ['a: 3 Cy|s1: .|b: 3 Cy|n: .|c: 3 Cy|s2: .|d: 3 Cy', 'H: Cy'], ['a: 4 Di|s1: .|b: 4 Di|n: .|c: 5 Ed|s2: .|d: 5 Ed', 'H: Ed'],
    ['a: 5 Ed|s1: .|b: 5 Ed|n: .|c: 5 Ed|s2: .|d: 5 Ed', 'H: Ed'], ['a: 6 Flo|s1: .|b: 6 Flo|n: .|c: 7 Gus|s2: .|d: 7 Gus', 'H: Gus'],
    ['a: 7 Gus|s1: .|b: 7 Gus|n: .|c: 7 Gus|s2: .|d: 7 Gus', 'H: Gus'],
  ],
};

test('Preview Results with NEXT, NEXTIF and SKIPIF, as Word previews them', () => {
  for (const [name, want] of Object.entries(WORD_PREVIEW)) {
    const session = new DocxSession(fixture(name));
    const got = group.map((_, i) => {
      // (What the preview's decorations show: mailmerge.js previewDecos, the plugin run on a document change.)
      session.merge = { records: group, fields: [], index: i, preview: true };
      const state = EditorState.create({ doc: session.doc, plugins: [mailMergePlugin(session)] });
      state.apply(state.tr.insertText(' ', 1));
      const body = mergeRecord(session.doc, mergeCursor(group, i, { preview: true }));
      const lines = [];
      body.forEach((p) => lines.push(norm(p.textContent)));
      return [lines.filter(Boolean).join('|'), norm(previewStory(session, session.story('rIdMmH')).textContent)];
    });
    assert.deepEqual(got, want, name);
  }
});
