import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import mammoth from 'mammoth';
import { unzipSync, strFromU8 } from 'fflate';
import { openDocument } from '../src/agent.js';

const load = (f) => new Uint8Array(fs.readFileSync(new URL(`./fixtures/${f}`, import.meta.url)));
const docXml = (bytes) => strFromU8(unzipSync(bytes)['word/document.xml']);

test('reading: paragraphs with ids, the outline, list numbers, tables, and a Markdown view', () => {
  const d = openDocument(load('fields.docx'));
  const ps = d.paragraphs();
  assert.equal(ps[0].id, 'p0');
  assert.equal(ps[0].style, 'Title');
  assert.deepEqual(d.outline().map((h) => `${h.level} ${h.text}`), ['1 Introduction', '1 Methods of measurement', '1 Results']);
  const methods = ps.find((p) => p.text === 'Methods of measurement');
  assert.equal(methods.list, '2.');
  const md = d.markdown();
  assert.match(md, /<!-- p\d+ -->\n# 2\. Methods of measurement/);
  const rich = openDocument(load('rich.docx'));
  const cell = rich.paragraphs().find((p) => p.table);
  assert.ok(cell && cell.table.table === 0);
  assert.match(rich.markdown(), /^\| .* \|$/m);
});

test('search finds text with its paragraph and place', () => {
  const d = openDocument(load('fields.docx'));
  const hits = d.search('quick brown fox');
  assert.ok(hits.length > 20);
  assert.equal(hits[0].match, 'quick brown fox');
  assert.equal(hits[0].text.slice(hits[0].start, hits[0].end), 'quick brown fox');
  assert.equal(d.search('(\\d+)\\. The quick', { regex: true })[0].match, '1. The quick');
});

test('edits are tracked changes under the agent\'s name, and Word reads them', async () => {
  const d = openDocument(load('fields.docx'), { author: 'Assistant' });
  const intro = d.outline()[0].id;
  assert.equal(d.replace('Introduction', 'Overview', { id: intro }), 1);
  const p = d.paragraphs().find((x) => x.text.startsWith('Front matter paragraph 1.'));
  d.setParagraphText(p.id, 'Front matter paragraph 1. A shorter, plainer paragraph.');
  const added = d.insertParagraph(intro, 'A new opening sentence.', { style: 'Normal' });
  d.deleteParagraph(d.paragraphs().find((x) => x.text.startsWith('Front matter paragraph 2.')).id);
  const changes = d.changes();
  assert.ok(changes.length >= 4 && changes.every((c) => c.author === 'Assistant'));
  assert.ok(changes.some((c) => c.kind === 'deletion' && c.text === 'Introduction'));
  assert.ok(changes.some((c) => c.kind === 'insertion' && c.text === 'Overview'));
  // The rewritten paragraph changes only what differs.
  assert.ok(changes.some((c) => c.kind === 'insertion' && /shorter, plainer/.test(c.text)));
  assert.equal(d.paragraphs().find((x) => x.id === added).text, 'A new opening sentence.');
  const bytes = d.save();
  const xml = docXml(bytes);
  assert.match(xml, /<w:ins [^>]*w:author="Assistant"/);
  assert.match(xml, /<w:del [^>]*w:author="Assistant"/);
  // An independent reader sees the document as it reads with the changes accepted.
  const text = (await mammoth.extractRawText({ buffer: Buffer.from(bytes) })).value;
  assert.match(text, /Overview/);
  // Accepting them: the text as edited, no revisions left.
  const again = openDocument(bytes);
  again.resolveChanges({ accept: true });
  assert.equal(again.changes().length, 0);
  assert.ok(again.paragraphs().some((x) => x.text === 'Overview'));
  assert.ok(!again.paragraphs().some((x) => x.text.startsWith('Front matter paragraph 2.')));
});

test('untracked edits, comments and replies, form fields', () => {
  const d = openDocument(load('fields.docx'));
  d.replace('Results', 'Findings', { all: true, tracked: false });
  assert.equal(d.changes().length, 0);
  assert.ok(d.outline().some((h) => h.text === 'Findings'));
  const target = d.search('lazy dog')[0];
  const id = d.addComment(target.id, 'Is this example needed?', { quote: 'lazy dog' });
  const reply = d.reply(id, 'Yes, keep it.', { author: 'Reviewer' });
  const comments = d.comments();
  assert.equal(comments.find((c) => c.id === id).quote, 'lazy dog');
  assert.equal(comments.find((c) => c.id === reply).replyTo, id);
  const xml = strFromU8(unzipSync(d.save())['word/comments.xml']);
  assert.match(xml, /Is this example needed\?/);
  const form = openDocument(load('form.docx'));
  const names = Object.keys(form.fields());
  assert.ok(names.length > 0);
  assert.deepEqual(form.setFields({ [names[0]]: 'Acme Ltd' }), []);
  assert.equal(form.fields()[names[0]], 'Acme Ltd');
  assert.throws(() => d.setParagraphText('p9999', 'x'), /No paragraph p9999/);
});

test('paragraph ids stay with their paragraphs through edits; new paragraphs get new ids', () => {
  const d = openDocument(load('fields.docx'));
  const before = d.outline();
  const added = d.insertParagraph(null, 'A first line, added before everything.');
  d.deleteParagraph(d.search('Front matter paragraph 1.')[0].id, { tracked: false });
  const after = d.outline();
  assert.deepEqual(after.map((h) => h.id), before.map((h) => h.id));
  assert.ok(!before.some((h) => h.id === added) && d.paragraphs()[0].id === added);
  // An id read before the edits still finds its paragraph.
  d.setParagraphText(before[1].id, 'Methods');
  assert.equal(d.paragraphs().find((p) => p.id === before[1].id).text, 'Methods');
  assert.throws(() => d.range('p99999'), /No paragraph/);
});
