import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { legacyKey, hashPassword, readProtection, withProtection, withoutEnforcement, checkPassword, lockedStyles, markedFinal, withoutFinal } from '../src/ooxml/protect.js';
import { sha512, sha1 } from '../src/ooxml/sha.js';
import { createHash } from 'node:crypto';
import { DocxSession } from '../src/ooxml/docx.js';

const load = (f) => new Uint8Array(fs.readFileSync(new URL(`./fixtures/${f}`, import.meta.url)));
const settings = (bytes) => strFromU8(unzipSync(bytes)['word/settings.xml']);

test('SHA-512 and SHA-1 agree with Node\'s', () => {
  for (const n of [0, 1, 55, 56, 64, 111, 112, 128, 300]) {
    const b = new Uint8Array(n).map((_, i) => (i * 37 + n) & 255);
    assert.equal(Buffer.from(sha512(b)).toString('hex'), createHash('sha512').update(b).digest('hex'));
    assert.equal(Buffer.from(sha1(b)).toString('hex'), createHash('sha1').update(b).digest('hex'));
  }
});

test('Word\'s password key and hash (ECMA-376\'s example; Word lifted the protection with these)', async () => {
  assert.equal(legacyKey('Example'), '7EEDCE64');
  // (Only the first 15 characters count, as in Word.)
  assert.equal(legacyKey('abcdefghijklmnopqrst'), legacyKey('abcdefghijklmno'));
  const salt = new Uint8Array(16);
  assert.equal(await hashPassword('Example', { salt, algorithm: 'SHA-512', spin: 100000 }), 'gCxNcoHiQlDc+D8yc54Xazfv8o51Yg5th1WFbpLW8Nc+RzbAsHs0WzJPDAmf8Ez3pN/8frhWLYyE7rq6wdIg7w==');
});

test('protection written, read, checked and lifted as Word keeps it', async () => {
  const xml = await withProtection('<w:settings><w:trackRevisions/><w:defaultTabStop w:val="720"/></w:settings>', { edit: 'comments', password: 'secret' });
  assert.match(xml, /<w:trackRevisions\/><w:documentProtection w:edit="comments" w:enforcement="1" w:cryptProviderType="rsaAES"[^>]*w:cryptAlgorithmSid="14" w:cryptSpinCount="100000" w:hash="[^"]+" w:salt="[^"]+"\/><w:defaultTabStop/);
  const p = readProtection(xml);
  assert.deepEqual([p.edit, p.enforced, p.formatting, p.password], ['comments', true, false, true]);
  assert.equal(await checkPassword(p.tag, 'wrong'), false);
  assert.equal(await checkPassword(p.tag, 'secret'), true);
  const lifted = readProtection(withoutEnforcement(xml));
  assert.deepEqual([lifted.edit, lifted.enforced, lifted.password], ['comments', false, false]);
  // The older way of writing the hash (w:hashValue, w:algorithmName...) is read too.
  const h = await hashPassword('pw', { salt: 'AAAAAAAAAAAAAAAAAAAAAA==', algorithm: 'SHA-1', spin: 10 });
  assert.equal(await checkPassword(`<w:documentProtection w:edit="readOnly" w:enforcement="1" w:algorithmName="SHA-1" w:hashValue="${h}" w:saltValue="AAAAAAAAAAAAAAAAAAAAAA==" w:spinCount="10"/>`, 'pw'), true);
});

test('a session protects and unprotects; the fixtures read as Word reads them', async () => {
  const s = new DocxSession(load('fields.docx')); // (no settings part: one is made)
  await s.protect({ edit: 'readOnly', password: 'pw' });
  assert.equal(s.protection.edit, 'readOnly');
  const again = new DocxSession(s.save());
  assert.ok(again.protection.enforced && again.protection.password);
  assert.equal(await again.unprotect('no'), false);
  assert.equal(await again.unprotect('pw'), true);
  assert.equal(again.protection.enforced, false);
  assert.match(settings(again.save()), /<w:documentProtection w:edit="readOnly" w:enforcement="0"\/>/);
  const sec = new DocxSession(load('protect-sections.docx'));
  assert.ok(sec.formsProtected);
  const fmt = new DocxSession(load('protect-format.docx')).protection;
  assert.ok(fmt.formatting && fmt.locked.has('Heading2') && !fmt.locked.has('Heading1'));
  const fin = new DocxSession(load('protect-final.docx'));
  assert.ok(fin.protection.final);
  fin.clearFinal();
  assert.ok(!fin.protection.final);
});

test('locked styles as Word reads them: a custom style by its own w:locked, a built-in one by its latent entry too', () => {
  const styles = '<w:styles><w:latentStyles w:defLockedState="0"><w:lsdException w:name="heading 2" w:locked="1"/></w:latentStyles>'
    + '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:locked/></w:style>'
    + '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:locked/></w:style>'
    + '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Mine"><w:name w:val="Mine"/><w:locked/></w:style></w:styles>';
  assert.deepEqual([...lockedStyles(styles)].sort(), ['Heading2', 'Mine', 'heading 2', 'mine']);
  assert.ok(markedFinal('<property name="_MarkAsFinal"><vt:bool>true</vt:bool></property>'));
  assert.ok(!markedFinal(withoutFinal('<p><property name="_MarkAsFinal"><vt:bool>true</vt:bool></property></p>')));
});
