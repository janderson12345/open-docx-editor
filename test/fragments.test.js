// XML kept from a part and parsed again on its own (a table's w:tblPr, a run's w:rPr in a tracked change) declares the
// prefixes it uses that its part's root declared: Word 2023's w16du:dateUtc in a tracked table change stopped
// documents opening in browsers (test/fixtures/w16du-tracked-table.docx, opened by the browser suite).
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { missingDecls } from '../src/ooxml/xml.js';

test('the prefixes a piece of XML uses are declared, Word\'s by their own namespaces', () => {
  const xml = '<w:tblPr><w:tblPrChange w:id="1" w16du:dateUtc="2026-11-18T14:00:00Z"><w:tblPr/></w:tblPrChange><w14:x/></w:tblPr>';
  const d = missingDecls(xml, ['w']);
  assert.match(d, /xmlns:w16du="http:\/\/schemas\.microsoft\.com\/office\/word\/2023\/wordml\/word16du"/);
  assert.match(d, /xmlns:w14="http:\/\/schemas\.microsoft\.com\/office\/word\/2010\/wordml"/);
  assert.doesNotMatch(d, /xmlns:w=/);
  // Any other prefix gets a stand-in, so the piece still parses; one its first tag declares isn't declared again.
  assert.match(missingDecls('<w:p><foo:bar/></w:p>', ['w']), /xmlns:foo="urn:odx:undeclared:foo"/);
  assert.equal(missingDecls('<w:tblPr xmlns:w14="x"><w14:y/></w:tblPr>', ['w']), '');
  // Text that looks like a prefix (a URL) isn't one; the prototype's names are only prefixes.
  assert.equal(missingDecls('<w:t>see http://example.com and a:b</w:t>', ['w']).includes('xmlns:http'), false);
  assert.match(missingDecls('<constructor:x __proto__:y="1"/>'), /xmlns:constructor="urn:odx:undeclared:constructor"/);
});

test('a piece using hundreds of prefixes is declared in one pass, a bounded number of them', () => {
  const xml = Array.from({ length: 5000 }, (_, i) => `<p${i}:x/>`).join('');
  const t0 = Date.now();
  const d = missingDecls(xml);
  assert.ok(Date.now() - t0 < 500);
  assert.ok((d.match(/xmlns:/g) || []).length <= 300);
});
