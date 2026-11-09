import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxPackage } from '../src/ooxml/package.js';
import { blankDocxBytes } from '../src/ooxml/template.js';

test('with a prefix of its own (co-editing), an editor\'s new relationship ids never take another editor\'s', () => {
  const a = DocxPackage.fromBytes(blankDocxBytes());
  const b = DocxPackage.fromBytes(blankDocxBytes());
  a.idPrefix = 'rIdc11n';
  b.idPrefix = 'rIdc22n';
  const ida = a.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'media/a.png');
  const idb = b.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'media/b.png');
  assert.notEqual(ida, idb);
  assert.match(ida, /^rIdc11n\d+$/);
  // The other's part, shared under its id, is added under that id.
  assert.equal(a.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'media/b.png', false, a.mainPath, idb), idb);
  // Without a prefix, the next free rIdN as before.
  const c = DocxPackage.fromBytes(blankDocxBytes());
  assert.match(c.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'media/c.png'), /^rId\d+$/);
});
