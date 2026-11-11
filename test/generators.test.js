import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { DocxSession } from '../src/ooxml/docx.js';

// Files from templating tools often give every picture the same drawing id and name (wp:docPr id="1"), which Word
// opens as it is: each picture is still shown, and saved, with its own picture.
test('pictures sharing one drawing id (as generators write them) are each shown and saved', () => {
  const files = unzipSync(new Uint8Array(fs.readFileSync(new URL('./fixtures/tiff.docx', import.meta.url))));
  files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace(/(<wp:docPr )id="\d+" name="[^"]*"/g, '$1id="1" name="Picture 1"'));
  const s = new DocxSession(zipSync(files), { imageUrl: () => 'blob:x' });
  const pics = [];
  s.doc.descendants((n) => { if (n.type.name === 'image') pics.push(n.attrs); });
  assert.equal(pics.length, 3);
  const xml = strFromU8(unzipSync(s.save())['word/document.xml']);
  const embeds = [...xml.matchAll(/r:embed="([^"]+)"/g)].map((m) => m[1]);
  assert.equal(new Set(embeds).size, 3, 'each keeps its own picture');
  assert.equal((xml.match(/<wp:docPr /g) || []).length, 3);
});
