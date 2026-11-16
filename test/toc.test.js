import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { DocxSession } from '../src/ooxml/docx.js';
import { tocSwitches, tocLevel } from '../src/editor/toc.js';

const load = (f) => new Uint8Array(fs.readFileSync(new URL(`./fixtures/${f}`, import.meta.url)));

// toc-u-only.docx, made with Word COM and its three tables updated in Word: TOC \u \h \z, TOC \o "1-3" \h \z and
// TOC \o "1-3" \h \z \u over headings by style, levels set on paragraphs directly (1, 5, 9), custom styles with
// levels 2 and 6, and a style based on Heading 1 set to body text. Each table's entries (text before the tab, TOC level)
// as Word wrote them, and the ones our switches pick from the same paragraphs, should be the same.
test('a TOC field\'s switches pick the entries Word does (\\u alone: outline levels only, 1-9)', () => {
  const s = new DocxSession(load('toc-u-only.docx'));
  const tables = [];
  let cur = null;
  const body = [];
  s.doc.forEach((p) => {
    let instr = '';
    p.forEach((c) => { if (c.type.name === 'opaque_inline' && c.attrs.kind === 'instrText') instr += c.attrs.xml.replace(/<[^>]*>/g, ''); });
    const m = /\bTOC\b.*/.exec(instr);
    if (m) tables.push((cur = { instr: m[0], words: [] }));
    const toc = /^TOC(\d)$/.exec(p.attrs.styleId || '');
    if (toc && cur) {
      let t = '', tab = false;
      p.forEach((c) => { if (c.type.name === 'tab') tab = true; else if (!tab && c.isText) t += c.text; });
      cur.words.push(`${t.trim()}:${toc[1]}`);
    } else if (!m && p.textContent.trim()) { cur = null; body.push(p); }
  });
  assert.equal(tables.length, 3);
  for (const { instr, words } of tables) {
    const sw = tocSwitches(instr);
    const ours = body.map((p) => [p.textContent.trim(), tocLevel(s, p, sw)]).filter(([, l]) => l != null).map(([t, l]) => `${t}:${l + 1}`);
    assert.deepEqual(ours, words, instr);
  }
  assert.ok(tables[0].words.includes('Direct level 9:9') && tables[0].words.includes('Custom style level 6:6'));
});

test('TOC switches: a bare TOC or \\o alone is \\o "1-9"; \\u alone takes levels 1-9', () => {
  assert.deepEqual(tocSwitches(' TOC '), { o: true, u: false, from: 0, to: 8 });
  assert.deepEqual(tocSwitches(' TOC \\o '), { o: true, u: false, from: 0, to: 8 });
  assert.deepEqual(tocSwitches(' TOC \\o "2-5" \\u '), { o: true, u: true, from: 1, to: 4 });
  assert.deepEqual(tocSwitches(' TOC \\u \\h \\z '), { o: false, u: true, from: 0, to: 8 });
});
