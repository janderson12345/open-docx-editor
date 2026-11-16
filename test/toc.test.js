import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { DocxSession } from '../src/ooxml/docx.js';
import { tocSwitches, tocLevel } from '../src/editor/toc.js';
import { unescapeXml } from '../src/ooxml/xml.js';

const load = (f) => new Uint8Array(fs.readFileSync(new URL(`./fixtures/${f}`, import.meta.url)));

/** Each TOC field's instruction and its entries as Word wrote them ("text before the tab:TOC level"), and our picks. */
function compare(file) {
  const s = new DocxSession(load(file));
  const tables = [];
  let cur = null;
  const body = [];
  s.doc.forEach((p) => {
    let instr = '';
    p.forEach((c) => { if (c.type.name === 'opaque_inline' && c.attrs.kind === 'instrText') instr += unescapeXml(/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/.exec(c.attrs.xml)?.[1] ?? ''); }); // (its text, as merge.js reads it)
    const m = /\bTOC\b.*/.exec(instr);
    if (m) tables.push((cur = { instr: m[0], words: [] }));
    const toc = /^TOC(\d)$/.exec(p.attrs.styleId || '');
    if (toc && cur) {
      let t = '', tab = false;
      p.forEach((c) => { if (c.type.name === 'tab') tab = true; else if (!tab && c.isText) t += c.text; });
      cur.words.push(`${t.trim()}:${toc[1]}`);
    } else if (!m && p.textContent.trim()) { cur = null; body.push(p); }
  });
  return tables.map(({ instr, words }) => {
    const sw = tocSwitches(instr);
    const ours = body.map((p) => [p.textContent.trim(), tocLevel(s, p, sw)]).filter(([, l]) => l != null).map(([t, l]) => `${t}:${l + 1}`);
    return { instr, words, ours };
  });
}

// toc-u-only.docx, made with Word COM and its three tables updated in Word: TOC \u \h \z, TOC \o "1-3" \h \z and
// TOC \o "1-3" \h \z \u over headings by style, levels set on paragraphs directly (1, 5, 9), custom styles with
// levels 2 and 6, and a style based on Heading 1 set to body text. Our switches should pick Word's entries.
test('a TOC field\'s switches pick the entries Word does (\\u alone: outline levels only, 1-9)', () => {
  const tables = compare('toc-u-only.docx');
  assert.equal(tables.length, 3);
  for (const { instr, words, ours } of tables) assert.deepEqual(ours, words, instr);
  assert.ok(tables[0].words.includes('Direct level 9:9') && tables[0].words.includes('Custom style level 6:6'));
});

// toc-t-switch.docx, the same paragraphs: TOC \t "Section Title,1,Deep Title,3", \o "1-3" \t "Section Title,1",
// \t "heading 2,1,Section Title" (a lower-case name, a name with no level) and \u \t "Deep Title,2", updated in Word.
test('a TOC field\'s \\t switch picks the entries Word does (named styles at their levels; \\t alone, those only)', () => {
  const tables = compare('toc-t-switch.docx');
  assert.equal(tables.length, 4);
  for (const { instr, words, ours } of tables) assert.deepEqual(ours, words, instr);
  assert.deepEqual(tables[3].words, ['Custom style level 6:2']);
});

test('TOC switches: a bare TOC or \\o alone is \\o "1-9"; \\u alone takes levels 1-9; \\t names styles', () => {
  const plain = (sw) => ({ ...sw, styles: Object.fromEntries(sw.styles) });
  assert.deepEqual(plain(tocSwitches(' TOC ')), { o: true, u: false, from: 0, to: 8, styles: {} });
  assert.deepEqual(plain(tocSwitches(' TOC \\o ')), { o: true, u: false, from: 0, to: 8, styles: {} });
  assert.deepEqual(plain(tocSwitches(' TOC \\o "2-5" \\u ')), { o: true, u: true, from: 1, to: 4, styles: {} });
  assert.deepEqual(plain(tocSwitches(' TOC \\u \\h \\z ')), { o: false, u: true, from: 0, to: 8, styles: {} });
  assert.deepEqual(plain(tocSwitches(' TOC \\t "Title,1,Sub Title;2,Note" ')), { o: false, u: false, from: 0, to: 8, styles: { title: 0, 'sub title': 1, note: 0 } });
  assert.deepEqual(plain(tocSwitches(' TOC \\o " 2 - 4 " ')), { o: true, u: false, from: 1, to: 3, styles: {} });
});

test('TOC switches: a field code with a long run of spaces is read in linear time', () => {
  const t0 = performance.now();
  for (const pad of [' '.repeat(40000), ' '.repeat(40000) + '"', ' '.repeat(20000) + '"' + ' '.repeat(20000)]) {
    assert.equal(tocSwitches(` TOC \\o${pad}x \\t${pad}`).from, 0);
  }
  assert.ok(performance.now() - t0 < 200, `${(performance.now() - t0).toFixed(0)}ms`); // (was 2.6s: 1.2s, 1.2s, 0.3s)
});
