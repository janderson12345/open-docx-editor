// tools/word-compare/lines.cjs, with one change for the generated documents: text on one baseline is split where a
// gap of over 12pt separates it, so the lines of columns side by side (and the parts of a line either side of a wide
// tab) are lines of their own instead of one line made of all of them. Word's PDFs and ours go through this alike.
// PDFs already done are skipped.
// usage: node tools/doc-generator/lines.cjs <pdfDir> <outDir>
const fs = require('fs');
const path = require('path');

const [dir, outDir] = process.argv.slice(2);
if (!outDir) { console.log('usage: node lines.cjs <pdfDir> <outDir>'); process.exit(1); }
fs.mkdirSync(outDir, { recursive: true });
// Symbol and Wingdings characters as Word's PDF text has them. The browser's print writes the font's own private-use
// code (U+F02D for Symbol's minus, U+F06C for Wingdings' round bullet) and can't write another without drawing another
// glyph, Word writes these as Unicode and every other code of these fonts (and all of Wingdings 2, Wingdings 3 and
// Webdings) as the private-use code too. Word's prints of every code 0x21-0xFF in each font, as w:sym, as private-use
// text and as the low code in the font's run, all three alike (..\why-tests\symtext, make.ps1). The same line held
// U+F02D in ours, U+2212 in Word's: 2,000 lines of bulleted lists in place but not matched. A no-op on Word's PDFs.
const WORD_SYMBOL = { 0x28: 0x28, 0x29: 0x29, 0x2b: 0x2b, 0x2d: 0x2212, 0x3d: 0x3d, 0x5e: 0x22a5, 0xae: 0x2192, 0xb7: 0x2022, 0xbe: 0x23af, 0xd3: 0xa9, 0xe3: 0xa9 };
const WORD_WINGDINGS = { 0x4a: 0x263a, 0x6c: 0x26ab, 0x6e: 0x25fc, 0x71: 0x2751, 0x73: 0x2b27, 0x74: 0x29eb, 0x75: 0x25c6, 0x76: 0x2756, 0xa7: 0x25aa, 0xd8: 0x27a2, 0xe0: 0x2192, 0xe8: 0x2794, 0xfc: 0x2713 };
const symbolTable = (name) => {
  const n = String(name || '').replace(/^[A-Z]{6}\+/, '').toLowerCase();
  return /^symbol(mt)?([-,].*)?$/.test(n) ? WORD_SYMBOL : /^wingdings(-regular)?([-,].*)?$/.test(n) ? WORD_WINGDINGS : null;
};
const asWord = (str, table) => str.replace(/[\uf021-\uf0ff]/g, (c) => { const u = table[c.charCodeAt(0) - 0xf000]; return u ? String.fromCodePoint(u) : c; });
(async () => {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  for (const f of fs.readdirSync(dir).filter((x) => /\.pdf$/i.test(x) && !fs.existsSync(path.join(outDir, x.replace(/\.pdf$/i, '.json'))))) {
    try {
      const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(path.join(dir, f))), disableFontFace: true, isEvalSupported: false, verbosity: 0, fontExtraProperties: true }).promise;
      const out = [];
      for (let p = 1; p <= doc.numPages; p++) {
        const page = await doc.getPage(p);
        const h = page.view[3];
        // Text items within 0.6pt of one baseline are one line (a run in another size sits a little higher or lower).
        const rows = new Map();
        const items = (await page.getTextContent()).items;
        // (The fonts' names only after the page's drawing is read, so only on a page with private-use text.)
        if (items.some((it) => /[\uf021-\uf0ff]/.test(it.str || ''))) {
          await page.getOperatorList();
          for (const it of items) {
            if (!/[\uf021-\uf0ff]/.test(it.str || '')) continue;
            let table = null;
            try { table = symbolTable(page.commonObjs.get(it.fontName).name); } catch { /* font not loaded: keep the text */ }
            if (table) it.str = asWord(it.str, table);
          }
        }
        for (const it of items) {
          if (!it.str || !it.str.trim()) continue;
          const y = Math.round((h - it.transform[5]) * 2) / 2; // baseline from the top, pt
          // Text turned on its side (a table's turned headings) is a line of its own, at where it starts: with the lines
          // across, items in other columns starting within 0.6pt of each other came together as one line, which a start
          // a fraction of a point off split differently (sample 1113's chemical tables, ~3 lines a page).
          if (Math.abs(it.transform[1]) > 1e-3 || Math.abs(it.transform[2]) > 1e-3) {
            const k = it.str.replace(/[\s\u200b]+/g, '').replace(/[\u2010\u2011]/g, '-').slice(0, 40);
            if (k.length >= 3) out.push([k, p, +(y * 4 / 3).toFixed(1), +((h - it.transform[5]) * 4 / 3).toFixed(3)]);
            continue;
          }
          let key = [...rows.keys()].find((k) => Math.abs(k - y) <= 0.6);
          if (key === undefined) rows.set((key = y), []);
          rows.get(key).push([it.transform[4], it.str, it.width, h - it.transform[5]]);
        }
        for (const [y, parts] of [...rows].sort((a, b) => a[0] - b[0])) {
          const segs = [];
          let end = -Infinity;
          for (const x of parts.sort((a, b) => a[0] - b[0])) {
            if (!segs.length || x[0] - end > 12) segs.push([]);
            segs[segs.length - 1].push(x);
            end = Math.max(end, x[0] + (x[2] || 0));
          }
          for (const seg of segs) {
            // (Hyphens as one: Word's PDF text keeps a U+2010 or U+2011 hyphen as it is, Edge's writes the same glyph as
            // U+002D, so the same line in the same place didn't match: sample 1113's hyphenated titles.)
            const k = seg.map((s) => s[1]).join('').replace(/[\s\u200b]+/g, '').replace(/[\u2010\u2011]/g, '-').slice(0, 40);
            // A line is [text, page, y to the half point in px, y exactly in px, x where it starts and ends in px]: the exact
            // y is the baseline of the piece holding most of its text (score.py's second measure; the rounded one stays the
            // first, comparable across runs); the x's let score.py tell a line cut into other pieces from text moved.
            const main = seg.reduce((m, s) => (s[1].trim().length > m[1].trim().length ? s : m));
            const x1 = Math.max(...seg.map((s) => s[0] + (s[2] || 0)));
            if (k.length >= 3) out.push([k, p, +(y * 4 / 3).toFixed(1), +(main[3] * 4 / 3).toFixed(3), +(seg[0][0] * 4 / 3).toFixed(1), +(x1 * 4 / 3).toFixed(1)]);
          }
        }
      }
      fs.writeFileSync(path.join(outDir, f.replace(/\.pdf$/i, '.json')), JSON.stringify(out));
    } catch (e) { console.log('FAILED', f, e.message); }
  }
})();
