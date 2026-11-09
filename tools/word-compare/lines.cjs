// The text lines of every PDF in a folder, as JSON beside them: [[key, page, baseline y in px], ...], the key being the
// line's first 40 characters without spaces (what lines are matched by). Word's PDFs and ours go through this alike.
// usage: node tools/word-compare/lines.cjs <pdfDir> <outDir>
const fs = require('fs');
const path = require('path');
const pdfjs = require('pdfjs-dist/legacy/build/pdf.js');

const [dir, outDir] = process.argv.slice(2);
if (!outDir) { console.log('usage: node lines.cjs <pdfDir> <outDir>'); process.exit(1); }
fs.mkdirSync(outDir, { recursive: true });
(async () => {
  for (const f of fs.readdirSync(dir).filter((x) => /\.pdf$/i.test(x))) {
    try {
      const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(path.join(dir, f))), disableFontFace: true, verbosity: 0 }).promise;
      const out = [];
      for (let p = 1; p <= doc.numPages; p++) {
        const page = await doc.getPage(p);
        const h = page.view[3];
        // Text items within 0.6pt of one baseline are one line (a run in another size sits a little higher or lower).
        const rows = new Map();
        for (const it of (await page.getTextContent()).items) {
          if (!it.str || !it.str.trim()) continue;
          const y = Math.round((h - it.transform[5]) * 2) / 2; // baseline from the top, pt
          let key = [...rows.keys()].find((k) => Math.abs(k - y) <= 0.6);
          if (key === undefined) rows.set((key = y), []);
          rows.get(key).push([it.transform[4], it.str]);
        }
        for (const [y, parts] of [...rows].sort((a, b) => a[0] - b[0])) {
          const k = parts.sort((a, b) => a[0] - b[0]).map((x) => x[1]).join('').replace(/[\s\u200b]+/g, '').slice(0, 40);
          if (k.length >= 3) out.push([k, p, +(y * 4 / 3).toFixed(1)]);
        }
      }
      fs.writeFileSync(path.join(outDir, f.replace(/\.pdf$/i, '.json')), JSON.stringify(out));
    } catch (e) { console.log('FAILED', f, e.message); }
  }
})();
