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
(async () => {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  for (const f of fs.readdirSync(dir).filter((x) => /\.pdf$/i.test(x) && !fs.existsSync(path.join(outDir, x.replace(/\.pdf$/i, '.json'))))) {
    try {
      const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(path.join(dir, f))), disableFontFace: true, isEvalSupported: false, verbosity: 0 }).promise;
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
          rows.get(key).push([it.transform[4], it.str, it.width]);
        }
        for (const [y, parts] of [...rows].sort((a, b) => a[0] - b[0])) {
          const segs = [];
          let end = -Infinity;
          for (const x of parts.sort((a, b) => a[0] - b[0])) {
            if (!segs.length || x[0] - end > 12) segs.push([]);
            segs[segs.length - 1].push(x[1]);
            end = Math.max(end, x[0] + (x[2] || 0));
          }
          for (const seg of segs) {
            const k = seg.join('').replace(/[\s\u200b]+/g, '').slice(0, 40);
            if (k.length >= 3) out.push([k, p, +(y * 4 / 3).toFixed(1)]);
          }
        }
      }
      fs.writeFileSync(path.join(outDir, f.replace(/\.pdf$/i, '.json')), JSON.stringify(out));
    } catch (e) { console.log('FAILED', f, e.message); }
  }
})();
