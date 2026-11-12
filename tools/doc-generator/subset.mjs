// A run folder holding only the documents with given features (their documents, Word's PDFs and lines copied), for
// checking a fix quickly: then `run.mjs <subDir> --steps ours,score,report`.
// usage: node tools/doc-generator/subset.mjs <runDir> <subDir> <feature>[,<feature>...]   (documents with any of them)
import fs from 'node:fs';
import path from 'node:path';

const [from, to, list] = process.argv.slice(2);
if (!list) { console.log('usage: node subset.mjs <runDir> <subDir> <feature>[,<feature>...]'); process.exit(1); }
const want = list.split(',');
const manifest = JSON.parse(fs.readFileSync(path.join(from, 'manifest.json'), 'utf8'));
const names = Object.keys(manifest).filter((n) => n === 's-plain' || manifest[n].features.some((f) => want.includes(f)));
for (const d of ['docs', 'word-pdf', 'lines-word']) fs.mkdirSync(path.join(to, d), { recursive: true });
const sub = {};
for (const n of names) {
  sub[n] = manifest[n];
  for (const [d, ext] of [['docs', '.docx'], ['word-pdf', '.pdf'], ['lines-word', '.json']]) {
    const src = path.join(from, d, n + ext);
    if (fs.existsSync(src)) fs.copyFileSync(src, path.join(to, d, n + ext));
  }
}
fs.writeFileSync(path.join(to, 'manifest.json'), JSON.stringify(sub, null, 1));
console.log(`${names.length} documents in ${to}`);
