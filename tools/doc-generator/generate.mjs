// Generates test documents from combinations of features (features.mjs): each feature on its own, every pair of
// features that can go together, and random mixes of three to six. Every document is reproducible: a pair's text comes
// from a seed derived from its name, a mix's features and text from its own seed.
// usage: node tools/doc-generator/generate.mjs <outDir> [--singles] [--pairs] [--random N] [--seed S] [--features a,b,c] [--list]
//   With none of --singles/--pairs/--random/--features: all singles, all pairs and 300 random mixes from seed 1.
//   --features a,b,c   one document of exactly these features (to regenerate a pair from a report)
//   --seed S --random 1  regenerates the mix r-S
//   --list             prints the features and exits
// Writes <outDir>/docs/<name>.docx and <outDir>/manifest.json ({ name: { features, seed } }).
import fs from 'node:fs';
import path from 'node:path';
import { FEATURES, BY_ID, compatible, compose } from './features.mjs';
import { pack, rng, hash } from './ooxml.mjs';

const args = process.argv.slice(2);
const flag = (n) => args.includes(n);
const opt = (n, d) => { const i = args.indexOf(n); return i >= 0 ? args[i + 1] : d; };
if (flag('--list')) {
  for (const f of FEATURES) console.log(`${f.id.padEnd(22)} ${f.kind.padEnd(8)} ${f.group === f.id ? '' : `[${f.group}] `}${f.desc}${f.bodyOnly ? ' (body only)' : ''}`);
  console.log(`${FEATURES.length} features`);
  process.exit(0);
}
const out = args.find((a, i) => !a.startsWith('--') && !(i > 0 && ['--random', '--seed', '--features'].includes(args[i - 1])));
if (!out) { console.log('usage: node tools/doc-generator/generate.mjs <outDir> [--singles] [--pairs] [--random N] [--seed S] [--features a,b,c] [--list]'); process.exit(1); }
const only = opt('--features');
const any = flag('--singles') || flag('--pairs') || opt('--random') || only;
const docs = [];
const title = (name, ids) => `Generated document ${name}: ${ids.length ? ids.join(' + ') : 'plain'}`;

if (only) {
  const ids = only.split(',').filter(Boolean);
  for (const id of ids) if (!BY_ID[id]) throw new Error(`no feature ${id} (see --list)`);
  if (!compatible(ids)) console.log('warning: these features are not normally combined');
  const name = ids.length === 1 ? `s-${ids[0]}` : ids.length === 2 ? `p-${ids[0]}__${ids[1]}` : `x-${ids.join('__')}`;
  docs.push({ name, ids, seed: hash(name) });
}
if (!any || flag('--singles')) {
  docs.push({ name: 's-plain', ids: [], seed: hash('s-plain') });
  for (const f of FEATURES) docs.push({ name: `s-${f.id}`, ids: [f.id], seed: hash(`s-${f.id}`) });
}
if (!any || flag('--pairs')) {
  for (let a = 0; a < FEATURES.length; a++) for (let b = a + 1; b < FEATURES.length; b++) {
    const ids = [FEATURES[a].id, FEATURES[b].id];
    if (!compatible(ids)) continue;
    const name = `p-${ids[0]}__${ids[1]}`;
    docs.push({ name, ids, seed: hash(name) });
  }
}
const nRandom = Number(opt('--random', any ? 0 : 300));
const seed0 = Number(opt('--seed', 1));
for (let k = 0; k < nRandom; k++) {
  const seed = seed0 + k;
  const r = rng(seed);
  const want = r.int(3, 6);
  const ids = [];
  for (let tries = 0; ids.length < want && tries < 200; tries++) {
    const f = r.pick(FEATURES).id;
    if (!ids.includes(f) && compatible([...ids, f])) ids.push(f);
  }
  docs.push({ name: `r-${seed}`, ids, seed });
}

fs.mkdirSync(path.join(out, 'docs'), { recursive: true });
const manifestPath = path.join(out, 'manifest.json');
const manifest = fs.existsSync(manifestPath) ? JSON.parse(fs.readFileSync(manifestPath, 'utf8')) : {};
for (const d of docs) {
  const parts = compose(d.ids, { seed: d.seed, title: title(d.name, d.ids) });
  fs.writeFileSync(path.join(out, 'docs', d.name + '.docx'), pack(parts));
  manifest[d.name] = { features: d.ids, seed: d.seed };
}
fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 1));
console.log(`generated ${docs.length} documents in ${path.join(out, 'docs')}`);
