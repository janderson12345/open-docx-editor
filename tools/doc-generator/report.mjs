// The report of a run: which features and feature combinations differ from Word, worst first.
// usage: node tools/doc-generator/report.mjs <outDir> [run]   → <outDir>/report-<run>.md and .json
import fs from 'node:fs';
import path from 'node:path';
import { FEATURES } from './features.mjs';

const [out, runName = 'ours'] = process.argv.slice(2);
const read = (f, d) => (fs.existsSync(path.join(out, f)) ? JSON.parse(fs.readFileSync(path.join(out, f), 'utf8')) : d);
const manifest = read('manifest.json', {});
const scores = read(`scores-${runName}.json`, {});
const wordFail = read('word-failures.json', {});
const names = Object.keys(manifest);
const scored = names.filter((n) => scores[n]);
const mean = (xs) => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : 0);
const f1 = (x) => (x == null ? '' : x.toFixed(1));
const regen = (n) => (n.startsWith('r-') ? `--seed ${manifest[n].seed} --random 1` : `--features ${manifest[n].features.join(',') || '(none)'}`);
const oursDir = path.join(out, `ours-${runName}`);
const oursMissing = names.filter((n) => !wordFail[n] && fs.existsSync(path.join(out, 'word-pdf', n + '.pdf')) && !fs.existsSync(path.join(oursDir, n + '.pdf')));

// Features: alone, and over every document that has them.
const alone = Object.fromEntries(FEATURES.map((f) => [f.id, scores[`s-${f.id}`]?.score]));
const plain = scores['s-plain']?.score;
const feats = FEATURES.map((f) => {
  const with_ = scored.filter((n) => manifest[n].features.includes(f.id)).map((n) => scores[n].score);
  return { id: f.id, desc: f.desc, alone: alone[f.id], mean: mean(with_), docs: with_.length, bad: with_.filter((s) => s > 5).length };
}).sort((a, b) => b.mean - a.mean);

// Pairs worse together than either feature alone.
const pairs = scored.filter((n) => n.startsWith('p-')).map((n) => {
  const [a, b] = manifest[n].features;
  const base = Math.max(alone[a] ?? 0, alone[b] ?? 0);
  return { name: n, a, b, score: scores[n].score, base, extra: scores[n].score - base };
}).filter((p) => p.extra > 2).sort((x, y) => y.extra - x.extra);

const worst = scored.map((n) => ({ name: n, ...scores[n] })).sort((a, b) => b.score - a.score);
const same = scored.filter((n) => scores[n].score === 0).length;

const md = [];
md.push(`# Generated documents compared with Word (run ${runName})`, '');
md.push(`${scored.length} of ${names.length} documents scored; ${same} (${(100 * same / Math.max(1, scored.length)).toFixed(1)}%) lay out exactly as Word's (every line on the same page within 1px). Mean score ${f1(mean(scored.map((n) => scores[n].score)))}; the plain document scores ${f1(plain)}.`);
const sum = (k) => scored.reduce((t, n) => t + scores[n][k], 0);
const lines = sum('wordLines'), within = sum('within1px'), pagesRight = scored.filter((n) => scores[n].wordPages === scores[n].ourPages).length;
md.push(`Lines: ${within} of Word's ${lines} (${(100 * within / Math.max(1, lines)).toFixed(2)}%) on the same page within 1px; page counts right in ${pagesRight} of ${scored.length}.`);
md.push(`Word could not print ${Object.keys(wordFail).length}; the editor did not print ${oursMissing.length} that Word did.`, '');
md.push('A score is the share of Word\'s lines we don\'t place on the same page within 1px, plus our lines Word doesn\'t have, plus 20 per page of difference: 0 is the same as Word.', '');
md.push('Regenerate any document with `node tools/doc-generator/generate.mjs <dir> <the arguments shown>`.', '');

md.push('## Features, worst first', '', '| feature | what | alone | mean with others | docs | docs over 5 |', '|---|---|---:|---:|---:|---:|');
for (const f of feats) md.push(`| ${f.id} | ${f.desc} | ${f1(f.alone)} | ${f1(f.mean)} | ${f.docs} | ${f.bad} |`);
md.push('');

md.push('## Pairs worse together than either alone', '', 'How much worse the pair is than the worse of its two features alone.', '', '| pair | score | worse feature alone | worse by |', '|---|---:|---:|---:|');
for (const p of pairs.slice(0, 150)) md.push(`| ${p.a} + ${p.b} | ${f1(p.score)} | ${f1(p.base)} | ${f1(p.extra)} |`);
if (pairs.length > 150) md.push(`| … ${pairs.length - 150} more in the JSON | | | |`);
md.push('');

md.push('## Documents, worst first', '', '| document | score | pages Word/ours | Word lines | missing | extra | other page | off >1px | first line off (text, page, y) | regenerate |', '|---|---:|---:|---:|---:|---:|---:|---:|---|---|');
for (const d of worst.slice(0, 200)) {
  const first = d.firstOff ? `${d.firstOff[0]} (p${d.firstOff[1]}, ${d.firstOff[2]})` : d.firstMissing ? `missing: ${d.firstMissing[0]}` : '';
  md.push(`| ${d.name} | ${f1(d.score)} | ${d.wordPages}/${d.ourPages} | ${d.wordLines} | ${d.missing} | ${d.extra} | ${d.otherPage} | ${d.drift} | ${first.replace(/\|/g, '/')} | \`${regen(d.name)}\` |`);
}
md.push('');
if (Object.keys(wordFail).length) {
  md.push('## Word could not print', '', '| document | why | regenerate |', '|---|---|---|');
  for (const [n, why] of Object.entries(wordFail)) md.push(`| ${n} | ${why.replace(/\|/g, '/')} | \`${manifest[n] ? regen(n) : ''}\` |`);
  md.push('');
}
if (oursMissing.length) md.push('## The editor did not print', '', oursMissing.map((n) => `- ${n}`).join('\n'), '');

fs.writeFileSync(path.join(out, `report-${runName}.md`), md.join('\n'));
fs.writeFileSync(path.join(out, `report-${runName}.json`), JSON.stringify({ scored: scored.length, total: names.length, same, plain, features: feats, pairs, worst, wordFailures: wordFail, oursMissing }, null, 1));
console.log(`report: ${scored.length} scored, ${same} the same as Word`);
