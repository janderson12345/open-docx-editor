// Runs the generated documents (generate.mjs) against Word: Word prints each to PDF (word-pdf.ps1, once: its PDFs are
// kept), the editor prints each (print-fast.mjs), the lines of both are extracted
// (lines.cjs) and compared (score.py), and report.mjs lists which feature combinations differ from
// Word, worst first.
// usage: node tools/doc-generator/run.mjs <outDir> [--steps word,ours,score,report] [--run name] [--timeout seconds]
//   Windows, with Word, Microsoft Edge, Node, Python 3 and PowerShell 7. Run from the repository. Steps already done
//   are picked up where they stopped (PDFs already made are kept). A manual tool: not part of the tests or CI.
import fs from 'node:fs';
import path from 'node:path';
import { spawn, execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const repo = path.resolve(here, '..', '..');
const args = process.argv.slice(2);
const opt = (n, d) => { const i = args.indexOf(n); return i >= 0 ? args[i + 1] : d; };
const out = args.find((a, i) => !a.startsWith('--') && !(i > 0 && args[i - 1].startsWith('--')));
if (!out) { console.log('usage: node tools/doc-generator/run.mjs <outDir> [--steps word,ours,score,report] [--run name] [--timeout seconds]'); process.exit(1); }
const OUT = path.resolve(out);
const steps = opt('--steps', 'word,ours,score,report').split(',');
const runName = opt('--run', 'ours');
const timeout = Number(opt('--timeout', 120)) * 1000;
const docsDir = path.join(OUT, 'docs');
const log = (s) => { const line = `${new Date().toISOString().slice(11, 19)} ${s}`; console.log(line); fs.appendFileSync(path.join(OUT, 'run.log'), line + '\n'); };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const docs = () => fs.readdirSync(docsDir).filter((f) => f.endsWith('.docx')).sort();

function step(exe, argv, { cwd = repo, logName } = {}) {
  return new Promise((resolve, reject) => {
    const p = spawn(exe, argv, { cwd, stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
    const o = fs.createWriteStream(path.join(OUT, `${logName}.log`));
    p.stdout.pipe(o); p.stderr.pipe(o);
    p.on('exit', (code) => (code === 0 ? resolve() : reject(new Error(`${exe} ${argv.join(' ')} failed (${code}); see ${logName}.log`))));
  });
}

// ---- Word ----

/** Is this process one of Word's? (The watcher only ever ends the Word its own script started.) */
const isWord = (pid) => { try { return /WINWORD\.EXE/i.test(execFileSync('tasklist', ['/FI', `PID eq ${pid}`, '/FO', 'CSV', '/NH'], { encoding: 'utf8' })); } catch { return false; } };
const kill = (pid) => { try { execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' }); } catch { /* gone already */ } };

async function word() {
  const pdfDir = path.join(OUT, 'word-pdf');
  fs.mkdirSync(pdfDir, { recursive: true });
  const failPath = path.join(OUT, 'word-failures.json');
  const failures = fs.existsSync(failPath) ? JSON.parse(fs.readFileSync(failPath, 'utf8')) : {};
  const saveFailures = () => fs.writeFileSync(failPath, JSON.stringify(failures, null, 1));
  let idle = 0;
  for (;;) {
    const todo = docs().filter((f) => { const n = f.slice(0, -5); return !failures[n] && !fs.existsSync(path.join(pdfDir, n + '.pdf')); });
    if (!todo.length) break;
    log(`word: ${todo.length} to print`);
    const list = path.join(OUT, 'word-list.txt'), status = path.join(OUT, 'word-status.txt');
    fs.writeFileSync(list, todo.map((f) => path.join(docsDir, f)).join('\r\n'));
    fs.writeFileSync(status, '');
    const child = spawn('pwsh', ['-NoProfile', '-NonInteractive', '-File', path.join(here, 'word-pdf.ps1'), '-List', list, '-Out', pdfDir, '-Status', status], { stdio: 'ignore', windowsHide: true });
    let exited = false; child.on('exit', () => { exited = true; });
    let wordPid = 0, current = null, since = Date.now(), seen = 0, done = 0, ended = false;
    for (;;) {
      await sleep(2000);
      const lines = fs.readFileSync(status, 'utf8').split(/\r?\n/).filter(Boolean);
      for (const l of lines.slice(seen)) {
        const [k, n, ...why] = l.split(' ');
        if (k === 'PID') wordPid = Number(n);
        else if (k === 'START') { current = n; since = Date.now(); }
        else if (k === 'DONE') { current = null; done++; if (done % 100 === 0) log(`word: ${done} printed this pass`); }
        else if (k === 'FAIL') { current = null; failures[n] = why.join(' ') || 'failed'; saveFailures(); log(`word: FAIL ${n} ${failures[n]}`); }
        else if (k === 'END') ended = true;
      }
      seen = lines.length;
      if (current && Date.now() - since > timeout) {
        log(`word: TIMEOUT ${current}; ending this run's Word (pid ${wordPid})`);
        failures[current] = 'TIMEOUT'; saveFailures();
        kill(child.pid);
        if (wordPid > 0 && isWord(wordPid)) kill(wordPid);
        break;
      }
      if (exited) {
        if (!ended && current) { failures[current] = 'Word stopped'; saveFailures(); log(`word: Word stopped at ${current}`); }
        if (wordPid > 0 && !ended && isWord(wordPid)) kill(wordPid);
        break;
      }
    }
    idle = done ? 0 : idle + 1;
    if (idle >= 3) { log('word: three passes printed nothing; stopping'); break; }
  }
  log(`word: ${fs.readdirSync(pdfDir).filter((f) => f.endsWith('.pdf')).length} PDFs, ${Object.keys(failures).length} failures`);
}

// ---- ours, lines, score, report ----

async function ours() {
  log('ours: building the demo');
  await step(process.execPath, ['build.mjs'], { logName: 'build' });
  log('ours: printing');
  await step(process.execPath, [path.join(here, 'print-fast.mjs'), docsDir, path.join(OUT, `ours-${runName}`)], { logName: `print-${runName}` });
  log(`ours: ${fs.readdirSync(path.join(OUT, `ours-${runName}`)).length} PDFs`);
}

async function score() {
  log('score: lines of Word\'s PDFs and ours');
  // (Word's lines once: lines.cjs skips the PDFs it has done, and Word's don't change.)
  await step(process.execPath, [path.join(here, 'lines.cjs'), path.join(OUT, 'word-pdf'), path.join(OUT, 'lines-word')], { logName: 'lines-word' });
  await step(process.execPath, [path.join(here, 'lines.cjs'), path.join(OUT, `ours-${runName}`), path.join(OUT, `lines-${runName}`)], { logName: `lines-${runName}` });
  await step('python', [path.join(here, 'score.py'), path.join(OUT, 'lines-word'), path.join(OUT, `lines-${runName}`), path.join(OUT, `scores-${runName}.json`)], { logName: `score-${runName}` });
}

async function report() {
  await step(process.execPath, [path.join(here, 'report.mjs'), OUT, runName], { logName: `report-${runName}` });
  log(`report: ${path.join(OUT, `report-${runName}.md`)}`);
}

const STEPS = { word, ours, score, report };
for (const s of steps) { if (!STEPS[s]) throw new Error(`no step ${s}`); await STEPS[s](); }
