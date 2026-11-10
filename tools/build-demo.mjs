// Rebuilds the standalone demo page in place: bundles tools/demo-entry.js and src/editor/editor.css into an
// existing open-docx-editor-demo.html (its largest <script> and the editor part of its <style> are replaced).
// usage: node tools/build-demo.mjs <path/to/open-docx-editor-demo.html>
import { build } from 'esbuild';
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
const root = fileURLToPath(new URL('..', import.meta.url));
const page = process.argv[2];
if (!page) throw new Error('usage: node tools/build-demo.mjs <demo.html>');
const out = await build({ entryPoints: [root + 'tools/demo-entry.js'], bundle: true, format: 'iife', minify: true, write: false });
const js = out.outputFiles[0].text.replace(/<\/script/g, '<\\/script');
const css = readFileSync(root + 'src/editor/editor.css', 'utf8');
let h = readFileSync(page, 'utf8');
const scripts = [...h.matchAll(/<script>([\s\S]*?)<\/script>/g)];
const big = scripts.reduce((a, b) => (b[1].length > a[1].length ? b : a));
const at = big.index + '<script>'.length;
h = h.slice(0, at) + js + h.slice(at + big[1].length);
const st = /<style>([\s\S]*?)<\/style>/.exec(h);
const keep = st[1].slice(0, st[1].indexOf('.odx-canvas {'));
const s0 = st.index + '<style>'.length;
h = h.slice(0, s0) + keep + css + '\n' + h.slice(s0 + st[1].length);
// (Read as UTF-8 when opened from disk: the bundle's regular expressions hold characters beyond ASCII.)
if (!/^\s*(<!doctype[^>]*>\s*)?<meta charset=/i.test(h)) h = `<meta charset="utf-8">\n${h}`;
writeFileSync(page, h);
console.log('built', page, Math.round(h.length / 1024) + ' KB');
