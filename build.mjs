// Bundles the library and the demo with esbuild. `--serve` starts a dev server.
import * as esbuild from 'esbuild';
import fs from 'node:fs';

const serve = process.argv.includes('--serve');
fs.mkdirSync('demo/dist', { recursive: true });
fs.copyFileSync('src/editor/editor.css', 'demo/dist/open-docx-editor.css');
const common = { bundle: true, format: 'esm', sourcemap: true, target: 'es2020', logLevel: 'info' };
await esbuild.build({ ...common, entryPoints: ['src/index.js'], outfile: 'dist/open-docx-editor.js', minify: true });
fs.copyFileSync('src/editor/editor.css', 'dist/open-docx-editor.css');
const ctx = await esbuild.context({ ...common, entryPoints: ['demo/main.js'], outfile: 'demo/dist/demo.js' });
if (serve) {
  await ctx.watch();
  const { port } = await ctx.serve({ servedir: 'demo', port: 8080 });
  console.log(`demo: http://localhost:${port}/`);
} else {
  await ctx.rebuild();
  await ctx.dispose();
}
