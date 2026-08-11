import * as esbuild from 'esbuild';
import { cp, mkdir, rm } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const outdir = path.join(root, 'dist');
const watch = process.argv.includes('--watch');

async function copyStatic() {
  await mkdir(outdir, { recursive: true });
  await cp(path.join(root, 'src/manifest.json'), path.join(outdir, 'manifest.json'));
  await cp(path.join(root, 'src/options/options.html'), path.join(outdir, 'options.html'));
  await cp(path.join(root, 'src/options/options.css'), path.join(outdir, 'options.css'));
  await cp(path.join(root, 'src/popup/popup.html'), path.join(outdir, 'popup.html'));
  await cp(path.join(root, 'src/popup/popup.css'), path.join(outdir, 'popup.css'));
  await cp(path.join(root, 'src/icons'), path.join(outdir, 'icons'), { recursive: true });
}

const options = {
  entryPoints: {
    content: path.join(root, 'src/content/index.ts'),
    background: path.join(root, 'src/background/service-worker.ts'),
    options: path.join(root, 'src/options/options.ts'),
    popup: path.join(root, 'src/popup/popup.ts'),
  },
  bundle: true,
  outdir,
  format: 'iife',
  target: ['chrome114'],
  sourcemap: watch ? 'inline' : false,
  minify: false,
  logLevel: 'info',
};

await rm(outdir, { recursive: true, force: true });
await copyStatic();

if (watch) {
  const ctx = await esbuild.context(options);
  await ctx.watch();
  console.log('[build] watching… reload the extension in chrome://extensions after each change');
} else {
  await esbuild.build(options);
  console.log('[build] wrote dist/ — load it via chrome://extensions → Load unpacked');
}
