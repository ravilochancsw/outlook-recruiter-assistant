import * as esbuild from 'esbuild';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));

await esbuild.build({
  entryPoints: [path.join(here, 'entry.ts')],
  outfile: path.join(here, '.build', 'lib.mjs'),
  bundle: true,
  format: 'esm',
  platform: 'neutral',
  target: ['es2022'],
  logLevel: 'warning',
});
