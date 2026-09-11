import fs from 'node:fs';
import path from 'node:path';
import { build } from '../packages/xlsx/node_modules/esbuild/lib/main.js';

const root = path.resolve(import.meta.dirname, '..');
const out = path.join(root, 'playground', 'dist');
fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(out, { recursive: true });
fs.copyFileSync(path.join(root, 'playground', 'index.html'), path.join(out, 'index.html'));
await build({
  entryPoints: [path.join(root, 'playground', 'src', 'main.js')],
  bundle: true, format: 'esm', platform: 'browser', target: ['es2020'],
  outfile: path.join(out, 'assets', 'app.js'), sourcemap: true, logLevel: 'info',
});
