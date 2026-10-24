import { build } from 'esbuild';
import { readdir } from 'node:fs/promises';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const result = await build({
  absWorkingDir: root, entryPoints: ['src/extension.ts', 'src/react/main.tsx'],
  bundle: true, write: false, metafile: true, outdir: 'out/inventory',
  packages: 'external', platform: 'neutral', logLevel: 'silent',
  loader: { '.less': 'empty', '.css': 'empty', '.svg': 'empty', '.png': 'empty', '.jpg': 'empty' },
});
const reachable = new Set(Object.keys(result.metafile.inputs).map(p => p.replaceAll('\\', '/')));
const files = (await readdir(resolve(root, 'src'), { recursive: true, withFileTypes: true }))
  .filter(item => item.isFile() && /\.(?:[cm]?[jt]sx?)$/.test(item.name))
  .map(item => `${item.parentPath}/${item.name}`.replaceAll('\\', '/').slice(root.replaceAll('\\', '/').length + 1));
const unreachable = files.filter(file => !reachable.has(file)).sort();
console.log(JSON.stringify({ reachableSourceFiles: [...reachable].filter(p => p.startsWith('src/')).length,
  note: 'Candidates only: type-only imports, tests, styles and assets require separate review before deletion.', unreachable }, null, 2));
