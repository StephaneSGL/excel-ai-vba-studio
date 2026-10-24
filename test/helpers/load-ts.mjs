import { build } from 'esbuild';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

export async function withModule(entry, run, mocks = {}) {
  const root = resolve(import.meta.dirname, '../..');
  const temporary = await mkdtemp(join(tmpdir(), 'excel-release-test-'));
  try {
    const outfile = join(temporary, 'module.cjs');
    await build({
      ...(typeof entry === 'string' ? { entryPoints: [resolve(root, entry)] } : {
        stdin: { contents: entry.source, resolveDir: root, sourcefile: 'release-test-entry.ts', loader: 'ts' },
      }), outfile, bundle: true,
      format: 'cjs', platform: 'node', logLevel: 'silent',
      plugins: [{ name: 'test-boundaries', setup(api) {
        api.onResolve({ filter: /.*/ }, ({ path }) => Object.hasOwn(mocks, path)
          ? { path, namespace: 'test-boundary' } : undefined);
        api.onLoad({ filter: /.*/, namespace: 'test-boundary' }, ({ path }) => ({ contents: mocks[path] }));
      } }],
    });
    await run(await import(pathToFileURL(outfile).href));
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
}
