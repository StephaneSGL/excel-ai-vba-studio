import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { mkdtemp, rm, truncate, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const root = resolve(import.meta.dirname, '..');
const buildDirectory = await mkdtemp(join(tmpdir(), 'excel-file-open-bounds-'));
const outputPath = join(buildDirectory, 'office-content.mjs');

globalThis.__officeFileOpenMock = {
  bytes: new Uint8Array([1, 2, 3, 4]),
  readCalls: 0,
  statCalls: 0,
  statSize: 4,
};

try {
  await build({
    entryPoints: [join(root, 'src', 'provider', 'handlers', 'officeContent.ts')],
    outfile: outputPath,
    bundle: true,
    format: 'esm',
    platform: 'node',
    logLevel: 'silent',
    tsconfig: join(root, 'tsconfig.json'),
    plugins: [{
      name: 'mock-vscode',
      setup(buildApi) {
        buildApi.onResolve({ filter: /^vscode$/ }, () => ({
          path: 'vscode',
          namespace: 'mock-vscode',
        }));
        buildApi.onLoad(
          { filter: /.*/, namespace: 'mock-vscode' },
          () => ({
            contents: `
              export const FilePermission = { Readonly: 1 };
              export class Uri {}
              export const workspace = {
                fs: {
                  async stat() {
                    const state = globalThis.__officeFileOpenMock;
                    state.statCalls += 1;
                    return { size: state.statSize, permissions: 0 };
                  },
                  async readFile() {
                    const state = globalThis.__officeFileOpenMock;
                    state.readCalls += 1;
                    return state.bytes;
                  },
                },
              };
            `,
          }),
        );
      },
    }],
  });

  const { emitFileOfficeOpen } = await import(
    `${pathToFileURL(outputPath).href}?cache=${Date.now()}`
  );
  const localPath = join(buildDirectory, 'bounded.csv');
  await writeFile(localPath, new Uint8Array([1, 2, 3, 4]));
  const uri = {
    scheme: 'file',
    fsPath: localPath,
    toString: () => `file:///${localPath.replaceAll('\\', '/')}`,
  };

  const emitted = [];
  const handler = {
    emit(event, payload) {
      emitted.push({ event, payload });
    },
  };
  await emitFileOfficeOpen(handler, uri, { nativeLoadGeneration: 'generation-1' });
  assert.equal(globalThis.__officeFileOpenMock.readCalls, 0, 'local files must use the bounded Node file handle');
  assert.ok(globalThis.__officeFileOpenMock.statCalls >= 1);
  assert.equal(emitted.length, 1);
  assert.equal(emitted[0].event, 'open');
  assert.equal(emitted[0].payload.bufferBase64, 'AQIDBA==');
  assert.equal(emitted[0].payload.nativeLoadGeneration, 'generation-1');
  assert.equal('buffer' in emitted[0].payload, false);
  assert.equal('error' in emitted[0].payload, false);

  globalThis.__officeFileOpenMock = {
    bytes: new Uint8Array([9]),
    readCalls: 0,
    statCalls: 0,
    statSize: 128 * 1024 * 1024 + 1,
  };
  await truncate(localPath, 128 * 1024 * 1024 + 1);
  emitted.length = 0;
  await emitFileOfficeOpen(handler, uri);
  assert.equal(
    globalThis.__officeFileOpenMock.readCalls,
    0,
    'an oversized local file must be rejected before workspace.fs.readFile or payload allocation',
  );
  assert.equal(emitted.length, 1);
  assert.match(emitted[0].payload.error, /134217728|open limit/i);
  assert.equal('bufferBase64' in emitted[0].payload, false);

  console.log('Local Office open bounds passed: host stat/read cap and base64 transport are enforced.');
} finally {
  delete globalThis.__officeFileOpenMock;
  await rm(buildDirectory, { recursive: true, force: true });
}
