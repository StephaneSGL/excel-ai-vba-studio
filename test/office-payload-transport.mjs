import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const root = resolve(import.meta.dirname, '..');
const directory = await mkdtemp(join(tmpdir(), 'office-payload-transport-'));
const output = join(directory, 'load-office-content.mjs');

try {
  await build({
    entryPoints: [join(root, 'src/react/util/loadOfficeContent.ts')],
    outfile: output,
    bundle: true,
    format: 'esm',
    platform: 'browser',
    logLevel: 'silent',
  });
  const { arrayBufferFromBase64Payload } = await import(
    `${pathToFileURL(output).href}?cache=${Date.now()}`
  );
  const source = Uint8Array.from({ length: 2_500_003 }, (_, index) => index % 251);
  const encoded = Buffer.from(source).toString('base64');
  const restored = new Uint8Array(arrayBufferFromBase64Payload(encoded));
  assert.deepEqual(restored, source);
  assert.throws(() => arrayBufferFromBase64Payload('not base64'), /Invalid/);
  assert.throws(
    () => arrayBufferFromBase64Payload({ length: 180_000_004 }),
    /oversized/,
  );
  console.log('Office payload transport passed: bounded base64 round-trip without number[] expansion.');
} finally {
  await rm(directory, { recursive: true, force: true });
}
