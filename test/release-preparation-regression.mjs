import assert from 'node:assert/strict';
import { copyFile, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

const temporary = await mkdtemp(join(tmpdir(), 'excel-release-prepare-'));
try {
  await mkdir(join(temporary, 'tools'));
  const runner = join(temporary, 'tools/prepare-release.mjs');
  await copyFile(resolve(import.meta.dirname, '../tools/prepare-release.mjs'), runner);
  const manifest = JSON.stringify({ version: '0.1.0' }, null, 2);
  await writeFile(join(temporary, 'package.json'), manifest);
  await writeFile(join(temporary, 'package-lock.json'), JSON.stringify({ version: '0.0.9' }));
  const run = () => spawnSync(process.execPath, [runner, '0.1.1'], { encoding: 'utf8', timeout: 10_000, windowsHide: true });
  assert.notEqual(run().status, 0, 'mismatched lockfile refused');
  assert.equal(await readFile(join(temporary, 'package.json'), 'utf8'), manifest, 'manifest unchanged on validation failure');
  await writeFile(join(temporary, 'package-lock.json'), JSON.stringify({ version: '0.1.0', packages: { '': { version: '0.1.0' } } }, null, 2));
  const result = run();
  assert.equal(result.status, 0, result.stderr);
  const lock = JSON.parse(await readFile(join(temporary, 'package-lock.json'), 'utf8'));
  assert.equal(lock.version, '0.1.1');
  assert.equal(lock.packages[''].version, '0.1.1');
} finally {
  await rm(temporary, { recursive: true, force: true });
}
console.log('Release preparation regression passed.');
