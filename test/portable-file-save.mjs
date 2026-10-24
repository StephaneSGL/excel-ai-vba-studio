import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { build } from 'esbuild';
import { mkdtemp, mkdir, readFile, writeFile, chmod, stat, symlink, link, readdir, realpath, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const root = resolve(import.meta.dirname, '..');
const temp = await mkdtemp(join(tmpdir(), 'excel-portable-save-'));
const digest = value => createHash('sha256').update(value).digest('hex');

function windowsShortPath(directory) {
  const result = spawnSync('cmd.exe', ['/d', '/s', '/c', 'for %I in ("%EXCEL_SYNTHETIC_DIRECTORY%") do @echo %~sI'], {
    encoding: 'utf8', timeout: 10_000, windowsHide: true, windowsVerbatimArguments: true,
    env: { ...process.env, EXCEL_SYNTHETIC_DIRECTORY: directory },
  });
  assert.ifError(result.error);
  assert.equal(result.status, 0, `Synthetic Windows short-path query failed: ${result.stderr}`);
  assert.ok(result.stdout.trim(), 'Windows short-path query must return a path');
  return result.stdout.trim();
}

function syntheticAttributes(file, action) {
  const result = spawnSync('python3', ['-c', `
import json, os, stat, struct, sys
path, action = sys.argv[1:]
if action == 'set':
    # st_mode reports ACL mask=r although the owning group has no access.
    entries = [(1, 6, 0xffffffff), (2, 4, os.getuid() + 1),
               (4, 0, 0xffffffff), (16, 4, 0xffffffff), (32, 0, 0xffffffff)]
    acl = struct.pack('<I', 2) + b''.join(struct.pack('<HHI', *entry) for entry in entries)
    os.setxattr(path, 'system.posix_acl_access', acl)
    os.setxattr(path, 'user.excel-synthetic-test', b'preserve synthetic metadata')
info = os.stat(path)
print(json.dumps({'mode': stat.S_IMODE(info.st_mode), 'uid': info.st_uid, 'gid': info.st_gid,
                  'acl': os.getxattr(path, 'system.posix_acl_access').hex(),
                  'xattr': os.getxattr(path, 'user.excel-synthetic-test').hex()}))
`, file, action], { encoding: 'utf8', timeout: 10_000, maxBuffer: 8192 });
  assert.ifError(result.error);
  assert.equal(result.status, 0, `Synthetic ACL fixture failed: ${result.stderr}`);
  return JSON.parse(result.stdout);
}

async function loadSaveWithSpawnWrapper(temp, behavior) {
  const output = join(temp, `save-${behavior}.mjs`);
  await build({
    entryPoints: [join(root, 'src/provider/portableFileSave.ts')], outfile: output,
    bundle: true, format: 'esm', platform: 'node', logLevel: 'silent',
    plugins: [{ name: 'synthetic-cp-boundary', setup(api) {
      api.onResolve({ filter: /^node:child_process$/ }, () => ({ path: 'child-process', namespace: 'test' }));
      api.onLoad({ filter: /.*/, namespace: 'test' }, () => ({ loader: 'js', contents: `
        import { createRequire } from 'node:module';
        import { fchmodSync, fstatSync } from 'node:fs';
        const realSpawn = createRequire(import.meta.url)('node:child_process').spawn;
        export function spawn(command, args, options) {
          const child = realSpawn(${behavior === 'unavailable' ? "'/nonexistent-excel-synthetic-cp'" : 'command'}, args, options);
          ${behavior === 'metadata-race' ? `child.once('close', () => {
            // Change ctime without changing bytes, uid/gid or the mode value.
            const fd = options.stdio[3];
            fchmodSync(fd, fstatSync(fd).mode & 0o777);
          });` : ''}
          return child;
        }
      ` }));
    } }],
  });
  return import(pathToFileURL(output).href);
}
try {
  const compiled = join(temp, 'save.mjs');
  await build({ entryPoints: [join(root, 'src/provider/portableFileSave.ts')], outfile: compiled,
    bundle: true, format: 'esm', platform: 'node', logLevel: 'silent' });
  const { createPortableFile, savePortableFile } = await import(pathToFileURL(compiled).href);
  const file = join(temp, 'synthetic.xlsx');
  const original = Buffer.from('synthetic original bytes');
  const edited = Buffer.from('synthetic edited bytes');
  assert.equal(await createPortableFile(file, original), digest(original));
  await assert.rejects(createPortableFile(file, edited), /already exists/);
  assert.deepEqual(await readFile(file), original);
  // Linux uses GNU cp for ACL/xattr preservation; Windows keeps its own writer.
  if (process.platform === 'linux') {
    await chmod(file, 0o640);
    assert.equal(await savePortableFile(file, edited, digest(original)), digest(edited));
    assert.deepEqual(await readFile(file), edited);
    assert.equal((await stat(file)).mode & 0o777, 0o640);
    const backups = join(temp, '.excel-ai-vba-backups');
    const recovery = await readdir(backups);
    assert.equal(recovery.length, 1);
    assert.deepEqual(await readFile(join(backups, recovery[0])), original);
    await assert.rejects(savePortableFile(file, original, digest(original)), /changed outside/);
    assert.deepEqual(await readFile(file), edited);
    await assert.rejects(savePortableFile(file, original, ''), /snapshot is missing/);
    await chmod(file, 0o444);
    await assert.rejects(savePortableFile(file, original, digest(edited)), /read-only/);
    await chmod(file, 0o640);
    const alias = join(temp, 'alias.xlsx');
    await symlink(file, alias);
    await assert.rejects(savePortableFile(alias, original, digest(edited)), /symbolic or hard links/);
    const hard = join(temp, 'hard.xlsx');
    await link(file, hard);
    await assert.rejects(savePortableFile(file, original, digest(edited)), /symbolic or hard links/);
    await rm(hard);
    await assert.rejects(createPortableFile(join(backups, 'new.xlsx'), edited), /Recovery backups/);
    const directoryLink = join(temp, 'linked-parent');
    await symlink(temp, directoryLink);
    await assert.rejects(createPortableFile(join(directoryLink, 'new.xlsx'), edited), /symbolic links/);
    // Concurrent edits from two editor sessions cannot both replace the baseline.
    const results = await Promise.allSettled([
      savePortableFile(file, Buffer.from('session one'), digest(edited)),
      savePortableFile(file, Buffer.from('session two'), digest(edited)),
    ]);
    assert.ok(results.some(result => result.status === 'fulfilled'));
    assert.ok(results.some(result => result.status === 'rejected'), 'stale parallel save must be rejected');

    const aclFile = join(temp, 'synthetic-acl.xlsx');
    await createPortableFile(aclFile, original);
    const attributes = syntheticAttributes(aclFile, 'set');
    assert.equal(attributes.mode, 0o640, 'the ACL mask must reproduce the old chmod-only widening case');
    assert.equal(await savePortableFile(aclFile, edited, digest(original)), digest(edited));
    assert.deepEqual(syntheticAttributes(aclFile, 'inspect'), attributes,
      'real GNU cp must preserve exact ACL/xattr bytes and ownership');
    assert.deepEqual(await readFile(aclFile), edited, 'attribute copying must preserve replacement content');

    const race = await loadSaveWithSpawnWrapper(temp, 'metadata-race');
    const beforeRace = await stat(aclFile, { bigint: true });
    await assert.rejects(race.savePortableFile(aclFile, original, digest(edited)), /changed during saving/);
    const afterRace = await stat(aclFile, { bigint: true });
    assert.equal(afterRace.mtimeNs, beforeRace.mtimeNs);
    assert.notEqual(afterRace.ctimeNs, beforeRace.ctimeNs, 'test must exercise a metadata-only change');
    assert.deepEqual(await readFile(aclFile), edited, 'metadata drift must leave the source content intact');

    const unavailable = await loadSaveWithSpawnWrapper(temp, 'unavailable');
    await assert.rejects(unavailable.savePortableFile(aclFile, original, digest(edited)), /could not preserve filesystem/);
    assert.deepEqual(await readFile(aclFile), edited, 'missing attribute helper must fail closed');
  } else {
    await assert.rejects(savePortableFile(file, edited, digest(original)), /qualified only on Linux/);
    if (process.platform === 'win32') {
      const canonicalTemp = await realpath(temp);
      const directory = join(canonicalTemp, 'SyntheticCaseParent');
      await mkdir(directory);
      const caseAlias = join(canonicalTemp, 'sYNTHETICcASEpARENT');
      const preservedName = 'PreserveCase.xlsx';
      assert.equal(await createPortableFile(join(caseAlias, preservedName), edited), digest(edited));
      assert.deepEqual(await readFile(join(directory, preservedName)), edited);
      assert.ok((await readdir(directory)).includes(preservedName), 'canonicalizing a parent must preserve the requested filename');

      const shortDirectory = windowsShortPath(directory);
      assert.equal(await realpath(shortDirectory), await realpath(directory));
      assert.equal(await createPortableFile(join(shortDirectory, 'short-alias.xlsx'), original), digest(original));
      assert.deepEqual(await readFile(join(directory, 'short-alias.xlsx')), original);
      console.log(`Windows synthetic paths: ${JSON.stringify({ lexicalTemp: temp, canonicalTemp, shortDirectory })}`);

      const nested = join(directory, 'Nested');
      await mkdir(nested);
      const junction = join(canonicalTemp, 'synthetic-junction');
      await symlink(directory, junction, 'junction');
      await assert.rejects(createPortableFile(join(junction, 'Nested', 'refused.xlsx'), edited), /symbolic links or junctions/);
      assert.deepEqual(await readdir(nested), [], 'an ancestor junction must not redirect the destination');

      const backups = join(canonicalTemp, '.EXCEL-AI-VBA-BACKUPS');
      await mkdir(backups);
      await assert.rejects(createPortableFile(join(backups, 'new.xlsx'), edited), /Recovery backups/);
      const shortBackups = windowsShortPath(backups);
      await assert.rejects(createPortableFile(join(shortBackups, 'alias.xlsx'), edited), /Recovery backups/);
      assert.deepEqual(await readdir(backups), [], 'recovery paths must remain protected through case and 8.3 aliases');
    }
  }
  assert.deepEqual((await readdir(temp)).filter(name => name.startsWith('.excel-ai-save-')), []);
  console.log('Portable file save passed: exclusive create, staged replacement, recovery, ACL/xattr preservation, metadata-race and stale/link/read-only refusals.');
} finally {
  await rm(temp, { recursive: true, force: true });
}
