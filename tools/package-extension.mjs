import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import process from 'node:process';
import { requestedTargets, withTargetIgnore } from './package-targets.mjs';

const root = resolve(import.meta.dirname, '..');
const manifest = JSON.parse(
  readFileSync(resolve(root, 'package.json'), 'utf8'),
);
const outputDirectory = resolve(root, 'output', 'vsix');
const vsceEntrypoint = resolve(
  root,
  'node_modules',
  '@vscode',
  'vsce',
  'vsce',
);

mkdirSync(outputDirectory, { recursive: true });

for (const target of requestedTargets()) {
  const outputPath = resolve(outputDirectory, `${manifest.name}-${target}-${manifest.version}.vsix`);
  const result = withTargetIgnore(root, target, ignoreFile => spawnSync(
    process.execPath,
    [vsceEntrypoint, 'package', '--target', target, '--no-dependencies', '--ignoreFile', ignoreFile, '--out', outputPath],
    { cwd: root, stdio: 'inherit', shell: false },
  ));
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
}
