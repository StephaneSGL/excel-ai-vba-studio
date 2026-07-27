import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import process from 'node:process';

const root = resolve(import.meta.dirname, '..');
const manifest = JSON.parse(
  readFileSync(resolve(root, 'package.json'), 'utf8'),
);
const outputDirectory = resolve(root, 'output', 'vsix');
const outputPath = resolve(
  outputDirectory,
  `${manifest.name}-win32-x64-${manifest.version}.vsix`,
);
const vsceEntrypoint = resolve(
  root,
  'node_modules',
  '@vscode',
  'vsce',
  'vsce',
);

mkdirSync(outputDirectory, { recursive: true });

const result = spawnSync(
  process.execPath,
  [
    vsceEntrypoint,
    'package',
    '--target',
    'win32-x64',
    '--no-dependencies',
    '--out',
    outputPath,
  ],
  {
    cwd: root,
    stdio: 'inherit',
    shell: false,
  },
);

if (result.error) {
  throw result.error;
}
process.exit(result.status ?? 1);
