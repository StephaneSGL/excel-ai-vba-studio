import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

execFileSync(
  process.platform === 'win32' ? 'python' : 'python3',
  ['-B', fileURLToPath(new URL('./atomic-writeback-regression.py', import.meta.url))],
  { stdio: 'inherit' },
);
