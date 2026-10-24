import { execFileSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const testDirectory = dirname(fileURLToPath(import.meta.url));
execFileSync(
  process.platform === 'win32' ? 'python' : 'python3',
  ['-B', join(testDirectory, 'native-zip-budget-regression.py')],
  { stdio: 'inherit' },
);
