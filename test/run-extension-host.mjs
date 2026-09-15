import { runTests } from '@vscode/test-electron';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const temporary = await mkdtemp(join(tmpdir(), 'excel-host-smoke-'));
try {
  await runTests({
    version: process.argv[2] || 'stable',
    extensionDevelopmentPath: root,
    extensionTestsPath: join(root, 'test/extension-host-smoke.cjs'),
    extensionTestsEnv: { EXCEL_AI_SMOKE_WORKSPACE: temporary, ELECTRON_RUN_AS_NODE: undefined },
    launchArgs: [temporary, '--disable-extensions', '--disable-gpu', '--skip-welcome', '--skip-release-notes',
      '--disable-workspace-trust', `--user-data-dir=${join(temporary, 'profile')}`, `--extensions-dir=${join(temporary, 'extensions')}`],
  });
} finally {
  await rm(temporary, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
}
