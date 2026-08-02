import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';
import process from 'node:process';

if (process.platform !== 'win32') {
  console.log('Native XLSM integration skipped: Windows and Microsoft Excel are required.');
  process.exit(0);
}

const root = resolve(import.meta.dirname, '..');

function runPowerShell(script, args = []) {
  const result = spawnSync(
    'powershell.exe',
    [
      '-NoLogo',
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      resolve(root, 'test', script),
      ...args,
    ],
    {
      cwd: root,
      encoding: 'utf8',
      shell: false,
      // Each native edit is deliberately isolated in its own Excel COM
      // transaction. The expanded workbook-object suite can exceed three
      // minutes on otherwise healthy machines, so retain a bounded but
      // realistic integration-test deadline.
      timeout: 600_000,
    },
  );

  process.stdout.write(result.stdout ?? '');
  process.stderr.write(result.stderr ?? '');
  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}

runPowerShell('run-native-edit-tests.ps1', [
  '-WorkbookPath',
  resolve(root, 'test', 'fixtures', 'NativeEditingSynthetic.xlsm'),
]);
runPowerShell('run-vba-preservation-test.ps1', [
  '-WorkbookPath',
  resolve(root, 'test', 'fixtures', 'DemoExcelUserForm.xlsm'),
]);
