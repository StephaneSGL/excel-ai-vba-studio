import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
for (const script of ['vba-bootstrap-integration.mjs', 'vba-designer-integration.mjs', 'vba-userform-inventory.mjs', 'native-security-preflight.mjs', 'native-xlsm-integration.mjs']) {
  const result = spawnSync(process.execPath, [resolve(root, 'test', script)], {
    cwd: root, stdio: 'inherit', windowsHide: true, timeout: 1_200_000,
    env: { ...process.env, EXCEL_AI_REQUIRE_NATIVE_TESTS: '1' },
  });
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
}
console.log('Required Excel/VBA native acceptance suites passed without prerequisite skips.');
