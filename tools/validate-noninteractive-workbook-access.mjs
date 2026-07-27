import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const security = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/security.ts'),
  'utf8',
);
const workbookService = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/workbookService.ts'),
  'utf8',
);
const manifest = JSON.parse(readFileSync(resolve(root, 'package.json'), 'utf8'));

assert.match(security, /export async function canonicalizeWorkbookUri/);
assert.match(security, /await assertLocalDrive\(uri\.fsPath\)/);
assert.match(security, /await fs\.promises\.realpath\(uri\.fsPath\)/);
assert.match(security, /await assertLocalDrive\(canonicalPath\)/);
assert.doesNotMatch(
  security,
  /confirmExactPath|authorizeWorkbookRead|isWorkbookInsideWorkspace|ALLOW_ONCE/,
);
assert.match(workbookService, /canonicalizeWorkbookUri/);
assert.doesNotMatch(workbookService, /authorizeWorkbookRead/);
assert.equal(
  manifest.capabilities?.untrustedWorkspaces?.supported,
  false,
  'native workbook automation must remain disabled in untrusted workspaces',
);

console.log(
  'Workbook access validation passed: no per-file prompt; canonical local-path and Workspace Trust checks remain.',
);
