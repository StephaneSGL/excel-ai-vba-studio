import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

if (process.platform !== 'win32') {
  console.log('VBA UserForm inventory skipped: Windows and Excel are required.');
  process.exit(0);
}

const root = resolve(import.meta.dirname, '..');
const exporter = resolve(root, 'scripts/office-ai-export.ps1');
const workbook = resolve(root, 'test/fixtures/DemoExcelUserForm.xlsm');
const temporaryDirectory = mkdtempSync(
  join(tmpdir(), 'excel-ai-vba-inventory-'),
);
const outputDirectory = join(temporaryDirectory, 'export');
const vbaDirectory = join(outputDirectory, 'vba');
const markdownPath = join(outputDirectory, 'workbook.md');
const jsonPath = join(outputDirectory, 'workbook.json');
const ownerMarker = '.excel-ai-vba-studio-owned';
const ownerMarkerContent =
  'excel-ai-vba-studio:managed-export-directory:v1\n';

function markOwned(directory) {
  mkdirSync(directory, { recursive: true });
  writeFileSync(join(directory, ownerMarker), ownerMarkerContent, 'utf8');
}

try {
  assert.ok(existsSync(workbook), 'VBA UserForm fixture is missing');
  markOwned(temporaryDirectory);
  markOwned(outputDirectory);
  markOwned(vbaDirectory);

  const result = spawnSync(
    'powershell.exe',
    [
      '-NoLogo',
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      exporter,
      '-WorkbookPath',
      workbook,
      '-OutputPath',
      markdownPath,
      '-StorageRoot',
      temporaryDirectory,
      '-JsonOutputPath',
      jsonPath,
      '-VbaOutputDirectory',
      vbaDirectory,
      '-MaxRows',
      '20',
      '-MaxColumns',
      '20',
      '-IncludeVba',
      'true',
    ],
    {
      cwd: root,
      encoding: 'utf8',
      shell: false,
      windowsHide: true,
      timeout: 120_000,
    },
  );
  assert.equal(result.error, undefined, String(result.error));
  assert.equal(result.status, 0, String(result.stderr || result.stdout));

  const exported = JSON.parse(readFileSync(jsonPath, 'utf8'));
  const userForms = exported.workbook?.vba?.userForms;
  assert.ok(Array.isArray(userForms), 'workbook.vba.userForms must be an array');
  const form = userForms.find(({ name }) => name === 'oUserForm');
  assert.ok(form, 'oUserForm inventory is missing');
  assert.ok(form.width > 0 && form.height > 0, 'UserForm geometry is invalid');
  assert.ok(Array.isArray(form.controls), 'UserForm controls must be an array');
  assert.ok(form.controls.length > 0, 'fixture UserForm controls were not inventoried');
  for (const control of form.controls) {
    assert.match(control.name, /^[A-Za-z_][A-Za-z0-9_]{0,30}$/);
    assert.ok(control.width > 0 && control.height > 0);
    assert.ok(typeof control.type === 'string' && control.type.length > 0);
    assert.ok(typeof control.typeName === 'string' && control.typeName.length > 0);
  }

  console.log(
    `VBA UserForm inventory passed: ${form.name}, ${form.controls.length} controls with native geometry.`,
  );
} finally {
  rmSync(temporaryDirectory, { recursive: true, force: true });
}
