import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import {
  copyFileSync,
  existsSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { spawnSync } from 'node:child_process';
import { basename, join, resolve } from 'node:path';
import { tmpdir } from 'node:os';

if (process.platform !== 'win32') {
  console.log('VBA write-back integration skipped: bundled helper targets Windows x64.');
  process.exit(0);
}

const root = resolve(import.meta.dirname, '..');
const helper = resolve(root, 'bin/win32-x64/excel-ai-vba-writeback.exe');
const fixture = resolve(root, 'test/fixtures/DemoExcelUserForm.xlsm');
const temp = mkdtempSync(join(tmpdir(), 'excel-ai-vba-writeback-test-'));
const workbook = join(temp, 'roundtrip.xlsm');
let requestIndex = 0;

const sha256 = (value) =>
  crypto.createHash('sha256').update(value).digest('hex');

function workbookHash() {
  return sha256(readFileSync(workbook));
}

function runHelper(request, expectedSuccess = true) {
  const requestPath = join(temp, `request-${requestIndex++}.json`);
  writeFileSync(requestPath, JSON.stringify({ schemaVersion: 1, ...request }), 'utf8');
  const result = spawnSync(helper, [requestPath], {
    cwd: temp,
    encoding: 'utf8',
    shell: false,
    windowsHide: true,
  });
  const output = JSON.parse(
    String(result.stdout)
      .trim()
      .split(/\r?\n/)
      .filter(Boolean)
      .at(-1),
  );
  assert.equal(output.ok, expectedSuccess, String(result.stderr || output.message));
  assert.equal(result.status === 0, expectedSuccess, JSON.stringify(output));
  return output;
}

function inspect() {
  return runHelper({ operation: 'inspect', workbookPath: workbook });
}

function excelProcessIds() {
  if (process.platform !== 'win32') return new Set();
  const result = spawnSync(
    'powershell.exe',
    [
      '-NoLogo',
      '-NoProfile',
      '-NonInteractive',
      '-Command',
      'Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }',
    ],
    { encoding: 'utf8', shell: false, windowsHide: true },
  );
  return new Set(
    String(result.stdout)
      .split(/\r?\n/)
      .map((line) => Number.parseInt(line.trim(), 10))
      .filter(Number.isSafeInteger),
  );
}

assert.ok(existsSync(helper), `Bundled helper missing: ${helper}`);
assert.ok(existsSync(fixture), `Fixture missing: ${fixture}`);
copyFileSync(fixture, workbook);
const excelBefore = excelProcessIds();

try {
  const before = inspect();
  const fingerprint = runHelper({
    operation: 'fingerprint',
    workbookPath: workbook,
  });
  assert.match(fingerprint.projectFingerprintSha256, /^[0-9a-f]{64}$/);
  assert.ok(fingerprint.projectStreamCount >= 20);
  assert.ok(fingerprint.projectStorageCount > 0);
  const beforeModules = new Map(before.modules.map((module) => [module.name, module]));
  assert.equal(before.protected, false);
  assert.equal(before.signed, false);
  assert.equal(beforeModules.get('mCode')?.componentKind, 'module');
  assert.equal(beforeModules.get('oUserForm')?.componentKind, 'userform');
  assert.ok(Object.keys(before.designerStreamsSha256).length > 0);

  const originalHash = workbookHash();
  const marker = 'ExcelAiVbaWritebackRoundTrip';
  const moduleSource = [
    'Attribute VB_Name = "mCode"',
    'Option Explicit',
    '',
    `Public Function ${marker}(ByVal query As String) As String`,
    `    ${marker} = "Trouve: " & query`,
    'End Function',
    '',
  ].join('\r\n');
  const applied = runHelper({
    operation: 'apply',
    workbookPath: workbook,
    expectedWorkbookSha256: originalHash,
    patches: [
      {
        moduleName: 'mCode',
        componentKind: 'module',
        source: moduleSource,
      },
    ],
  });
  assert.equal(applied.changed, true);
  assert.deepEqual(applied.modifiedModules, ['mCode']);
  assert.ok(applied.backupPath && existsSync(applied.backupPath));
  assert.equal(basename(applied.backupPath).includes(originalHash.slice(0, 12)), true);
  assert.equal(sha256(readFileSync(applied.backupPath)), originalHash);
  assert.equal(applied.workbookSha256, workbookHash());

  const afterModule = inspect();
  const afterModuleMap = new Map(
    afterModule.modules.map((module) => [module.name, module]),
  );
  assert.match(afterModuleMap.get('mCode').source, new RegExp(marker));
  assert.deepEqual(
    afterModule.designerStreamsSha256,
    before.designerStreamsSha256,
  );
  for (const [name, module] of beforeModules) {
    if (name !== 'mCode') {
      assert.equal(afterModuleMap.get(name)?.sourceSha256, module.sourceSha256);
    }
  }

  const stale = runHelper(
    {
      operation: 'apply',
      workbookPath: workbook,
      expectedWorkbookSha256: originalHash,
      patches: [
        {
          moduleName: 'mCode',
          componentKind: 'module',
          source: moduleSource,
        },
      ],
    },
    false,
  );
  assert.equal(stale.code, 'STALE_WORKBOOK');

  const formModule = afterModuleMap.get('oUserForm');
  const designerSource = [
    'VERSION 5.00',
    'Begin VB.UserForm oUserForm',
    '   Caption = "Designer preserved"',
    'End',
    '',
  ].join('\r\n');
  const formSource = `${designerSource}${formModule.source}\r\nPrivate Sub cmdBonjour_Click()\r\n    MsgBox "Bonjour depuis VS Code", vbInformation\r\nEnd Sub\r\n`;
  const formApplied = runHelper({
    operation: 'apply',
    workbookPath: workbook,
    expectedWorkbookSha256: workbookHash(),
    patches: [
      {
        moduleName: 'oUserForm',
        componentKind: 'userform',
        source: formSource,
        expectedDesignerSha256: sha256(Buffer.from(designerSource, 'utf8')),
      },
    ],
  });
  assert.equal(formApplied.changed, true);
  const afterForm = inspect();
  assert.match(
    afterForm.modules.find((module) => module.name === 'oUserForm').source,
    /Bonjour depuis VS Code/,
  );
  assert.deepEqual(afterForm.designerStreamsSha256, before.designerStreamsSha256);

  const newModuleName = 'mCopilotRecherche';
  const newModuleSource = [
    `Attribute VB_Name = "${newModuleName}"`,
    'Option Explicit',
    '',
    'Public Function RechercheCopilot(ByVal terme As String) As String',
    '    RechercheCopilot = "Résultat: " & terme',
    'End Function',
    '',
  ].join('\r\n');
  const newModuleApplied = runHelper({
    operation: 'apply',
    workbookPath: workbook,
    expectedWorkbookSha256: workbookHash(),
    patches: [
      {
        moduleName: newModuleName,
        componentKind: 'module',
        source: newModuleSource,
      },
    ],
  });
  assert.equal(newModuleApplied.changed, true);
  const afterNewModule = inspect();
  assert.match(
    afterNewModule.modules.find((module) => module.name === newModuleName).source,
    /RechercheCopilot/,
  );
  assert.deepEqual(
    afterNewModule.designerStreamsSha256,
    before.designerStreamsSha256,
  );

  const changedDesigner = runHelper(
    {
      operation: 'apply',
      workbookPath: workbook,
      expectedWorkbookSha256: workbookHash(),
      patches: [
        {
          moduleName: 'oUserForm',
          componentKind: 'userform',
          source: formSource,
          expectedDesignerSha256: '0'.repeat(64),
        },
      ],
    },
    false,
  );
  assert.equal(changedDesigner.code, 'USERFORM_DESIGNER_CHANGED');

  const newForm = runHelper(
    {
      operation: 'apply',
      workbookPath: workbook,
      expectedWorkbookSha256: workbookHash(),
      patches: [
        {
          moduleName: 'UserFormNeverCreate',
          componentKind: 'userform',
          source: formSource.replaceAll('oUserForm', 'UserFormNeverCreate'),
          expectedDesignerSha256: sha256(
            Buffer.from(
              designerSource.replaceAll('oUserForm', 'UserFormNeverCreate'),
              'utf8',
            ),
          ),
        },
      ],
    },
    false,
  );
  assert.equal(newForm.code, 'NEW_USERFORM_UNSUPPORTED');

  const excelAfter = excelProcessIds();
  const newExcelProcesses = [...excelAfter].filter((pid) => !excelBefore.has(pid));
  assert.deepEqual(newExcelProcesses, []);
  console.log('VBA write-back regression passed: existing/new modules + UserForm code, designer preserved, stale writes blocked.');
} finally {
  rmSync(temp, { recursive: true, force: true });
}
