import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import { spawnSync } from 'node:child_process';
import {
  copyFileSync,
  existsSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs';
import { join, resolve } from 'node:path';
import { tmpdir } from 'node:os';

function normalizeDiagnostic(value) {
  return String(value).replace(/\s+/g, ' ').trim();
}

assert.equal(
  normalizeDiagnostic('already\r\n exists;\tset replaceExisting=true'),
  'already exists; set replaceExisting=true',
);
assert.equal(
  normalizeDiagnostic('Excel refused\r\n ActiveX insertion'),
  'Excel refused ActiveX insertion',
);

if (process.platform !== 'win32') {
  console.log('VBA designer integration skipped: Windows and Excel are required.');
  process.exit(0);
}

const root = resolve(import.meta.dirname, '..');
const engine = resolve(root, 'scripts/apply-vba-designer.ps1');
const helper = resolve(root, 'bin/win32-x64/excel-ai-vba-writeback.exe');
const fixture = resolve(root, 'test/fixtures/DemoExcelUserForm.xlsm');
const temporaryDirectory = mkdtempSync(
  join(tmpdir(), 'excel-ai-vba-designer-test-'),
);
let helperRequestIndex = 0;

const base64 = (value) => Buffer.from(value, 'utf8').toString('base64');
const sha256File = (filePath) =>
  crypto.createHash('sha256').update(readFileSync(filePath)).digest('hex');

function excelProcessIds() {
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

function excelAutomationReady() {
  const probe = [
    "$ErrorActionPreference = 'Stop'",
    '$excel = $null',
    '$workbook = $null',
    '$ready = $false',
    'try {',
    '  $excel = New-Object -ComObject Excel.Application',
    '  $excel.AutomationSecurity = 3',
    '  $excel.DisplayAlerts = $false',
    '  $excel.EnableEvents = $false',
    '  $excel.Visible = $false',
    '  $workbook = $excel.Workbooks.Add()',
    '  [void]$workbook.VBProject.VBComponents.Count',
    '  $ready = $true',
    '} catch {',
    '  $ready = $false',
    '} finally {',
    '  if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }',
    '  if ($null -ne $excel) { try { $excel.Quit() } catch {} }',
    '  if ($null -ne $workbook -and [Runtime.InteropServices.Marshal]::IsComObject($workbook)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) }',
    '  if ($null -ne $excel -and [Runtime.InteropServices.Marshal]::IsComObject($excel)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }',
    '  [GC]::Collect()',
    '  [GC]::WaitForPendingFinalizers()',
    '}',
    'if ($ready) { exit 0 } else { exit 2 }',
  ].join('\n');
  const result = spawnSync(
    'powershell.exe',
    ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', probe],
    { encoding: 'utf8', shell: false, windowsHide: true, timeout: 30_000 },
  );
  return result.status === 0;
}

function activeXCreationReady() {
  const probe = [
    "$ErrorActionPreference = 'Stop'",
    '$excel = $null',
    '$workbook = $null',
    '$worksheet = $null',
    '$objects = $null',
    '$control = $null',
    '$ready = $false',
    'try {',
    '  $excel = New-Object -ComObject Excel.Application',
    '  $excel.AutomationSecurity = 3',
    '  $excel.DisplayAlerts = $false',
    '  $excel.EnableEvents = $false',
    '  $excel.Visible = $false',
    '  $workbook = $excel.Workbooks.Add()',
    '  $worksheet = $workbook.Worksheets.Item(1)',
    '  $objects = $worksheet.OLEObjects()',
    "  $control = $objects.Add('Forms.CommandButton.1', [Type]::Missing, $false, $false, [Type]::Missing, [Type]::Missing, [Type]::Missing, 20, 20, 120, 28)",
    '  $ready = $null -ne $control',
    '} catch {',
    '  $ready = $false',
    '} finally {',
    '  if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }',
    '  if ($null -ne $excel) { try { $excel.Quit() } catch {} }',
    '  foreach ($item in @($control, $objects, $worksheet, $workbook, $excel)) {',
    '    if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item) } catch {} }',
    '  }',
    '  [GC]::Collect()',
    '  [GC]::WaitForPendingFinalizers()',
    '}',
    'if ($ready) { exit 0 } else { exit 2 }',
  ].join('\n');
  const result = spawnSync(
    'powershell.exe',
    ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', probe],
    { encoding: 'utf8', shell: false, windowsHide: true, timeout: 30_000 },
  );
  return result.status === 0;
}

function runDesigner(requestPath) {
  return spawnSync(
    'powershell.exe',
    [
      '-NoLogo',
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      engine,
      '-RequestPathBase64',
      base64(requestPath),
      '-HelperPathBase64',
      base64(helper),
    ],
    {
      cwd: temporaryDirectory,
      encoding: 'utf8',
      shell: false,
      windowsHide: true,
      timeout: 120_000,
    },
  );
}

function parseLastJson(stdout) {
  const line = String(stdout)
    .replace(/\r/g, '')
    .split('\n')
    .map((value) => value.trim())
    .filter(Boolean)
    .at(-1);
  assert.ok(line, 'designer did not return a JSON result');
  return JSON.parse(line);
}

function inspectWorkbook(workbookPath) {
  const requestPath = join(
    temporaryDirectory,
    `inspect-${helperRequestIndex++}.json`,
  );
  writeFileSync(
    requestPath,
    JSON.stringify({
      schemaVersion: 1,
      operation: 'inspect',
      workbookPath,
    }),
    'utf8',
  );
  const result = spawnSync(helper, [requestPath], {
    cwd: temporaryDirectory,
    encoding: 'utf8',
    shell: false,
    windowsHide: true,
    timeout: 30_000,
  });
  assert.equal(result.status, 0, String(result.stderr || result.stdout));
  return parseLastJson(result.stdout);
}

function assertNoNewExcelProcess(before) {
  const after = excelProcessIds();
  const residual = [...after].filter((processId) => !before.has(processId));
  assert.deepEqual(residual, [], `residual Excel process IDs: ${residual}`);
}

assert.ok(existsSync(engine), 'VBA designer engine is missing');
assert.ok(existsSync(helper), 'native helper is missing');
assert.ok(statSync(helper).size > 1_000_000, 'native helper is too small');
assert.ok(existsSync(fixture), 'UserForm fixture is missing');

if (!excelAutomationReady()) {
  rmSync(temporaryDirectory, { recursive: true, force: true });
  console.log(
    'VBA designer integration skipped: Excel COM or user-enabled AccessVBOM is unavailable.',
  );
  process.exit(0);
}

const activeXAvailable = activeXCreationReady();
const excelBefore = excelProcessIds();

try {
  const workbookPath = join(temporaryDirectory, 'designer-success.xlsm');
  copyFileSync(fixture, workbookPath);
  const originalHash = sha256File(workbookPath);
  const generatedControls = [
    {
      type: 'label',
      name: 'lblTitle',
      caption: 'Generated safely',
      left: 18,
      top: 16,
      width: 180,
      height: 20,
    },
    {
      type: 'textBox',
      name: 'txtValue',
      left: 18,
      top: 48,
      width: 130,
      height: 22,
    },
    {
      type: 'commandButton',
      name: 'cmdClose',
      caption: 'Close',
      left: 164,
      top: 46,
      width: 90,
      height: 28,
    },
    {
      type: 'comboBox',
      name: 'cboChoice',
      left: 18,
      top: 84,
      width: 130,
      height: 22,
    },
    {
      type: 'listBox',
      name: 'lstChoice',
      left: 164,
      top: 84,
      width: 130,
      height: 54,
    },
    {
      type: 'checkBox',
      name: 'chkEnabled',
      caption: 'Enabled',
      left: 18,
      top: 122,
      width: 100,
      height: 20,
    },
    {
      type: 'optionButton',
      name: 'optMode',
      caption: 'Mode',
      left: 18,
      top: 150,
      width: 100,
      height: 20,
    },
    {
      type: 'toggleButton',
      name: 'tglState',
      caption: 'Toggle',
      left: 164,
      top: 150,
      width: 90,
      height: 26,
    },
    {
      type: 'frame',
      name: 'fraGroup',
      caption: 'Group',
      left: 18,
      top: 188,
      width: 130,
      height: 70,
    },
    {
      type: 'image',
      name: 'imgPreview',
      left: 164,
      top: 188,
      width: 70,
      height: 60,
    },
    {
      type: 'spinButton',
      name: 'spnValue',
      left: 18,
      top: 276,
      width: 20,
      height: 48,
    },
    {
      type: 'scrollBar',
      name: 'scrValue',
      left: 54,
      top: 286,
      width: 200,
      height: 20,
    },
  ];
  const keyDownProcedure = [
    'Private Sub txtGenerated_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)',
    '    If KeyCode = vbKeyReturn Then',
    '        Me.Caption = "Validated"',
    '    End If',
    'End Sub',
  ].join('\r\n');
  const initializeProcedure = [
    'Private Sub UserForm_Initialize()',
    '    Me.Caption = "Ready"',
    'End Sub',
  ].join('\r\n');
  const replacedInitializeProcedure = [
    'Private Sub UserForm_Initialize()',
    '    Me.Caption = "Ready after replacement"',
    'End Sub',
  ].join('\r\n');
  const requestPath = join(temporaryDirectory, 'designer-success.json');
  writeFileSync(
    requestPath,
    JSON.stringify({
      schemaVersion: 2,
      workbookPath,
      expectedWorkbookSha256: originalHash,
      allowedCustomActiveXProgIds: ['Forms.ToggleButton.1'],
      operations: [
        {
          kind: 'createUserForm',
          name: 'frmGenerated',
          caption: 'Generated form',
          width: 420,
          height: 380,
          source: [
            'Option Explicit',
            '',
            'Private Sub cmdClose_Click()',
            '    Unload Me',
            'End Sub',
          ].join('\r\n'),
          controls: generatedControls,
        },
        {
          kind: 'addUserFormControl',
          formName: 'oUserForm',
          control: {
            type: 'textBox',
            name: 'txtGenerated',
            left: 18,
            top: 190,
            width: 130,
            height: 22,
          },
        },
        {
          kind: 'updateUserFormControl',
          formName: 'oUserForm',
          name: 'txtGenerated',
          changes: {
            left: 24,
            top: 196,
            width: 156,
            height: 24,
            enabled: true,
            visible: true,
            tabIndex: 0,
            controlTipText: 'Generated editor',
          },
        },
        {
          kind: 'setUserFormEventHandler',
          formName: 'oUserForm',
          objectName: 'txtGenerated',
          eventName: 'KeyDown',
          procedureSource: keyDownProcedure,
        },
        {
          kind: 'setUserFormEventHandler',
          formName: 'frmGenerated',
          objectName: 'UserForm',
          eventName: 'Initialize',
          procedureSource: initializeProcedure,
        },
        {
          kind: 'createWorksheetButton',
          sheetName: 'Data',
          name: 'btnGenerated',
          caption: 'Run generated',
          macroName: 'mcode.showuserform',
          left: 20,
          top: 20,
          width: 120,
          height: 28,
        },
        ...(activeXAvailable ? [{
          kind: 'createWorksheetActiveXControl',
          sheetName: 'Data',
          control: {
            type: 'commandButton',
            name: 'btnActiveX',
            caption: 'Open form',
            left: 160,
            top: 20,
            width: 120,
            height: 28,
          },
        },
        {
          kind: 'bindWorksheetActiveXMacro',
          sheetName: 'Data',
          name: 'btnActiveX',
          macroName: 'mCode.ShowUserForm',
        },
        {
          kind: 'createWorksheetActiveXControl',
          sheetName: 'Data',
          control: {
            type: 'customActiveX',
            progId: 'Forms.ToggleButton.1',
            name: 'tglAllowlisted',
            caption: 'Allowed custom',
            left: 300,
            top: 20,
            width: 120,
            height: 28,
          },
        }] : []),
      ],
    }),
    'utf8',
  );

  const result = runDesigner(requestPath);
  assert.equal(result.error, undefined, String(result.error));
  assert.equal(result.status, 0, String(result.stderr || result.stdout));
  const output = parseLastJson(result.stdout);
  assert.equal(output.ok, true);
  assert.equal(output.targetWorkbookPath, workbookPath);
  assert.equal(output.sourceWorkbookPath, workbookPath);
  assert.equal(output.convertedToXlsm, false);
  assert.equal(output.changed, true);
  assert.deepEqual(output.createdUserForms, ['frmGenerated']);
  assert.deepEqual(output.addedControls, [
    ...generatedControls.map(
      control => `frmGenerated.${control.name}`,
    ),
    'oUserForm.txtGenerated',
  ]);
  assert.deepEqual(output.updatedControls, ['oUserForm.txtGenerated']);
  assert.deepEqual(output.updatedEventHandlers, [
    'oUserForm.txtGenerated_KeyDown',
    'frmGenerated.UserForm_Initialize',
  ]);
  assert.deepEqual(output.createdButtons, ['Data.btnGenerated']);
  assert.deepEqual(output.assignedButtons, []);
  assert.deepEqual(
    output.createdActiveXControls,
    activeXAvailable ? ['Data.btnActiveX', 'Data.tglAllowlisted'] : [],
  );
  assert.deepEqual(
    output.boundActiveXControls,
    activeXAvailable ? ['Data.btnActiveX'] : [],
  );
  assert.equal(output.macrosExecuted, false);
  assert.equal(output.accessVbomChanged, false);
  assert.equal(output.designerVerified, true);
  assert.match(output.workbookSha256, /^[0-9a-f]{64}$/);
  assert.equal(sha256File(workbookPath), output.workbookSha256);
  assert.ok(existsSync(output.backupPath));
  assert.equal(sha256File(output.backupPath), originalHash);

  const inspection = inspectWorkbook(workbookPath);
  assert.equal(inspection.ok, true);
  assert.equal(inspection.workbookSha256, output.workbookSha256);
  const modules = new Map(
    inspection.modules.map((module) => [module.name, module]),
  );
  assert.equal(modules.get('frmGenerated')?.componentKind, 'userform');
  assert.equal(modules.get('oUserForm')?.componentKind, 'userform');
  assert.match(modules.get('frmGenerated')?.source || '', /Ready/);
  assert.match(modules.get('oUserForm')?.source || '', /txtGenerated_KeyDown/);
  const designerEntries = Object.entries(inspection.designerStreamsSha256);
  assert.ok(
    designerEntries.some(
      ([name, digest]) =>
        name.startsWith('frmgenerated/') && /^[0-9a-f]{64}$/.test(digest),
    ),
  );

  const assignmentHash = sha256File(workbookPath);
  const assignmentRequestPath = join(
    temporaryDirectory,
    'designer-assignment.json',
  );
  writeFileSync(
    assignmentRequestPath,
    JSON.stringify({
      schemaVersion: 2,
      workbookPath,
      expectedWorkbookSha256: assignmentHash,
      operations: [
        {
          kind: 'assignWorksheetButtonMacro',
          sheetName: 'Data',
          name: 'btnGenerated',
          macroName: 'ShowUserForm',
        },
        {
          kind: 'setUserFormEventHandler',
          formName: 'frmGenerated',
          objectName: 'UserForm',
          eventName: 'Initialize',
          procedureSource: replacedInitializeProcedure,
          replaceExisting: true,
        },
      ],
    }),
    'utf8',
  );
  const assignmentResult = runDesigner(assignmentRequestPath);
  assert.equal(
    assignmentResult.status,
    0,
    String(assignmentResult.stderr || assignmentResult.stdout),
  );
  const assignmentOutput = parseLastJson(assignmentResult.stdout);
  assert.deepEqual(assignmentOutput.createdButtons, []);
  assert.deepEqual(assignmentOutput.assignedButtons, ['Data.btnGenerated']);
  assert.deepEqual(assignmentOutput.updatedControls, []);
  assert.deepEqual(assignmentOutput.updatedEventHandlers, [
    'frmGenerated.UserForm_Initialize',
  ]);
  assert.deepEqual(assignmentOutput.createdActiveXControls, []);
  assert.deepEqual(assignmentOutput.boundActiveXControls, []);
  assert.equal(sha256File(workbookPath), assignmentOutput.workbookSha256);
  assert.equal(sha256File(assignmentOutput.backupPath), assignmentHash);

  const existingUserFormEventHash = sha256File(workbookPath);
  const existingUserFormEventRequestPath = join(
    temporaryDirectory,
    'designer-existing-userform-event.json',
  );
  writeFileSync(
    existingUserFormEventRequestPath,
    JSON.stringify({
      schemaVersion: 2,
      workbookPath,
      expectedWorkbookSha256: existingUserFormEventHash,
      operations: [
        {
          kind: 'setUserFormEventHandler',
          formName: 'frmGenerated',
          objectName: 'UserForm',
          eventName: 'Initialize',
          procedureSource: initializeProcedure,
        },
      ],
    }),
    'utf8',
  );
  const existingUserFormEventResult = runDesigner(
    existingUserFormEventRequestPath,
  );
  assert.notEqual(
    existingUserFormEventResult.status,
    0,
    'existing UserForm event replacement must require explicit opt-in',
  );
  assert.match(
    normalizeDiagnostic(existingUserFormEventResult.stderr),
    /already exists; set replaceExisting=true/i,
  );
  assert.equal(
    sha256File(workbookPath),
    existingUserFormEventHash,
    'existing UserForm event refusal changed the workbook',
  );

  if (activeXAvailable) {
  const existingHandlerHash = sha256File(workbookPath);
  const existingHandlerRequestPath = join(
    temporaryDirectory,
    'designer-existing-handler.json',
  );
  writeFileSync(
    existingHandlerRequestPath,
    JSON.stringify({
      schemaVersion: 2,
      workbookPath,
      expectedWorkbookSha256: existingHandlerHash,
      operations: [
        {
          kind: 'bindWorksheetActiveXMacro',
          sheetName: 'Data',
          name: 'btnActiveX',
          macroName: 'mCode.ShowUserForm',
        },
      ],
    }),
    'utf8',
  );
  const existingHandlerResult = runDesigner(existingHandlerRequestPath);
  assert.notEqual(
    existingHandlerResult.status,
    0,
    'existing ActiveX handler replacement must fail',
  );
  assert.match(
    String(existingHandlerResult.stderr),
    /event handler 'btnActiveX_Click' already exists/i,
  );
  assert.equal(
    sha256File(workbookPath),
    existingHandlerHash,
    'existing-handler refusal changed the workbook',
  );
  }

  const blockedCustomWorkbookPath = join(
    temporaryDirectory,
    'designer-blocked-custom.xlsm',
  );
  copyFileSync(fixture, blockedCustomWorkbookPath);
  const blockedCustomHash = sha256File(blockedCustomWorkbookPath);
  const blockedCustomRequestPath = join(
    temporaryDirectory,
    'designer-blocked-custom.json',
  );
  writeFileSync(
    blockedCustomRequestPath,
    JSON.stringify({
      schemaVersion: 2,
      workbookPath: blockedCustomWorkbookPath,
      expectedWorkbookSha256: blockedCustomHash,
      operations: [
        {
          kind: 'createWorksheetActiveXControl',
          sheetName: 'Data',
          control: {
            type: 'customActiveX',
            progId: 'Forms.CommandButton.1',
            name: 'btnBlockedCustom',
            left: 20,
            top: 20,
            width: 120,
            height: 28,
          },
        },
      ],
    }),
    'utf8',
  );
  const blockedCustomResult = runDesigner(blockedCustomRequestPath);
  assert.notEqual(
    blockedCustomResult.status,
    0,
    'non-allowlisted custom ActiveX creation must fail',
  );
  assert.match(
    String(blockedCustomResult.stderr),
    /Custom ActiveX ProgID is not allowlisted/i,
  );
  assert.equal(
    sha256File(blockedCustomWorkbookPath),
    blockedCustomHash,
    'custom ActiveX allowlist refusal changed the workbook',
  );
  if (!activeXAvailable) {
    const blockedByExcelWorkbookPath = join(
      temporaryDirectory,
      'designer-activex-policy.xlsm',
    );
    copyFileSync(fixture, blockedByExcelWorkbookPath);
    const blockedByExcelHash = sha256File(blockedByExcelWorkbookPath);
    const blockedByExcelRequestPath = join(
      temporaryDirectory,
      'designer-activex-policy.json',
    );
    writeFileSync(
      blockedByExcelRequestPath,
      JSON.stringify({
        schemaVersion: 2,
        workbookPath: blockedByExcelWorkbookPath,
        expectedWorkbookSha256: blockedByExcelHash,
        operations: [
          {
            kind: 'createWorksheetActiveXControl',
            sheetName: 'Data',
            control: {
              type: 'commandButton',
              name: 'btnPolicyBlocked',
              left: 20,
              top: 20,
              width: 120,
              height: 28,
            },
          },
        ],
      }),
      'utf8',
    );
    const blockedByExcelResult = runDesigner(blockedByExcelRequestPath);
    assert.notEqual(
      blockedByExcelResult.status,
      0,
      'Office ActiveX policy probe unexpectedly succeeded',
    );
    assert.match(
      normalizeDiagnostic(blockedByExcelResult.stderr),
      /Excel refused ActiveX insertion/i,
    );
    assert.equal(
      sha256File(blockedByExcelWorkbookPath),
      blockedByExcelHash,
      'Office ActiveX refusal changed the workbook',
    );
  }
  assert.ok(
    designerEntries.some(
      ([name, digest]) =>
        name.startsWith('ouserform/') && /^[0-9a-f]{64}$/.test(digest),
    ),
  );

  const rollbackWorkbookPath = join(
    temporaryDirectory,
    'designer-rollback.xlsm',
  );
  copyFileSync(fixture, rollbackWorkbookPath);
  const rollbackHash = sha256File(rollbackWorkbookPath);
  const rollbackRequestPath = join(
    temporaryDirectory,
    'designer-rollback.json',
  );
  writeFileSync(
    rollbackRequestPath,
    JSON.stringify({
      schemaVersion: 2,
      workbookPath: rollbackWorkbookPath,
      expectedWorkbookSha256: '0'.repeat(64),
      operations: [
        {
          kind: 'createUserForm',
          name: 'frmMustNotExist',
        },
      ],
    }),
    'utf8',
  );
  const rollbackResult = runDesigner(rollbackRequestPath);
  assert.notEqual(rollbackResult.status, 0, 'stale hash request must fail');
  assert.equal(
    sha256File(rollbackWorkbookPath),
    rollbackHash,
    'failed transaction changed the workbook',
  );

  const missingMacroWorkbookPath = join(
    temporaryDirectory,
    'designer-missing-macro.xlsm',
  );
  copyFileSync(fixture, missingMacroWorkbookPath);
  const missingMacroHash = sha256File(missingMacroWorkbookPath);
  const missingMacroRequestPath = join(
    temporaryDirectory,
    'designer-missing-macro.json',
  );
  writeFileSync(
    missingMacroRequestPath,
    JSON.stringify({
      schemaVersion: 2,
      workbookPath: missingMacroWorkbookPath,
      expectedWorkbookSha256: missingMacroHash,
      operations: [
        {
          kind: 'createWorksheetButton',
          sheetName: 'Data',
          name: 'btnMissingMacro',
          caption: 'Must fail',
          macroName: 'mCode.DoesNotExist',
          left: 20,
          top: 20,
          width: 120,
          height: 28,
        },
      ],
    }),
    'utf8',
  );
  const missingMacroResult = runDesigner(missingMacroRequestPath);
  assert.notEqual(
    missingMacroResult.status,
    0,
    'missing macro assignment must fail',
  );
  assert.match(
    String(missingMacroResult.stderr),
    /Public macro\s+procedure\s+'mCode\.DoesNotExist' was not found/,
  );
  assert.equal(
    sha256File(missingMacroWorkbookPath),
    missingMacroHash,
    'missing macro validation changed the workbook',
  );

  assertNoNewExcelProcess(excelBefore);
  console.log(
    `VBA designer integration passed: UserForms, 12 controls, Form button assignment, custom ProgID allowlist, backup, native verification and rollback. Worksheet ActiveX: ${
      activeXAvailable
        ? 'creation and binding verified'
        : 'Office policy blocked insertion; safe refusal verified'
    }.`,
  );
} finally {
  let processError;
  try {
    assertNoNewExcelProcess(excelBefore);
  } catch (error) {
    processError = error;
  }
  rmSync(temporaryDirectory, { recursive: true, force: true });
  if (processError) {
    throw processError;
  }
}
