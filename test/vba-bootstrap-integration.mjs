import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { spawnSync } from 'node:child_process';
import { dirname, join, resolve } from 'node:path';
import { tmpdir } from 'node:os';

if (process.platform !== 'win32') {
  console.log('XLSX VBA bootstrap integration skipped: Windows and Excel are required.');
  process.exit(0);
}

const root = resolve(import.meta.dirname, '..');
const script = resolve(root, 'scripts/prepare-macro-workbook.ps1');
const helper = resolve(root, 'bin/win32-x64/excel-ai-vba-writeback.exe');
const fixture = resolve(root, 'test/fixtures/Excel-AI-VBA-Studio-Demo-base.xlsx');
const temp = mkdtempSync(join(tmpdir(), 'excel-ai-vba-bootstrap-test-'));
let requestIndex = 0;

const sha256File = (filePath) =>
  crypto.createHash('sha256').update(readFileSync(filePath)).digest('hex');
const base64 = (value) => Buffer.from(value, 'utf8').toString('base64');

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
    { encoding: 'utf8', shell: false, windowsHide: true },
  );
  return result.status === 0;
}

function runBootstrap(workbookPath, componentFile, source, expectedSuccess = true) {
  const result = spawnSync(
    'powershell.exe',
    [
      '-NoLogo',
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      script,
      '-WorkbookPathBase64',
      base64(workbookPath),
      '-ComponentFileBase64',
      base64(componentFile),
      '-SourceBase64',
      base64(source),
    ],
    {
      cwd: dirname(workbookPath),
      encoding: 'utf8',
      shell: false,
      windowsHide: true,
      timeout: 90_000,
    },
  );
  assert.equal(
    result.status === 0,
    expectedSuccess,
    `${result.stdout}\n${result.stderr}`,
  );
  if (!expectedSuccess) {
    return result;
  }
  const jsonLine = String(result.stdout)
    .trim()
    .split(/\r?\n/)
    .filter(Boolean)
    .at(-1);
  const output = JSON.parse(jsonLine);
  assert.equal(output.ok, true);
  assert.match(String(result.stdout), /OWNED_EXCEL_PID\|\d+/);
  return output;
}

function inspect(workbookPath) {
  const requestPath = join(temp, `inspect-${requestIndex++}.json`);
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
    cwd: temp,
    encoding: 'utf8',
    shell: false,
    windowsHide: true,
  });
  assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
  return JSON.parse(
    String(result.stdout)
      .trim()
      .split(/\r?\n/)
      .filter(Boolean)
      .at(-1),
  );
}

assert.ok(existsSync(script), `Bootstrap script missing: ${script}`);
assert.ok(existsSync(helper), `Bundled helper missing: ${helper}`);
assert.ok(existsSync(fixture), `Fixture missing: ${fixture}`);

const excelBefore = excelProcessIds();
if (!excelAutomationReady()) {
  console.log(
    'XLSX VBA bootstrap integration skipped: Excel COM or user-enabled AccessVBOM is unavailable.',
  );
  rmSync(temp, { recursive: true, force: true });
  process.exit(0);
}

try {
  const moduleWorkbook = join(temp, 'module.xlsx');
  copyFileSync(fixture, moduleWorkbook);
  const moduleSourceHash = sha256File(moduleWorkbook);
  const moduleSource = [
    'Attribute VB_Name = "CalculateBudget"',
    'Option Explicit',
    '',
    'Public Function CalculateBudget(ByVal planned As Currency, ByVal actual As Currency) As Currency',
    'Attribute CalculateBudget.VB_Description = "Budget variance"',
    '    CalculateBudget = planned - actual',
    'End Function',
    '',
  ].join('\r\n');
  const moduleResult = runBootstrap(
    moduleWorkbook,
    'CalculateBudget.bas',
    moduleSource,
  );
  const moduleTarget = join(temp, 'module.xlsm');
  assert.equal(moduleResult.sourceWorkbookPath, moduleWorkbook);
  assert.equal(moduleResult.targetWorkbookPath, moduleTarget);
  assert.equal(moduleResult.convertedToXlsm, true);
  assert.deepEqual(moduleResult.modifiedModules, ['CalculateBudget']);
  assert.equal(sha256File(moduleWorkbook), moduleSourceHash);
  assert.equal(sha256File(moduleTarget), moduleResult.workbookSha256);
  const moduleInspection = inspect(moduleTarget);
  const calculateBudget = moduleInspection.modules.find(
    (module) => module.name === 'CalculateBudget',
  );
  assert.equal(calculateBudget?.componentKind, 'module');
  assert.match(calculateBudget?.source ?? '', /Function CalculateBudget/);

  const classWorkbook = join(temp, 'class.xlsx');
  copyFileSync(fixture, classWorkbook);
  const classSourceHash = sha256File(classWorkbook);
  const classSource = [
    'VERSION 1.0 CLASS',
    'BEGIN',
    '  MultiUse = -1',
    'END',
    'Attribute VB_Name = "BudgetItem"',
    'Attribute VB_GlobalNameSpace = False',
    'Attribute VB_Creatable = False',
    'Attribute VB_PredeclaredId = False',
    'Attribute VB_Exposed = False',
    'Option Explicit',
    '',
    'Private mAmount As Currency',
    '',
    'Public Property Get Amount() As Currency',
    '    Amount = mAmount',
    'End Property',
    '',
    'Public Property Let Amount(ByVal value As Currency)',
    '    mAmount = value',
    'End Property',
    '',
  ].join('\r\n');
  const classResult = runBootstrap(
    classWorkbook,
    'BudgetItem.cls',
    classSource,
  );
  const classTarget = join(temp, 'class.xlsm');
  assert.equal(classResult.targetWorkbookPath, classTarget);
  assert.equal(sha256File(classWorkbook), classSourceHash);
  const classInspection = inspect(classTarget);
  const budgetItem = classInspection.modules.find(
    (module) => module.name === 'BudgetItem',
  );
  assert.equal(budgetItem?.componentKind, 'class');
  assert.match(budgetItem?.source ?? '', /Property Get Amount/);

  const backupDirectory = join(temp, '.excel-ai-vba-backups');
  mkdirSync(backupDirectory);
  const backupWorkbook = join(backupDirectory, 'forbidden.xlsx');
  copyFileSync(fixture, backupWorkbook);
  const rejected = runBootstrap(
    backupWorkbook,
    'Forbidden.bas',
    'Attribute VB_Name = "Forbidden"\r\nOption Explicit\r\n',
    false,
  );
  assert.match(
    `${rejected.stdout}\n${rejected.stderr}`,
    /\.excel-ai-vba-backups/i,
  );
  assert.equal(existsSync(join(backupDirectory, 'forbidden.xlsm')), false);

  const collisionWorkbook = join(temp, 'collision.xlsx');
  const collisionTarget = join(temp, 'collision.xlsm');
  copyFileSync(fixture, collisionWorkbook);
  copyFileSync(fixture, collisionTarget);
  const collisionHash = sha256File(collisionTarget);
  const collision = runBootstrap(
    collisionWorkbook,
    'Collision.bas',
    'Attribute VB_Name = "Collision"\r\nOption Explicit\r\n',
    false,
  );
  assert.match(`${collision.stdout}\n${collision.stderr}`, /existe deja/i);
  assert.equal(sha256File(collisionTarget), collisionHash);

  const excelAfter = excelProcessIds();
  const leaked = [...excelAfter].filter((processId) => !excelBefore.has(processId));
  assert.deepEqual(leaked, [], `Excel processes leaked: ${leaked.join(', ')}`);

  console.log(
    'XLSX VBA bootstrap integration passed: module/class persisted, source preserved, backups rejected, no Excel leak.',
  );
} finally {
  rmSync(temp, { recursive: true, force: true });
}
