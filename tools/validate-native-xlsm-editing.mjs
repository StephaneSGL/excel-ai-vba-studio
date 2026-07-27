import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (relativePath) => readFile(path.join(root, relativePath), 'utf8');

const [
  bridge,
  commonHandler,
  officeProvider,
  excelView,
  excelWriter,
  spreadsheet,
  script,
  vbaCli,
] = await Promise.all([
  read('src/provider/nativeExcelBridge.ts'),
  read('src/provider/compress/commonHandler.ts'),
  read('src/provider/officeViewerProvider.ts'),
  read('src/react/view/excel/Excel.tsx'),
  read('src/react/view/excel/excel_writer.ts'),
  read('src/react/view/excel/x-spreadsheet/index.ts'),
  read('scripts/office-ai-apply-edits.ps1'),
  read('native/vba-writeback/cli.py'),
]);

assert.match(bridge, /MAX_NATIVE_EDIT_OPERATIONS\s*=\s*10_000/);
assert.match(bridge, /MAX_NATIVE_EDIT_PAYLOAD_BYTES\s*=\s*4 \* 1024 \* 1024/);
assert.match(bridge, /extname\(workbookPath\).*\.xlsm/s);
assert.match(bridge, /windowsHide:\s*true/);
assert.match(bridge, /fs\.rm\(payloadPath,\s*\{\s*force:\s*true\s*\}\)/);
assert.match(bridge, /parseOwnedExcelProcessId/);
assert.match(bridge, /terminateExactProcess/);
assert.match(bridge, /expectedWorkbookSha256/);
assert.match(bridge, /assertNoReparsePointChain/);
assert.match(bridge, /getFileSha256/);
assert.match(bridge, /parseNativeEditResult/);
assert.match(bridge, /must change a value or style/);
assert.match(bridge, /contains an invalid value/);
assert.match(
  commonHandler,
  /nativeLoadGeneration\s*=\s*supportsNativeMacroEditing\(uri\)[\s\S]+?randomUUID\(\)/,
);
assert.match(
  commonHandler,
  /payload\?\.nativeLoadGeneration\s*!==\s*nativeLoadGeneration/,
);
assert.match(commonHandler, /sendQueue\.then/);
assert.match(commonHandler, /hasUnsavedChanges/);
assert.match(commonHandler, /Recharger le fichier/);
assert.match(commonHandler, /requestBackup/);
assert.match(commonHandler, /requestSaveAs/);
assert.match(commonHandler, /requestHostSave/);
assert.match(commonHandler, /instanceof vscode\.TabInputCustom/);
assert.match(commonHandler, /handler\.panel\.active/);
assert.match(commonHandler, /payload\.expectedWorkbookSha256/);
assert.match(officeProvider, /CustomEditorProvider<OfficeCustomDocument>/);
assert.match(officeProvider, /onDidChangeCustomDocument/);
assert.match(officeProvider, /saveCustomDocument/);
assert.match(officeProvider, /revertCustomDocument/);
assert.match(officeProvider, /backupCustomDocument/);
assert.match(officeProvider, /openContext\.backupId/);
assert.match(officeProvider, /currentSha256 !== parsed\.sourceSha256/);
assert.match(officeProvider, /document\.hasSession\(\)/);
assert.doesNotMatch(
  officeProvider,
  /executeCommand\('workbench\.action\.files\.save'\)/,
);
assert.match(excelView, /crypto\.subtle\.digest\('SHA-256',\s*buffer\)/);
assert.match(excelView, /nativeSnapshotRef/);
assert.match(excelView, /loadedWorkbookSha256Ref/);
assert.match(excelView, /sourceSha256/);
assert.match(excelView, /payload\.backupSourceSha256 !== workbookSha256/);
assert.match(excelView, /openGenerationRef/);
assert.match(excelView, /handler\.emit\('clean'\)/);
assert.match(excelView, /setEditorBusyState\(true\)/);
assert.match(excelView, /readOnly=\{readOnly \|\| editorBusy\}/);
assert.match(excelView, /setEditorBusyState\(false\)/);
assert.match(spreadsheet, /setMode\(mode:\s*'edit' \| 'read'\)/);
assert.match(excelWriter, /Unsupported spreadsheet save format/);

assert.match(script, /\$excel\.Visible\s*=\s*\$false/);
assert.match(script, /\$excel\.AutomationSecurity\s*=\s*3/);
assert.match(script, /\$excel\.EnableEvents\s*=\s*\$false/);
assert.match(script, /\$excel\.DisplayAlerts\s*=\s*\$false/);
assert.match(script, /\$excel\.AskToUpdateLinks\s*=\s*\$false/);
assert.match(script, /\$excel\.Calculation\s*=\s*-4135/);
assert.match(script, /\$excel\.CalculateBeforeSave\s*=\s*\$false/);
assert.match(script, /\[IO\.File\]::Replace\(/);
assert.match(script, /\.excel-ai-vba-backups/);
assert.match(script, /expectedWorkbookSha256/);
assert.match(script, /Assert-NoReparsePointChain/);
assert.match(script, /Test-MacroWorkbookPackage/);
assert.match(script, /sourceHasVbaProject/);
assert.match(script, /Get-StreamSha256Hex/);
assert.match(
  script,
  /\[IO\.FileShare\]::Read\s+-bor\s+\[IO\.FileShare\]::Delete/,
);
assert.match(script, /Get-PackagePreservationState/);
assert.match(script, /Compare-PackagePreservationState/);
assert.match(script, /Get-VbaProjectFingerprint/);
assert.match(vbaCli, /def project_stream_fingerprint/);
assert.match(vbaCli, /projectFingerprintSha256/);
assert.match(script, /Automatic rollback/);
assert.match(script, /Restore-DisplacedWorkbook/);
assert.match(script, /Restore-MissingWorkbook/);
assert.match(script, /\$workbookSha256\s+-cne\s+\$validatedWorkSha256/);
assert.match(script, /Displaced version retained at/);
assert.match(script, /OWNED_EXCEL_PID\|/);
assert.match(script, /Network and UNC paths are not supported/);
assert.match(script, /NumberFormatLocal/);
assert.match(script, /International\(32\)/);
assert.match(script, /ColorIndex\s*=\s*-4105/);
assert.doesNotMatch(script, /jj\/mm\/aaaa/);
assert.doesNotMatch(script, /\.SaveCopyAs\(/);
assert.doesNotMatch(script, /Copy-Item[\s\S]+?-Destination\s+\$workbookFullPath/);
assert.doesNotMatch(script, /\$workbooks\.Open\(\$workbookFullPath/);
assert.doesNotMatch(script, /\.Run\(|RunAutoMacros|Application\.Run/);

console.log('Native XLSM editing safety validation passed.');
