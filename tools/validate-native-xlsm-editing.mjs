import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (relativePath) => readFile(path.join(root, relativePath), 'utf8');

const [
  nativeTypes,
  workbookObjects,
  nativeDiff,
  bridge,
  commonHandler,
  officeProvider,
  excelView,
  excelWriter,
  spreadsheet,
  script,
  vbaCli,
] = await Promise.all([
  read('src/common/nativeExcelEdits.ts'),
  read('src/common/excelWorkbookObjects.ts'),
  read('src/react/view/excel/native_edit_diff.ts'),
  read('src/provider/nativeExcelBridge.ts'),
  read('src/provider/compress/commonHandler.ts'),
  read('src/provider/officeViewerProvider.ts'),
  read('src/react/view/excel/Excel.tsx'),
  read('src/react/view/excel/excel_writer.ts'),
  read('src/react/view/excel/x-spreadsheet/index.ts'),
  read('scripts/office-ai-apply-edits.ps1'),
  read('native/vba-writeback/cli.py'),
]);

assert.match(nativeTypes, /type NativeExcelEditOperation/);
assert.match(nativeTypes, /kind:\s*'columnWidth'/);
assert.match(nativeTypes, /kind:\s*'rowHeight'/);
assert.match(nativeTypes, /kind:\s*'addConditionalFormatting'/);
assert.match(nativeTypes, /kind:\s*'clearConditionalFormatting'/);
assert.match(nativeTypes, /NativeExcelConditionalFormattingRule/);
for (const kind of [
  'createTable', 'updateTable', 'deleteTable',
  'createChart', 'updateChart', 'deleteChart',
]) {
  assert.match(nativeTypes, new RegExp(`kind:\\s*'${kind}'`));
}
assert.match(workbookObjects, /EXCEL_CHART_TYPES/);
assert.match(workbookObjects, /buildExcelTableStyleCatalog/);
assert.match(workbookObjects, /SheetChartSeriesData/);
assert.match(nativeDiff, /buildConditionalFormattingOperations/);
assert.match(nativeDiff, /buildObjectOperations/);
assert.match(nativeDiff, /tables:\s*_tables/);
assert.match(nativeDiff, /charts:\s*_charts/);
assert.match(nativeDiff, /hasNativeChartParts:\s*_hasNativeChartParts/);
assert.match(nativeDiff, /unsupportedNativeChartCount:\s*_unsupportedNativeChartCount/);
assert.match(nativeDiff, /normalizeConditionalRule/);
assert.match(nativeDiff, /conditionalFormattings:\s*_conditionalFormattings/);
assert.match(nativeDiff, /conditional-formatting/);
assert.match(nativeDiff, /kind:\s*'columnWidth'/);
assert.match(nativeDiff, /kind:\s*'rowHeight'/);
assert.match(bridge, /MAX_NATIVE_EDIT_OPERATIONS\s*=\s*10_000/);
assert.match(
  bridge,
  /MAX_CONDITIONAL_FORMATTING_ADDS_PER_SHEET\s*=\s*64/,
);
assert.match(bridge, /MAX_NATIVE_EDIT_PAYLOAD_BYTES\s*=\s*4 \* 1024 \* 1024/);
assert.match(bridge, /\['\.xlsx', '\.xlsm'\]\.includes\(workbookExtension\)/);
assert.match(bridge, /normalizeTableData/);
assert.match(bridge, /normalizeChartData/);
assert.match(bridge, /VALID_CHART_TYPES/);
assert.match(bridge, /normalizeTableName/);
assert.match(bridge, /isValidExcelTableName/);
assert.match(bridge, /normalizeExcelTableName/);
assert.match(bridge, /cannot be an Excel formula/);
assert.match(bridge, /windowsHide:\s*true/);
assert.match(bridge, /fs\.rm\(payloadPath,\s*\{\s*force:\s*true\s*\}\)/);
assert.match(bridge, /parseOwnedExcelProcessId/);
assert.match(bridge, /terminateExactProcess/);
assert.match(bridge, /\$candidate\.Kill\(\)/);
assert.doesNotMatch(bridge, /Stop-Process -Id \$expectedPid/);
assert.match(bridge, /expectedWorkbookSha256/);
assert.match(bridge, /assertNoReparsePointChain/);
assert.match(bridge, /getFileSha256/);
assert.match(bridge, /parseNativeEditResult/);
assert.match(bridge, /must change a value or style/);
assert.match(bridge, /contains an invalid value/);
assert.match(bridge, /normalizeRangeRef/);
assert.match(bridge, /assertAllowedKeys/);
assert.match(bridge, /NativeExcelEditOperation/);
assert.match(
  commonHandler,
  /nativeLoadGeneration\s*=\s*supportsNativeWorkbookObjectEditing\(uri\)[\s\S]+?randomUUID\(\)/,
);
assert.doesNotMatch(commonHandler, /hasWorkbookObjectOperation/);
assert.match(
  commonHandler,
  /state\.readOnlyReason\s*!==\s*'native-excel-editing'[\s\S]+?extname\(uri\.fsPath\)\.toLowerCase\(\)\s*!==\s*'\.xlsx'/,
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
assert.match(excelView, /hasExistingNativeObjects/);
assert.match(excelView, /sheet\.tables\?\.length/);
assert.match(excelView, /sheet\.charts\?\.length/);
assert.match(excelView, /function isOoxmlSaveAsExt\(ext: string\)/);
assert.match(excelView, /isOoxmlSaveAsExt\(ext\)[\s\S]+?sheetsContainUnsafeNativeOoxmlObjects\(spreadSheet\.getData\(\)\)/);
const unsafeSaveAsHelper = excelView.match(
  /function sheetsContainUnsafeNativeOoxmlObjects\([\s\S]+?\{([\s\S]+?)\n\}/,
)?.[1] ?? '';
assert.match(unsafeSaveAsHelper, /sheet\.charts\?\.length/);
assert.match(unsafeSaveAsHelper, /sheet\.hasNativeChartParts === true/);
assert.match(unsafeSaveAsHelper, /sheet\.tables\?\.length/);
assert.doesNotMatch(excelView, /nativeObjectSaveAsBlocked/);
assert.match(excelView, /allowSaveAs:\s*!preserveSourceIntegrity/);
assert.match(excelView, /allowSaveAs=\{!blocksSaveAs\(readOnlyReason\)\}/);
assert.match(excelView, /tableaux et graphiques Excel natifs/);
assert.match(excelView, /Choisissez CSV, TSV, ODS ou XLS pour un export aplati/);
assert.match(spreadsheet, /hasNativeWorkbookObjects\(\): boolean/);
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
assert.match(script, /Apply-ColumnWidthOperation/);
assert.match(script, /Apply-RowHeightOperation/);
assert.match(script, /Apply-ConditionalFormattingOperation/);
assert.match(script, /Apply-ClearConditionalFormattingOperation/);
assert.match(script, /Apply-CreateTableOperation/);
assert.match(script, /Apply-UpdateTableOperation/);
assert.match(script, /Apply-DeleteTableOperation/);
assert.match(script, /Apply-CreateChartOperation/);
assert.match(script, /Apply-UpdateChartOperation/);
assert.match(script, /Apply-DeleteChartOperation/);
assert.match(script, /\.ListObjects\.Item|\$worksheet\.ListObjects/);
assert.match(script, /\$worksheet\.ChartObjects\(\)/);
assert.match(script, /\.SetSourceData\(/);
assert.match(script, /\.NewSeries\(\)/);
assert.match(script, /\.FullSeriesCollection\(\)/);
assert.match(script, /\.IsFiltered\s*=/);
assert.match(script, /\.Unlist\(\)/);
assert.doesNotMatch(script, /\$listObject\.Delete\(\)/);
assert.match(script, /Test-NativeExcelTableName/);
assert.match(script, /cannot be an Excel formula/);
assert.match(script, /Assert-NativeWorkbookObjectOperations/);
assert.match(script, /Get-NativeUntargetedWorkbookObjectSnapshot/);
assert.match(script, /\$workbooks\.Open\(\$workPath, 0, \$true\)/);
assert.match(script, /AllowWorkbookObjectChanges/);
assert.match(script, /SetLastPriority/);
assert.match(script, /AddColorScale\(3\)/);
assert.match(script, /AddDatabar\(\)/);
assert.match(script, /AddIconSetCondition\(\)/);
assert.match(script, /protected worksheet control/);
assert.doesNotMatch(script, /jj\/mm\/aaaa/);
assert.doesNotMatch(script, /\.SaveCopyAs\(/);
assert.doesNotMatch(script, /Copy-Item[\s\S]+?-Destination\s+\$workbookFullPath/);
assert.doesNotMatch(script, /\$workbooks\.Open\(\$workbookFullPath/);
assert.doesNotMatch(script, /\.Run\(|RunAutoMacros|Application\.Run/);

console.log('Native XLSX/XLSM editing safety validation passed.');
