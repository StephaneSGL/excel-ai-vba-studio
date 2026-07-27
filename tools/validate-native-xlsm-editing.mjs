import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (relativePath) => readFile(path.join(root, relativePath), 'utf8');

const [bridge, script] = await Promise.all([
  read('src/provider/nativeExcelBridge.ts'),
  read('scripts/office-ai-apply-edits.ps1'),
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
