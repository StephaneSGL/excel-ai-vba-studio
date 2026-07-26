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
assert.match(bridge, /must change a value or style/);
assert.match(bridge, /contains an invalid value/);

assert.match(script, /\$excel\.Visible\s*=\s*\$false/);
assert.match(script, /\$excel\.AutomationSecurity\s*=\s*3/);
assert.match(script, /\$excel\.EnableEvents\s*=\s*\$false/);
assert.match(script, /\$excel\.DisplayAlerts\s*=\s*\$false/);
assert.match(script, /\$excel\.AskToUpdateLinks\s*=\s*\$false/);
assert.match(script, /\$workbook\.SaveCopyAs\(\$backupPath\)/);
assert.match(script, /OWNED_EXCEL_PID\|/);
assert.match(script, /Network and UNC workbook paths are not supported/);
assert.match(script, /NumberFormatLocal/);
assert.match(script, /International\(32\)/);
assert.match(script, /ColorIndex\s*=\s*-4105/);
assert.doesNotMatch(script, /jj\/mm\/aaaa/);
assert.match(script, /Copy-Item[\s\S]+?\$workbookFullPath -Force/);
assert.doesNotMatch(script, /\.Run\(|RunAutoMacros|Application\.Run/);

console.log('Native XLSM editing safety validation passed.');
