import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const scripts = ['office-ai-export.ps1', 'prepare-macro-workbook.ps1', 'office-ai-apply-edits.ps1', 'apply-vba-designer.ps1'];
let openCount = 0;
for (const name of scripts) {
  const source = readFileSync(join(root, 'scripts', name), 'utf8');
  assert.match(source, /\. \(Join-Path \$PSScriptRoot 'excel-automation-safety\.ps1'\)/);
  assert.match(source, /Get-ExcelAutomationProcessIds[\s\S]+?\$excel = New-Object -ComObject Excel\.Application[\s\S]+?Get-ExcelAutomationProcessIdentity[\s\S]+?\$excel\.AutomationSecurity\s*=\s*3/);
  assert.match(source, /OWNED_EXCEL_PID\|'\s*\+\s*\$\w+\.ProcessId\s*\+\s*'\|'\s*\+\s*\$\w+\.StartTimeUtcTicks[\s\S]+?\[Console\]::Out\.Flush\(\)/);
  assert.match(source, /Set-ExcelAutomationSafety \$\w+ -BeforeGuard\s*\r?\n\s*(?:\$workbooks = \$excel\.Workbooks\s*\r?\n\s*)?\$\w+ = \$workbooks\.Add\(-4167\)/);
  const lines = source.split(/\r?\n/);
  for (let index = 0; index < lines.length; index++) {
    if (!/\$workbooks\d*\.Open\(/i.test(lines[index])) continue;
    openCount++;
    assert.match(lines[index - 1], /Set-ExcelAutomationSafety \$\w+\s*$/, `${name}:${index + 1} must set/verify policy before Open`);
    assert.match(lines[index + 1], /Assert-ExcelAutomationSafety \$\w+\s*$/, `${name}:${index + 1} must verify policy after Open`);
    assert.match(lines[index], /\.Open\(\$\w+, 0, \$(?:true|false)\)/i);
  }
}
assert.equal(openCount, 6, 'All six workbook opens must stay behind the shared policy.');

const powershell = process.platform === 'win32'
  ? join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
  : 'pwsh';
const result = spawnSync(powershell, ['-NoLogo', '-NoProfile', '-NonInteractive', '-File', join(root, 'test/excel-automation-safety.ps1'), '-RepositoryRoot', root], {
  encoding: 'utf8', timeout: 60_000, windowsHide: true,
});
if (result.error?.code === 'ENOENT' && process.platform !== 'win32') {
  console.log('Excel automation wiring verified; synthetic PowerShell behavior tests skipped (pwsh unavailable).');
} else {
  assert.ifError(result.error);
  assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
  assert.match(result.stdout, /Synthetic Excel automation safety tests passed/);
  console.log('Excel automation safety passed: all six call sites, read-back failures, policy drift, ownership, cleanup, and PID reuse.');
}
