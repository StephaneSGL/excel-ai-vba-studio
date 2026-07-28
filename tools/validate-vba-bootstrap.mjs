import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const script = readFileSync(
  resolve(root, 'scripts/prepare-macro-workbook.ps1'),
  'utf8',
);
const security = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/security.ts'),
  'utf8',
);
const service = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/workbookService.ts'),
  'utf8',
);
const languageTool = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/languageModelTool.ts'),
  'utf8',
);
const manifest = readFileSync(resolve(root, 'package.json'), 'utf8');

for (const [pattern, message] of [
  [/AutomationSecurity\s*=\s*3/, 'macro automation disablement is missing'],
  [/OWNED_EXCEL_PID\|/, 'owned Excel process reporting is missing'],
  [/GetWindowThreadProcessId/, 'exact Excel PID ownership is missing'],
  [/\.excel-ai-vba-backups/, 'managed-backup refusal is missing'],
  [/Assert-NoReparsePointChain/, 'reparse-point checks are missing'],
  [/\.SaveAs\(\$stagingPath,\s*52\)/, 'XLSM SaveAs conversion is missing'],
  [/\.VBProject/, 'VBProject access is missing'],
  [/\.VBComponents/, 'VBComponent insertion is missing'],
  [/xl\/vbaProject\.bin/, 'persisted VBA package verification is missing'],
  [/sourceSha256Before/, 'source-preservation hash gate is missing'],
  [/function Get-Sha256[\s\S]+?SHA256\]::Create/, 'self-contained SHA-256 implementation is missing'],
  [/\[IO\.File\]::Move\(\$stagingPath,\s*\$targetPath\)/, 'non-overwriting atomic target commit is missing'],
  [/FinalReleaseComObject/, 'deterministic COM release is missing'],
]) {
  assert.match(script, pattern, message);
}

assert.doesNotMatch(
  script,
  /Set-ItemProperty|New-ItemProperty|reg\.exe|ExecuteExcel4Macro|\.Run\(/i,
  'bootstrap must not edit registry settings or execute macros',
);
assert.doesNotMatch(
  script,
  /Get-FileHash/,
  'bootstrap must not depend on PowerShell module autoloading for SHA-256',
);
assert.match(
  security,
  /assertNotManagedBackupPath[\s\S]+?\.excel-ai-vba-backups/,
  'extension-side managed-backup refusal is missing',
);
assert.match(
  service,
  /writeVbaFromTool[\s\S]+?prepare-macro-workbook\.ps1/,
  'XLSX bootstrap is not connected to the workbook service',
);
assert.match(
  service,
  /nouveau UserForm est refusée[\s\S]+?faux fichier \.frm/,
  'new UserForm fail-closed policy is missing',
);
assert.match(
  service,
  /targetWorkbookPath[\s\S]+?\.excel-ai-vba-backups/,
  'generated Copilot instructions do not enforce target-path continuity',
);
assert.match(
  languageTool,
  /EXCEL_AI_VBA_WRITE_TOOL[\s\S]+?writeVbaFromTool/,
  'Copilot write tool is not connected to the safe workbook service',
);
assert.match(
  languageTool,
  /targetWorkbookPath:\s*writeResult\.targetWorkbookPath/,
  'Copilot result does not expose the actual target workbook',
);
assert.match(
  manifest,
  /\.excel-ai-vba-backups/,
  'language-model tool contract does not reject managed backups',
);
assert.match(
  manifest,
  /\.xlsx[\s\S]+?targetWorkbookPath/,
  'language-model tool contract does not describe safe XLSX conversion',
);

console.log(
  'VBA bootstrap validation passed: safe XLSX conversion, explicit target, fail-closed UserForms.',
);
