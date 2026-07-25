import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const workbookService = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/workbookService.ts'),
  'utf8',
);
const panel = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/vbaStudioPanel.ts'),
  'utf8',
);
const exporter = readFileSync(
  resolve(root, 'scripts/office-ai-export.ps1'),
  'utf8',
);

assert.doesNotMatch(
  workbookService,
  /removeOwnedDirectory\(\s*paths\.vbaDirectory/,
  'refresh must not delete a VBA directory that VS Code/Copilot is indexing',
);
assert.match(
  workbookService,
  /vbaStudioPanel\.prepare\(result\)[\s\S]+?vbaStudioPanel\.open\(result\)/,
  'opening VBA Studio must prepare real sources before showing the panel',
);
assert.match(
  panel,
  /createWebviewPanel\(\s*'excelAiVbaStudio\.vbaStudio'/,
  'the integrated VBA Studio webview is missing',
);
assert.match(panel, /workspace\.applyEdit/, 'VBA Studio saves must update real VS Code documents');
assert.match(panel, /workbench\.action\.chat\.open/, 'VBA Studio must expose a Copilot action');
assert.match(panel, /copilot-instructions\.md|#excelVbaWorkbook/, 'Copilot context is missing');
for (const extension of ['.bas', '.cls', '.frm']) {
  assert.ok(panel.includes(extension), `${extension} source support is missing`);
}
assert.match(panel, /Content-Security-Policy/, 'the VBA Studio webview needs a CSP');
assert.match(panel, /MAX_SOURCE_CHARACTERS/, 'VBA source writes must remain bounded');
assert.match(
  exporter,
  /previous exporter manifest[\s\S]+?managedNames/,
  'refresh must selectively replace exported files and preserve workspace-authored modules',
);
assert.doesNotMatch(
  exporter,
  /foreach \(\$managedPattern in @\('\*\.bas'/,
  'refresh must not glob-delete every Copilot-editable VBA source file',
);

console.log('VBA Studio validation passed: stable workspace, real sources, Copilot bridge, secure panel.');
