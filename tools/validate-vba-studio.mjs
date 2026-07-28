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
const writeback = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/vbaWritebackService.ts'),
  'utf8',
);
const userFormPreview = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/userFormPreview.ts'),
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
assert.ok(
  panel.indexOf('webview.onDidReceiveMessage') < panel.indexOf('webview.html ='),
  'the webview message receiver must be registered before HTML can post ready',
);
assert.match(panel, /workspace\.applyEdit/, 'VBA Studio saves must update real VS Code documents');
assert.match(
  panel,
  /onDidSaveTextDocument[\s\S]+?writebackService\.applySource/,
  'saved Copilot edits must trigger automatic workbook reinjection',
);
assert.match(panel, /workbench\.action\.chat\.open/, 'VBA Studio must expose a Copilot action');
assert.match(panel, /copilot-instructions\.md|#excelVbaWorkbook/, 'Copilot context is missing');
for (const extension of ['.bas', '.cls', '.frm']) {
  assert.ok(panel.includes(extension), `${extension} source support is missing`);
}
assert.match(panel, /Content-Security-Policy/, 'the VBA Studio webview needs a CSP');
assert.match(panel, /MAX_SOURCE_CHARACTERS/, 'VBA source writes must remain bounded');
assert.match(writeback, /expectedWorkbookSha256/, 'write-back must reject stale workbook baselines');
assert.match(writeback, /excel-ai-vba-writeback\.exe/, 'the bundled binary write-back helper is missing');
assert.match(writeback, /designerSha256/, 'UserForm designer changes must be detected');
assert.match(
  userFormPreview,
  /Content-Security-Policy[\s\S]+?script-src 'nonce-\$\{nonce\}'/,
  'interactive UserForm preview must keep a nonce-scoped CSP',
);
assert.match(
  userFormPreview,
  /enableScripts:\s*true[\s\S]+?localResourceRoots:\s*\[\]/,
  'interactive UserForm preview must deny local-resource access',
);
assert.match(
  userFormPreview,
  /Aucun code VBA exécuté/,
  'interactive preview must state that it never executes VBA',
);
for (const controlType of [
  'type="text"',
  'type="checkbox"',
  'type="radio"',
  'type="number"',
  'type="range"',
  '<select',
  '<fieldset',
]) {
  assert.ok(
    userFormPreview.includes(controlType),
    `interactive UserForm preview is missing ${controlType}`,
  );
}
assert.match(
  userFormPreview,
  /openExcel[\s\S]+?excelAiVbaStudio\.openExcel[\s\S]+?openVbe[\s\S]+?excelAiVbaStudio\.openVbe/,
  'UserForm preview must hand real event testing to native Excel/VBE',
);
assert.match(
  workbookService,
  /showUserFormPreview\([\s\S]+?context\.workbookUri/,
  'UserForm preview must receive the exact source workbook',
);
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
