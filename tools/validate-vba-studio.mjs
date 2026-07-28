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
const interactionGraph = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/vbaInteractionGraph.ts'),
  'utf8',
);
const manifest = JSON.parse(
  readFileSync(resolve(root, 'package.json'), 'utf8'),
);

const htmlMarker = 'return `<!doctype html>';
const htmlStart = panel.indexOf(htmlMarker);
const htmlEnd = panel.indexOf('</html>`;', htmlStart);
assert.ok(htmlStart >= 0 && htmlEnd > htmlStart, 'VBA Studio HTML template is missing');
const htmlTemplate = panel.slice(
  htmlStart + 'return `'.length,
  htmlEnd + '</html>'.length,
);
const renderedHtml = Function(
  'nonce',
  `return \`${htmlTemplate}\`;`,
)('validation-nonce');
const webviewScript = /<script nonce="validation-nonce">([\s\S]*?)<\/script>/.exec(
  renderedHtml,
)?.[1];
assert.ok(webviewScript, 'VBA Studio webview script is missing');
assert.doesNotThrow(
  () => Function(webviewScript),
  'VBA Studio generated webview JavaScript must parse',
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
assert.match(panel, /controls-view/);
assert.match(panel, /designer-view/);
assert.match(panel, /renderDesigner/);
assert.match(panel, /addVisualControl/);
assert.match(panel, /updateVisualControl/);
assert.match(panel, /setUserFormEvent/);
assert.match(panel, /replaceExisting/);
assert.match(panel, /assignFormButton/);
assert.match(panel, /bindActiveX/);
assert.match(panel, /createWorksheetActiveXControl/);
assert.match(
  panel,
  /Simulation uniquement[\s\S]+?Aucune macro exécutée/,
  'Controls view must label simulated clicks and avoid VBA execution',
);
assert.match(interactionGraph, /extractPublicZeroArgumentMacros/);
assert.match(interactionGraph, /userFormsOpened/);
assert.match(interactionGraph, /activeXMacroTarget/);
assert.doesNotMatch(interactionGraph, /\.Run\s*\(/i);
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
assert.match(workbookService, /openVbaDeveloperWindow/);
assert.match(
  workbookService,
  /vscode\.openFolder[\s\S]+?forceNewWindow:\s*true/,
  'Developer mode must open a separate VS Code window',
);
assert.match(workbookService, /DEVELOPER_MARKER_NAME/);
assert.match(workbookService, /pathIsInside\(workspaceDirectory,\s*exportsRoot\)/);
assert.match(workbookService, /workbookSha256/);
assert.ok(
  manifest.activationEvents.includes(
    'workspaceContains:.excel-ai-vba-studio-project.json',
  ),
  'Developer workspace marker activation is missing',
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
assert.match(exporter, /worksheetButtons/);
assert.match(exporter, /worksheetActiveXControls/);
assert.match(exporter, /userForms\s*=\s*@\(\)/);
assert.match(exporter, /Get-VbaControlTypeLabel/);
assert.match(exporter, /workbook\.vba\.userForms\.controls/);
assert.match(
  exporter,
  /This is metadata only; no OnAction or event is invoked/,
);
assert.doesNotMatch(exporter, /\.Run\s*\(/i, 'exporter must never run a macro');
assert.equal(
  manifest.scripts?.['test:vba-inventory'],
  'node test/vba-userform-inventory.mjs',
);
assert.match(manifest.scripts?.validate ?? '', /test:vba-inventory/);

console.log('VBA Studio validation passed: separate Developer window, visual UserForm designer, event binding, static control graph, real sources, secure panel.');
