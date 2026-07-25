import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const excel = readFileSync(
  resolve(root, 'src/react/view/excel/Excel.tsx'),
  'utf8',
);
const studio = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/vbaStudioPanel.ts'),
  'utf8',
);

assert.match(
  excel,
  /observeVscodeThemeChange\([\s\S]+?setVscodeDark\(isVscodeEditorDark\(\)\)/,
  'the spreadsheet must observe VS Code theme changes',
);
assert.match(
  excel,
  /document\.body\.classList\.add\('office-adaptive'\)/,
  'the spreadsheet must always use VS Code adaptive colors',
);
assert.match(
  excel,
  /style\.colorScheme = themedDark \? 'dark' : 'light'/,
  'native spreadsheet controls must follow the active VS Code color scheme',
);
assert.doesNotMatch(
  excel,
  /loadExcelColorMode|saveExcelColorMode|className="dark-mode-toggle"/,
  'a stale manual light override must not replace the VS Code theme',
);

for (const themeClass of [
  'vscode-light',
  'vscode-dark',
  'vscode-high-contrast',
  'vscode-high-contrast-light',
]) {
  assert.ok(
    studio.includes(themeClass),
    `VBA Studio is missing ${themeClass} theme support`,
  );
}
for (const token of [
  '--vscode-editor-background',
  '--vscode-editor-foreground',
  '--vscode-list-activeSelectionBackground',
  '--vscode-focusBorder',
]) {
  assert.ok(studio.includes(token), `VBA Studio is missing ${token}`);
}

console.log('Theme validation passed: spreadsheet and VBA Studio follow VS Code automatically.');
