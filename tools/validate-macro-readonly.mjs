import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (relativePath) =>
  readFile(path.join(root, relativePath), 'utf8');

const [
  officeContent,
  commonHandler,
  excelUi,
  spreadsheetOptions,
  toolbar,
  sheet,
] = await Promise.all([
  read('src/provider/handlers/officeContent.ts'),
  read('src/provider/compress/commonHandler.ts'),
  read('src/react/view/excel/Excel.tsx'),
  read('src/react/view/excel/x-spreadsheet/index.ts'),
  read('src/react/view/excel/x-spreadsheet/component/toolbar/index.js'),
  read('src/react/view/excel/x-spreadsheet/component/sheet.js'),
]);

const protectedSet = officeContent.match(
  /MACRO_OR_LEGACY_EXTENSIONS\s*=\s*new Set\(\[([^\]]+)\]\)/
);
assert.ok(protectedSet, 'macro/legacy extension allowlist is missing');
const protectedExtensions = [
  ...protectedSet[1].matchAll(/['"]([^'"]+)['"]/g),
].map((match) => match[1]).sort();
assert.deepEqual(
  protectedExtensions,
  ['.xls', '.xlsm'],
  'only .xlsm and .xls should require a macro-safe write path'
);

for (const editable of ['.xlsx', '.csv', '.tsv']) {
  assert.ok(
    !protectedExtensions.includes(editable),
    `${editable} must remain editable`
  );
}

assert.match(
  officeContent,
  /readOnlyReason:\s*'macro-preservation'/,
  'legacy .xls payloads must explain macro-preservation read-only mode'
);
assert.match(
  officeContent,
  /supportsNativeMacroEditing[\s\S]+?=== '\.xlsm'/,
  '.xlsm must use the targeted native editing mode'
);
assert.match(
  officeContent,
  /readOnly:\s*false,\s*readOnlyReason:\s*'native-excel-editing'/,
  'writable .xlsm files must open in editable native mode'
);
assert.match(
  officeContent,
  /\.\.\.readOnlyState/,
  'read-only state must be propagated to the webview payload'
);

const saveHandler = commonHandler.match(
  /\.on\('save',[\s\S]+?\.on\('saveAs'/
)?.[0] ?? '';
const saveAsHandler = commonHandler.match(
  /\.on\('saveAs',[\s\S]+?\.on\('openVbaDeveloper'/
)?.[0] ?? '';
assert.match(
  saveHandler,
  /readOnlyReason === 'macro-preservation'/,
  'direct save must reject macro-preservation documents in the extension host'
);
assert.match(
  saveHandler,
  /readOnlyReason === 'native-excel-editing'/,
  'direct binary save must reject native-editing XLSM documents'
);
assert.match(
  saveAsHandler,
  /readOnlyReason === 'macro-preservation'/,
  'Save As must reject macro-preservation documents in the extension host'
);
assert.match(
  saveAsHandler,
  /readOnlyReason === 'native-excel-editing'/,
  'Save As must reject native-editing XLSM documents'
);
assert.match(
  commonHandler,
  /\.on\('saveNative'[\s\S]+?applyNativeExcelEdits/,
  'XLSM save must use targeted native Excel operations'
);
assert.match(
  commonHandler,
  /excelAiVbaStudio\.openVbaDeveloper/,
  'the internal VBA Studio action must remain available'
);
assert.match(
  commonHandler,
  /excelAiVbaStudio\.openVbaExplorer/,
  'the VBA source explorer action must remain available'
);

assert.match(
  excelUi,
  /allowSaveAs:\s*!preserveMacros/,
  'macro-preserving modes must suppress Save As in the grid'
);
assert.match(
  excelUi,
  /mode:\s*fileReadOnly \? 'read' : 'edit'/,
  'webview read-only mode must follow the extension-host payload'
);
assert.match(
  excelUi,
  /viewer\.macroWriteBlocked/,
  'the webview must show an explicit blocked-write message'
);
assert.match(
  excelUi,
  /buildNativeExcelEditPlan/,
  'the webview must diff XLSM edits before saving'
);
assert.match(
  excelUi,
  /handler\.emit\('saveNative', plan\.operations\)/,
  'supported XLSM edits must be sent to the native save handler'
);
assert.match(
  excelUi,
  /openVbaDeveloper/,
  'the protected-format banner must retain an internal VBA Studio action'
);

assert.match(
  spreadsheetOptions,
  /allowSaveAs\?: boolean/,
  'spreadsheet options must carry the Save As policy'
);
assert.match(
  toolbar,
  /allowSaveAs === false[\s\S]+?saveAsEl\.el\.hide/,
  'Save As must be hidden when the policy forbids it'
);
assert.match(
  sheet,
  /type === 'save-as'[\s\S]+?allowSaveAs === false\) return/,
  'Save As events must be suppressed when the policy forbids them'
);

console.log(
  'Macro safety validation passed: .xlsm uses native editing, .xls remains protected, and .xlsx/.csv/.tsv remain editable.'
);
