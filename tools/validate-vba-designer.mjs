import assert from 'node:assert/strict';
import { existsSync, readFileSync, statSync } from 'node:fs';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const manifest = JSON.parse(readFileSync(resolve(root, 'package.json'), 'utf8'));
const types = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/types.ts'),
  'utf8',
);
const languageTool = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/languageModelTool.ts'),
  'utf8',
);
const service = readFileSync(
  resolve(root, 'src/excelAiVbaStudio/workbookService.ts'),
  'utf8',
);
const scriptPath = resolve(root, 'scripts/apply-vba-designer.ps1');
const helperPath = resolve(
  root,
  'bin/win32-x64/excel-ai-vba-writeback.exe',
);

assert.ok(existsSync(scriptPath), 'VBA designer PowerShell engine is missing');
assert.ok(existsSync(helperPath), 'native VBA inspection helper is missing');
assert.ok(
  statSync(helperPath).size > 1_000_000,
  'native VBA inspection helper is implausibly small',
);

const script = readFileSync(scriptPath, 'utf8');
assert.match(script, /AutomationSecurity\s*=\s*3/);
assert.match(script, /\[IO\.File\]::Replace/);
assert.match(
  script,
  /\$backupHash\s*=\s*Get-Sha256\s+\$backupPath[\s\S]+?\$backupHash\s+-cne\s+\$originalHash/,
  'designer commit must verify the displaced backup hash',
);
assert.match(
  script,
  /\$commitCompleted\s*=\s*\$true[\s\S]+?\$rollbackStagingPath[\s\S]+?\[IO\.File\]::Replace\([\s\S]+?\$rollbackStagingPath,[\s\S]+?\$workbookPath,[\s\S]+?\$failedReplacementPath[\s\S]+?ROLLBACK_OK/,
  'designer must restore the original after any post-replacement failure',
);
assert.match(script, /designerStreamsSha256/);
assert.match(
  script,
  /function Assert-MacroProcedureExists[\s\S]+?Public macro procedure[\s\S]+?Assert-MacroProcedureExists \$components \$macroName/,
  'worksheet buttons must target an existing public standard-module macro',
);
assert.match(
  script,
  /\$expectedButtons\s*=\s*\[System\.Collections\.Generic\.List\[object\]\]::new\(\)[\s\S]+?expectedButton\.sheetName/,
  'button verification must preserve worksheet names containing dots',
);
assert.match(script, /assignWorksheetButtonMacro/);
assert.match(script, /updateUserFormControl/);
assert.match(script, /setUserFormEventHandler/);
assert.match(script, /replaceExisting=true/);
assert.match(script, /expectedUpdatedControls/);
assert.match(script, /expectedEventHandlers/);
assert.match(script, /createWorksheetActiveXControl/);
assert.match(script, /bindWorksheetActiveXMacro/);
assert.match(script, /allowedCustomActiveXProgIds/);
assert.match(script, /Custom ActiveX ProgID is not allowlisted/);
assert.match(script, /ActiveX event handler[\s\S]+?refusing to overwrite/);
assert.match(script, /expectedActiveXControls/);
assert.match(script, /expectedActiveXBindings/);
assert.match(script, /OWNED_EXCEL_PID\|/);
assert.match(script, /requestSize\s+-gt\s+1MB/);
assert.doesNotMatch(script, /\.Run\s*\(/i, 'designer must never run a macro');
assert.doesNotMatch(script, /RunAutoMacros/i, 'designer must never run auto macros');
assert.doesNotMatch(
  script,
  /(?:New|Set)-ItemProperty|Microsoft\.Win32\.Registry/i,
  'designer must never change registry or AccessVBOM',
);

const tools = manifest.contributes?.languageModelTools ?? [];
const designTool = tools.find(
  ({ name }) => name === 'excel_ai_vba_designWorkbook',
);
assert.ok(designTool, 'VBA designer language-model tool is missing');
assert.equal(designTool.toolReferenceName, 'excelVbaDesignWorkbook');
assert.equal(designTool.canBeReferencedInPrompt, true);
assert.match(designTool.modelDescription, /\.xlsm/i);
assert.match(designTool.modelDescription, /sauvegarde/i);
assert.match(designTool.modelDescription, /Aucune macro/i);
assert.match(designTool.modelDescription, /AccessVBOM/i);
assert.ok(
  manifest.activationEvents.includes(
    'onLanguageModelTool:excel_ai_vba_designWorkbook',
  ),
  'VBA designer activation event is missing',
);

const inputSchema = designTool.inputSchema;
assert.equal(inputSchema.type, 'object');
assert.equal(inputSchema.additionalProperties, false);
assert.deepEqual(inputSchema.required, ['operations']);
const operationsSchema = inputSchema.properties.operations;
assert.equal(operationsSchema.type, 'array');
assert.equal(operationsSchema.minItems, 1);
assert.equal(operationsSchema.maxItems, 100);
assert.equal(operationsSchema.items.oneOf.length, 7);
for (const operationSchema of operationsSchema.items.oneOf) {
  assert.equal(operationSchema.type, 'object');
  assert.equal(operationSchema.additionalProperties, false);
}
const [
  createFormSchema,
  updateControlSchema,
  eventHandlerSchema,
  addControlSchema,
  createButtonSchema,
  assignOrBindSchema,
  createActiveXSchema,
] =
  operationsSchema.items.oneOf;
assert.deepEqual(createFormSchema.properties.kind.enum, ['createUserForm']);
assert.equal(
  createFormSchema.properties.controls.items.additionalProperties,
  false,
);
assert.deepEqual(updateControlSchema.properties.kind.enum, [
  'updateUserFormControl',
]);
assert.equal(updateControlSchema.properties.changes.minProperties, 1);
assert.equal(updateControlSchema.properties.changes.additionalProperties, false);
assert.deepEqual(eventHandlerSchema.properties.kind.enum, [
  'setUserFormEventHandler',
]);
assert.equal(eventHandlerSchema.properties.procedureSource.maxLength, 200000);
assert.equal(eventHandlerSchema.properties.replaceExisting.default, false);
assert.deepEqual(addControlSchema.properties.kind.enum, [
  'addUserFormControl',
]);
assert.equal(addControlSchema.properties.control.additionalProperties, false);
assert.deepEqual(createButtonSchema.properties.kind.enum, [
  'createWorksheetButton',
]);
assert.deepEqual(assignOrBindSchema.properties.kind.enum, [
  'assignWorksheetButtonMacro',
  'bindWorksheetActiveXMacro',
]);
assert.deepEqual(createActiveXSchema.properties.kind.enum, [
  'createWorksheetActiveXControl',
]);
assert.equal(
  createActiveXSchema.properties.control.additionalProperties,
  false,
);

assert.match(types, /EXCEL_AI_VBA_DESIGN_TOOL/);
assert.match(types, /VbaUserFormControlType/);
assert.match(types, /VbaDesignToolInput/);
assert.match(types, /designerVerified:\s*true/);
assert.match(types, /customActiveX/);
assert.match(types, /assignedButtons/);
assert.match(types, /updatedControls/);
assert.match(types, /updatedEventHandlers/);
assert.match(types, /createdActiveXControls/);
assert.match(types, /boundActiveXControls/);
assert.match(languageTool, /parseDesignInput/);
assert.match(languageTool, /MAX_VBA_DESIGN_OPERATIONS\s*=\s*100/);
assert.match(
  languageTool,
  /registerTool\(EXCEL_AI_VBA_DESIGN_TOOL,\s*designTool\)/,
);
assert.match(languageTool, /designVbaFromTool/);
assert.match(service, /async designVbaFromTool/);
assert.match(service, /expectedWorkbookSha256/);
assert.match(service, /MAX_VBA_DESIGN_REQUEST_BYTES/);
assert.match(service, /apply-vba-designer\.ps1/);
assert.match(service, /cleanupOwnedExcel:\s*true/);
assert.match(service, /hashFileSha256\(backupPath\)/);
assert.match(service, /includeVba:\s*true/);
assert.match(service, /allowedCustomActiveXProgIds/);
assert.match(service, /stdout\.matchAll\(\/OWNED_EXCEL_PID/);

assert.equal(
  manifest.scripts?.['validate:vba-designer'],
  'node tools/validate-vba-designer.mjs',
);
assert.equal(
  manifest.scripts?.['test:vba-designer'],
  'node test/vba-designer-integration.mjs',
);
assert.match(manifest.scripts?.validate ?? '', /validate:vba-designer/);
assert.match(manifest.scripts?.validate ?? '', /test:vba-designer/);

console.log(
  'VBA designer validation passed: visual control updates, bounded event handlers, transactional backup, native verification, no macro execution.',
);
