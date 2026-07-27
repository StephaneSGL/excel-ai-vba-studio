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
assert.match(script, /designerStreamsSha256/);
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
assert.equal(operationsSchema.items.oneOf.length, 3);
for (const operationSchema of operationsSchema.items.oneOf) {
  assert.equal(operationSchema.type, 'object');
  assert.equal(operationSchema.additionalProperties, false);
}
const [createFormSchema, addControlSchema, createButtonSchema] =
  operationsSchema.items.oneOf;
assert.deepEqual(createFormSchema.properties.kind.enum, ['createUserForm']);
assert.equal(
  createFormSchema.properties.controls.items.additionalProperties,
  false,
);
assert.deepEqual(addControlSchema.properties.kind.enum, [
  'addUserFormControl',
]);
assert.equal(addControlSchema.properties.control.additionalProperties, false);
assert.deepEqual(createButtonSchema.properties.kind.enum, [
  'createWorksheetButton',
]);

assert.match(types, /EXCEL_AI_VBA_DESIGN_TOOL/);
assert.match(types, /VbaUserFormControlType/);
assert.match(types, /VbaDesignToolInput/);
assert.match(types, /designerVerified:\s*true/);
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
  'VBA designer validation passed: bounded XLSM tool, transactional backup, native designer verification, no macro execution.',
);
