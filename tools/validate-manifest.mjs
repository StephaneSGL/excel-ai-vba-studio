import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import process from 'node:process';

const root = resolve(import.meta.dirname, '..');
const manifestPath = resolve(root, 'package.json');
const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));
const errors = [];

function expect(condition, message) {
  if (!condition) {
    errors.push(message);
  }
}

function sorted(values) {
  return [...values].sort((left, right) => left.localeCompare(right));
}

expect(manifest.name === 'excel-ai-vba-studio', 'name must be excel-ai-vba-studio');
expect(manifest.publisher === 'steph-tools', 'publisher must be steph-tools');
expect(manifest.displayName === 'Excel AI & VBA Studio', 'unexpected displayName');
expect(/^\d+\.\d+\.\d+$/.test(manifest.version), 'version must use major.minor.patch');
expect(manifest.preview === true, 'initial Marketplace listing must remain marked Preview');
expect(manifest.engines?.vscode === '>=1.95.0', 'engines.vscode must be >=1.95.0');
expect(JSON.stringify(manifest.extensionKind) === JSON.stringify(['ui']), 'extensionKind must be ["ui"]');
expect(manifest.main === './out/extension.js', 'desktop entry point must be ./out/extension.js');
expect(!Object.hasOwn(manifest, 'browser'), 'Windows-only extension must not expose a browser entry point');
expect(
  manifest.license === 'SEE LICENSE IN LICENSE',
  'mixed licensing must be described by the packaged LICENSE file',
);
expect(
  manifest.repository?.url === 'https://github.com/StephaneSGL/excel-ai-vba-studio.git',
  'repository must point to the public GitHub repository',
);
expect(
  manifest.homepage === 'https://github.com/StephaneSGL/excel-ai-vba-studio#readme',
  'homepage must point to the GitHub README',
);
expect(
  manifest.bugs?.url === 'https://github.com/StephaneSGL/excel-ai-vba-studio/issues',
  'bugs must point to GitHub Issues',
);
expect(manifest.icon === 'image/marketplace-icon.png', 'use the original Marketplace icon');
expect(existsSync(resolve(root, manifest.icon ?? '')), 'Marketplace icon does not exist');

const expectedCommands = sorted([
  'excelAiVbaStudio.exportWorkbook',
  'excelAiVbaStudio.copyWorkbookContext',
  'excelAiVbaStudio.openWorkbookContext',
  'excelAiVbaStudio.copyGeneratedContext',
  'excelAiVbaStudio.openExcel',
  'excelAiVbaStudio.openVbe',
  'excelAiVbaStudio.openVbaDeveloper',
  'excelAiVbaStudio.openVbaExplorer',
  'excelAiVbaStudio.askCopilotAboutWorkbook',
  'excelAiVbaStudio.refreshExplorer',
  'excelAiVbaStudio.cleanExports',
]);
const actualCommands = sorted((manifest.contributes?.commands ?? []).map(({ command }) => command));
expect(
  JSON.stringify(actualCommands) === JSON.stringify(expectedCommands),
  `commands differ from the supported set:\nexpected ${expectedCommands.join(', ')}\nactual   ${actualCommands.join(', ')}`,
);
expect(actualCommands.every((command) => command.startsWith('excelAiVbaStudio.')), 'all command IDs must use the new prefix');

const forbiddenContributions = [
  'authentication',
  'grammars',
  'iconThemes',
  'languages',
  'snippets',
  'themes',
  'walkthroughs',
];
for (const contribution of forbiddenContributions) {
  expect(!Object.hasOwn(manifest.contributes ?? {}, contribution), `forbidden contribution: ${contribution}`);
}

const customEditors = manifest.contributes?.customEditors ?? [];
expect(customEditors.length === 1, 'exactly one Excel/CSV custom editor must be declared');
expect(customEditors[0]?.viewType === 'excelAiVbaStudio.officeViewer', 'unexpected custom editor viewType');
const allowedPatterns = sorted(['*.xlsx', '*.xlsm', '*.xls', '*.csv', '*.tsv']);
const actualPatterns = sorted((customEditors[0]?.selector ?? []).map(({ filenamePattern }) => filenamePattern));
expect(
  JSON.stringify(actualPatterns) === JSON.stringify(allowedPatterns),
  `custom editor selectors must be limited to ${allowedPatterns.join(', ')}`,
);

const explorerViews = manifest.contributes?.views?.explorer ?? [];
expect(
  explorerViews.length === 2
    && explorerViews.some(({ id }) => id === 'excelAiVbaExplorer')
    && explorerViews.some(({ id }) => id === 'excelAiVbaProperties'),
  'Excel & VBA explorer and properties views are required',
);

const tools = manifest.contributes?.languageModelTools ?? [];
expect(tools.length === 2, 'exactly two language-model tools must be declared');
const readTool = tools.find(({ name }) => name === 'excel_ai_vba_readWorkbook');
const writeTool = tools.find(({ name }) => name === 'excel_ai_vba_writeModule');
expect(readTool?.toolReferenceName === 'excelVbaWorkbook', 'unexpected workbook read tool reference');
expect(writeTool?.toolReferenceName === 'excelVbaWriteModule', 'unexpected VBA write tool reference');
expect(
  tools.every(({ canBeReferencedInPrompt }) => canBeReferencedInPrompt === true),
  'language-model tools must be prompt-referenceable',
);
expect(
  JSON.stringify(writeTool?.inputSchema?.required) === JSON.stringify(['componentFile', 'source']),
  'VBA write tool must require componentFile and source',
);

const settings = manifest.contributes?.configuration?.properties ?? {};
expect(settings['excelAiVbaStudio.maxRows']?.default === 200, 'maxRows default must be 200');
expect(settings['excelAiVbaStudio.maxRows']?.maximum === 5000, 'maxRows maximum must be 5000');
expect(settings['excelAiVbaStudio.maxColumns']?.default === 50, 'maxColumns default must be 50');
expect(settings['excelAiVbaStudio.maxColumns']?.maximum === 256, 'maxColumns maximum must be 256');
expect(settings['excelAiVbaStudio.includeVba']?.default === false, 'includeVba must be opt-in');
expect(!JSON.stringify(settings).toLowerCase().includes('telemetry'), 'telemetry setting must not exist');

const allDependencies = {
  ...(manifest.dependencies ?? {}),
  ...(manifest.devDependencies ?? {}),
};
expect(!Object.hasOwn(allDependencies, '@vscode/extension-telemetry'), 'telemetry dependency must not exist');
expect(
  allDependencies.xlsx === 'https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz',
  'xlsx must use the pinned official SheetJS 0.20.3 package',
);

const requiredFiles = [
  'LICENSE',
  'LICENSING.md',
  'LICENSES/POLYFORM-NONCOMMERCIAL-1.0.0.md',
  'LICENSES/OFFICE-VIEWER-MIT.txt',
  'LICENSES/X-DATA-SPREADSHEET-MIT.txt',
  'README.md',
  'CHANGELOG.md',
  'SUPPORT.md',
  'PRIVACY.md',
  'NOTICE.md',
  'THIRD_PARTY_NOTICES.md',
  'scripts/office-ai-export.ps1',
  'scripts/open-excel-developer.ps1',
];
for (const file of requiredFiles) {
  expect(existsSync(resolve(root, file)), `required file is missing: ${file}`);
}

if (errors.length > 0) {
  console.error(`Manifest validation failed (${errors.length}):`);
  for (const error of errors) {
    console.error(`- ${error}`);
  }
  process.exit(1);
}

console.log(`Manifest valid: ${manifest.publisher}.${manifest.name}@${manifest.version} (Preview, win32-x64 release target)`);
