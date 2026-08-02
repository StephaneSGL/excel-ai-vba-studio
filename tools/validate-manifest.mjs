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
  'excelAiVbaStudio.openSecurityCenter',
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
expect(tools.length === 4, 'exactly four language-model tools must be declared');
const readTool = tools.find(({ name }) => name === 'excel_ai_vba_readWorkbook');
const writeTool = tools.find(({ name }) => name === 'excel_ai_vba_writeModule');
const designTool = tools.find(({ name }) => name === 'excel_ai_vba_designWorkbook');
const workbookDesignTool = tools.find(
  ({ name }) => name === 'excel_ai_workbook_designObjects',
);
expect(readTool?.toolReferenceName === 'excelVbaWorkbook', 'unexpected workbook read tool reference');
expect(writeTool?.toolReferenceName === 'excelVbaWriteModule', 'unexpected VBA write tool reference');
expect(
  designTool?.toolReferenceName === 'excelVbaDesignWorkbook',
  'unexpected VBA designer tool reference',
);
expect(
  workbookDesignTool?.toolReferenceName === 'excelWorkbookDesign',
  'unexpected workbook-object designer tool reference',
);
expect(
  tools.every(({ canBeReferencedInPrompt }) => canBeReferencedInPrompt === true),
  'language-model tools must be prompt-referenceable',
);
expect(
  JSON.stringify(writeTool?.inputSchema?.required) === JSON.stringify(['componentFile', 'source']),
  'VBA write tool must require componentFile and source',
);
expect(
  JSON.stringify(designTool?.inputSchema?.required) === JSON.stringify(['operations']),
  'VBA designer tool must require operations',
);
expect(
  designTool?.inputSchema?.properties?.operations?.minItems === 1
    && designTool?.inputSchema?.properties?.operations?.maxItems === 100
    && designTool?.inputSchema?.properties?.operations?.items?.oneOf?.length === 7,
  'VBA designer operations schema must expose the seven bounded operation shapes',
);
expect(
  workbookDesignTool?.inputSchema?.properties?.operations?.minItems === 1
    && workbookDesignTool?.inputSchema?.properties?.operations?.maxItems === 100
    && workbookDesignTool?.inputSchema?.properties?.operations?.items?.oneOf?.length === 4,
  'workbook-object designer must expose four bounded conditional table/chart operation shapes',
);
expect(
  workbookDesignTool?.inputSchema?.required?.includes('workbookPath')
    && workbookDesignTool?.inputSchema?.properties?.workbookPath?.minLength >= 3,
  'mutating workbook-object tool must require the explicitly confirmed workbookPath',
);

const workbookOperations =
  workbookDesignTool?.inputSchema?.properties?.operations?.items?.oneOf ?? [];
const tableWriteOperation = workbookOperations.find(operation =>
  operation?.properties?.kind?.enum?.includes('createWorksheetTable')
    && operation?.properties?.kind?.enum?.includes('updateWorksheetTable')
);
const chartWriteOperation = workbookOperations.find(operation =>
  operation?.properties?.kind?.enum?.includes('createWorksheetChart')
    && operation?.properties?.kind?.enum?.includes('updateWorksheetChart')
);

function operationCondition(operation, kind) {
  return operation?.allOf?.[0]?.oneOf?.find(condition =>
    condition?.properties?.kind?.enum?.includes(kind)
  );
}

const createTableCondition = operationCondition(tableWriteOperation, 'createWorksheetTable');
const updateTableCondition = operationCondition(tableWriteOperation, 'updateWorksheetTable');
const createChartCondition = operationCondition(chartWriteOperation, 'createWorksheetChart');
const updateChartCondition = operationCondition(chartWriteOperation, 'updateWorksheetChart');
expect(
  createTableCondition?.not?.required?.includes('name')
    && updateTableCondition?.required?.includes('name'),
  'table create must reject top-level name and table update must require it',
);
expect(
  createChartCondition?.not?.required?.includes('name')
    && updateChartCondition?.required?.includes('name'),
  'chart create must reject top-level name and chart update must require it',
);

const chartSchema = chartWriteOperation?.properties?.chart;
const chartSeriesSchema = chartSchema?.properties?.series;
const a1Pattern = chartSchema?.properties?.sourceRangeRef?.pattern;
expect(
  chartSchema?.properties?.chartType?.enum?.length === 101
    && new Set(chartSchema.properties.chartType.enum).size === 101
    && !chartSchema.properties.chartType.enum.includes(-2)
    && !chartSchema.properties.chartType.enum.includes(140)
    && chartSeriesSchema?.items?.properties?.chartType?.enum?.length === 101
    && !chartSeriesSchema.items.properties.chartType.enum.includes(-2)
    && !chartSeriesSchema.items.properties.chartType.enum.includes(140),
  'writable chartType schemas must expose 101 offline-safe values and reject suggested and network-backed maps',
);
expect(
  typeof a1Pattern === 'string'
    && tableWriteOperation?.properties?.table?.properties?.rangeRef?.pattern === a1Pattern
    && ['categoryRange', 'valuesRange', 'xValuesRange', 'bubbleSizesRange']
      .every(property => chartSeriesSchema?.items?.properties?.[property]?.pattern === a1Pattern),
  'all table and chart ranges must use the same bounded local A1 pattern',
);
expect(
  chartSeriesSchema?.items?.properties?.nameRange?.pattern
    === '^\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]{0,6}$',
  'series nameRange must identify exactly one local A1 cell',
);
expect(
  chartSeriesSchema?.minItems === 1
    && chartSeriesSchema?.maxItems === 255
    && chartSchema?.anyOf?.some(option => option?.required?.includes('sourceRangeRef'))
    && chartSchema?.anyOf?.some(option => option?.required?.includes('series'))
    && chartSchema?.allOf?.some(condition =>
      condition?.not?.required?.includes('sourceRangeRef')
      && condition?.not?.required?.includes('series')
    ),
  'chart schema must require a source range or at least one explicit series',
);

const expectedAxisProperties = [
  'visible', 'title', 'minimumScale', 'maximumScale', 'majorUnit', 'minorUnit',
  'logarithmic', 'reverseOrder', 'numberFormat', 'majorGridlines', 'minorGridlines',
].sort();
expect(
  ['categoryAxis', 'valueAxis', 'secondaryCategoryAxis', 'secondaryValueAxis']
    .every(axisName => {
      const axis = chartSchema?.properties?.[axisName];
      return axis?.additionalProperties === false
        && JSON.stringify(Object.keys(axis?.properties ?? {}).sort())
          === JSON.stringify(expectedAxisProperties);
    }),
  'all four chart axes must expose the complete closed option set',
);

const dataLabelsSchema = chartSeriesSchema?.items?.properties?.dataLabels;
expect(
  dataLabelsSchema?.additionalProperties === false
    && JSON.stringify(Object.keys(dataLabelsSchema?.properties ?? {}).sort()) === JSON.stringify([
      'position', 'showBubbleSize', 'showCategoryName', 'showPercentage',
      'showSeriesName', 'showValue',
    ])
    && dataLabelsSchema?.properties?.position?.enum?.includes('outsideEnd'),
  'chart data labels must expose the complete closed option set',
);
expect(
  manifest.activationEvents.includes(
    'onLanguageModelTool:excel_ai_workbook_designObjects',
  ),
  'workbook-object designer activation event is missing',
);

const settings = manifest.contributes?.configuration?.properties ?? {};
expect(settings['excelAiVbaStudio.maxRows']?.default === 200, 'maxRows default must be 200');
expect(settings['excelAiVbaStudio.maxRows']?.maximum === 5000, 'maxRows maximum must be 5000');
expect(settings['excelAiVbaStudio.maxColumns']?.default === 50, 'maxColumns default must be 50');
expect(settings['excelAiVbaStudio.maxColumns']?.maximum === 256, 'maxColumns maximum must be 256');
expect(settings['excelAiVbaStudio.includeVba']?.default === false, 'includeVba must be opt-in');
expect(
  Array.isArray(settings['excelAiVbaStudio.allowedCustomActiveXProgIds']?.default)
    && settings['excelAiVbaStudio.allowedCustomActiveXProgIds'].default.length === 0
    && settings['excelAiVbaStudio.allowedCustomActiveXProgIds']?.maxItems === 32
    && settings['excelAiVbaStudio.allowedCustomActiveXProgIds']?.uniqueItems === true,
  'custom ActiveX ProgIDs must use an empty, bounded, unique opt-in allowlist',
);
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
  'scripts/inspect-office-security.ps1',
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
