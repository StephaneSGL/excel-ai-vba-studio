import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const read = path => readFileSync(resolve(root, path), 'utf8');
const ribbon = read('src/react/view/excel/ExcelRibbon.tsx');
const spreadsheet = read('src/react/view/excel/x-spreadsheet/index.ts');
const conditionalFormatting = read('src/react/view/excel/x-spreadsheet/core/conditional_format.js');
const table = read('src/react/view/excel/x-spreadsheet/component/table.js');
const excel = read('src/react/view/excel/Excel.tsx');
const handler = read('src/provider/compress/commonHandler.ts');
const workbookService = read('src/excelAiVbaStudio/workbookService.ts');

assert.doesNotMatch(
  ribbon,
  /Bient[oô]t|Coming soon/i,
  'the ribbon must not expose unfinished-feature badges',
);

for (const preset of [
  'greaterThan',
  'lessThan',
  'equal',
  'containsText',
  'colorScale',
  'dataBar',
  'iconSet',
]) {
  assert.ok(ribbon.includes(preset), `conditional formatting is missing ${preset}`);
}

for (const operation of [
  'formatSelectionAsTable',
  'textToColumns',
  'removeDuplicateRows',
  'addSubtotal',
  'addForecastRow',
  'formulaAudit',
  'insertImage',
]) {
  assert.ok(
    spreadsheet.includes(`${operation}(`),
    `integrated spreadsheet operation is missing: ${operation}`,
  );
}

assert.match(
  conditionalFormatting,
  /const icon = \(glyph, color\) => \(\{ glyph, color \}\)/,
  'icon-set formatting must provide an explicit glyph and colour',
);
assert.match(
  table,
  /fillStyle: icon\.color[\s\S]+?fillText\(icon\.glyph/,
  'the canvas must draw conditional-format icons using their explicit colour',
);

assert.match(
  excel,
  /askCopilotAboutWorkbook', request/,
  'the ribbon request must be forwarded by the webview',
);
assert.match(
  handler,
  /askCopilotAboutWorkbook[\s\S]+?resourceUri: uri,[\s\S]+?request/,
  'the webview handler must forward targeted requests to the extension',
);
assert.match(
  ribbon,
  /onOpenExcel[\s\S]+?Open in Excel[\s\S]+?onOpenVbe[\s\S]+?Open native VBE/,
  'the ribbon must expose explicit native Excel and VBE handoff actions',
);
assert.match(
  excel,
  /handler\.emit\('openExcel'\)[\s\S]+?handler\.emit\('openVbe'\)/,
  'the webview must forward native Excel and VBE handoff actions',
);
assert.match(
  handler,
  /openExcel[\s\S]+?excelAiVbaStudio\.openExcel[\s\S]+?openVbe[\s\S]+?excelAiVbaStudio\.openVbe/,
  'the extension handler must route native Excel and VBE commands',
);
assert.match(
  workbookService,
  /async openExcel\([\s\S]+?showVbe = false[\s\S]+?runExcelLauncher/,
  'the workbook service must implement native Excel and VBE handoff',
);
assert.match(
  workbookService,
  /requestedTask[\s\S]+?T[aâ]che demand[ée]e/i,
  'targeted ribbon requests must be added to the Copilot prompt',
);

console.log('Integrated feature validation passed: conditional formatting and ribbon actions are active.');
