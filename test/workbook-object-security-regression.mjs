import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const read = relative => readFile(join(root, relative), 'utf8');

const [
  preflight,
  reader,
  officeContent,
  officeProvider,
  objectTool,
  nativeDiff,
  chartReader,
  nativeScript,
  chartDesigner,
  tableDesigner,
] = await Promise.all([
  read('src/react/view/excel/ooxml-zip-preflight.ts'),
  read('src/react/view/excel/excel_reader.ts'),
  read('src/provider/handlers/officeContent.ts'),
  read('src/provider/officeViewerProvider.ts'),
  read('src/excelAiVbaStudio/workbookObjectTool.ts'),
  read('src/react/view/excel/native_edit_diff.ts'),
  read('src/react/view/excel/chart-ooxml-reader.ts'),
  read('scripts/office-ai-apply-edits.ps1'),
  read('src/react/view/excel/chart-designer.tsx'),
  read('src/react/view/excel/table-designer.tsx'),
]);

assert.match(preflight, /validateOoxmlZipInflationBounds/);
assert.match(preflight, /new DecompressionStream\('deflate-raw'\)/);
assert.match(preflight, /produced > entry\.uncompressedSize/);
assert.match(preflight, /headerId === 0x0001 \|\| headerId === 0x7075/);
assert.ok(
  reader.indexOf('validateOoxmlZipInflationBounds(buffer, zipMetadata)')
    < reader.indexOf("import('jszip')"),
  'real streamed inflation bounds must be checked before JSZip loads the package',
);
const boundedReadBody = officeContent.match(
  /export async function readUriBytesWithLimit[\s\S]+?\n\}/,
)?.[0] ?? '';
const stableLocalReadBody = officeContent.match(
  /async function readStableLocalFileBytes[\s\S]+?\n\}/,
)?.[0] ?? '';
assert.match(boundedReadBody, /bytes\.byteLength > maximumBytes/);
assert.match(boundedReadBody, /uri\.scheme === 'file'/);
assert.match(stableLocalReadBody, /fs\.open\(filePath, 'r'\)/);
assert.match(stableLocalReadBody, /Buffer\.allocUnsafe\(before\.size\)/);
assert.ok(
  (stableLocalReadBody.match(/handle\.stat\(\)/g) ?? []).length >= 2,
  'local Office reads must verify size twice on the same open file handle',
);
assert.match(officeProvider, /const currentBytes = await readUriBytes\(uri\)/);
assert.match(officeContent, /bufferBase64:\s*bytesToPayloadBase64\(data\)/);
assert.doesNotMatch(
  officeContent.match(/export async function emitVirtualOfficeOpen[\s\S]+?\n\}/)?.[0] ?? '',
  /bytesToPayloadBuffer\(data\)/,
);
assert.ok(
  boundedReadBody.indexOf('workspace.fs.stat(uri)')
    < boundedReadBody.indexOf('workspace.fs.readFile(uri)'),
  'virtual Office files must be size-checked before the provider reads them into memory',
);

assert.match(objectTool, /confirmationMessages/);
assert.match(objectTool, /Chemin canonique complet/);
assert.match(objectTool, /canonicalizeWorkbookUri/);
assert.match(objectTool, /hors de l’espace de travail/);
assert.match(objectTool, /MAX_WORKBOOK_OBJECT_TRANSACTION_CELLS/);
assert.match(nativeDiff, /preserveAnchor:\s*true/);
assert.match(nativeDiff, /preserveSeries:\s*true/);
assert.match(chartReader, /PivotChart non éditable conservé nativement/);
assert.match(nativeScript, /function ConvertFrom-NativeSeriesFormula/);
assert.doesNotMatch(nativeScript, /function Test-NativeSeriesFormulaRange/);
assert.match(nativeScript, /xl\/charts\/\(\?:chart\|chartEx\|style\|colors\)/);
assert.match(nativeScript, /Get-NativePreservedChartSnapshot/);
assert.match(nativeScript, /NativeSmoothSeriesChartTypes/);
assert.match(nativeScript, /NativePercentageDataLabelChartTypes/);
assert.match(nativeScript, /function Test-NativeDataLabelPositionSupported/);
assert.match(nativeScript, /Native chart smooth requires a line or scatter series type/);
assert.match(nativeScript, /Native chart showBubbleSize requires a bubble series type/);
assert.match(nativeScript, /Native chart showPercentage requires a pie or doughnut series type/);
assert.match(nativeScript, /Native chart data label position is not supported by this chart type/);
assert.match(nativeScript, /dataLabels must explicitly define at least one show option/);
assert.match(nativeScript, /dataLabels\.position requires at least one enabled show option/);
assert.match(
	nativeScript,
	/if \(-not \$hasEnabledShowOption\) \{[\s\S]+?\$series\.HasDataLabels = \$false[\s\S]+?else \{[\s\S]+?\$series\.ApplyDataLabels\(\)/,
	'all-false labels must be removed without invoking ApplyDataLabels',
);
assert.match(nativeScript, /Chart data-label removal verification failed/);
assert.match(nativeScript, /NumberFormatLinked = \$true/);
assert.match(nativeScript, /NumberFormatLinked = \$false/);
assert.match(
	nativeScript,
	/NumberFormat = \$numberFormat[\s\S]+?NumberFormatLinked = \$false[\s\S]+?\$Definition\.numberFormat = \$normalizedNumberFormat/,
	'custom axis formats must be verified against Excel’s locale-normalized value',
);
assert.match(nativeScript, /number-format linked-state verification failed/);
assert.ok(
  nativeScript.indexOf('Assert-NativeChartDefinition $chartDefinition')
    < nativeScript.indexOf('$excel = New-Object -ComObject Excel.Application'),
  'series compatibility must be validated before Excel COM is created',
);
assert.match(
  nativeScript,
  /Has-Property \$seriesDefinition 'smooth'[\s\S]+?NativeSmoothSeriesChartTypes -contains \$effectiveSeriesType[\s\S]+?\$series\.Smooth/,
  'verification may access Series.Smooth only behind an effective-type guard',
);
assert.match(
  nativeScript,
  /\$propertyName -ceq 'showPercentage'[\s\S]+?NativePercentageDataLabelChartTypes -contains \$effectiveSeriesType[\s\S]+?continue/,
  'unsupported false percentage labels must not touch the COM property',
);
assert.match(
  nativeScript,
  /\$propertyName -ceq 'showBubbleSize'[\s\S]+?NativeBubbleChartTypes -contains \$effectiveSeriesType[\s\S]+?continue/,
  'unsupported false bubble-size labels must not touch the COM property',
);
assert.ok(
  nativeScript.indexOf('Native table creation with headerRow=false is disabled')
    < nativeScript.indexOf('$excel = New-Object -ComObject Excel.Application'),
  'headerless table creation must be rejected before Excel COM is created',
);
assert.ok(
  nativeScript.indexOf('Native table creation with totalsRow=true is disabled')
    < nativeScript.indexOf('$excel = New-Object -ComObject Excel.Application'),
  'totals-row creation must be rejected before Excel COM is created',
);
assert.match(nativeScript, /Native table totalsRow transitions are disabled/);
assert.match(nativeScript, /Native tables with an existing totals row cannot be resized/);
assert.doesNotMatch(nativeScript, /\.ShowTotals\s*=\s*\$(?:true|false)/i);
assert.match(nativeScript, /ShowHeaders -ne \[bool\]\$definition\.headerRow/);
assert.match(chartDesigner, /existingChart\?\.legend\?\.position === 'custom'/);
assert.match(
  chartDesigner,
  /disabled: option\.value === 'custom' && !canPreserveCustomLegend/,
);
assert.match(tableDesigner, /checked=\{draft\.headerRow\}[\s\S]+?disabled/);
assert.match(tableDesigner, /checked=\{draft\.totalsRow\}[\s\S]+?disabled/);
assert.match(nativeScript, /persistent backup alternate data streams are not a/);
assert.ok(
  nativeScript.indexOf('persistent backup alternate data streams are not a')
    < nativeScript.indexOf('$rollbackResult = Restore-DisplacedWorkbook'),
  'rollback must verify backup ADS before restoring the displaced workbook',
);

console.log('Workbook-object security regression checks passed.');
