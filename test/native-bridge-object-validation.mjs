import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

if (process.platform !== 'win32') {
  console.log('Native bridge object validation skipped outside Windows.');
  process.exit(0);
}

const root = resolve(import.meta.dirname, '..');
const buildDirectory = await mkdtemp(join(tmpdir(), 'native-bridge-validation-'));
const outputPath = join(buildDirectory, 'native-bridge.mjs');
const chart = (overrides = {}) => ({
  id: 'chart:test',
  name: 'SalesChart',
  chartType: 51,
  sourceRangeRef: 'A1:B5',
  plotBy: 'columns',
  anchor: { left: 10, top: 10, width: 300, height: 200 },
  ...overrides,
});
const operation = chartValue => ({
  kind: 'updateChart',
  sheetName: 'Données',
  name: 'SalesChart',
  chart: chartValue,
  preserveAnchor: true,
  preserveSeries: true,
});

try {
  await build({
    entryPoints: [join(root, 'src', 'provider', 'nativeExcelBridge.ts')],
    outfile: outputPath,
    bundle: true,
    format: 'esm',
    platform: 'node',
    logLevel: 'silent',
    tsconfig: join(root, 'tsconfig.json'),
    plugins: [{
      name: 'mock-vscode',
      setup(buildApi) {
        buildApi.onResolve({ filter: /^vscode$/ }, () => ({
          path: 'vscode',
          namespace: 'mock-vscode',
        }));
        buildApi.onLoad(
          { filter: /.*/, namespace: 'mock-vscode' },
          () => ({ contents: 'export const extensions = { getExtension: () => undefined };' }),
        );
      },
    }],
  });
  const { validateNativeEdits } = await import(
    `${pathToFileURL(outputPath).href}?cache=${Date.now()}`
  );
  const [validated] = validateNativeEdits(
    'C:\\work\\book.xlsx',
    [operation(chart())],
  );
  assert.equal(validated.preserveAnchor, true);
  assert.equal(validated.preserveSeries, true);
	const [absoluteRangeValidated] = validateNativeEdits(
		'C:\\work\\book.xlsx',
		[operation(chart({ sourceRangeRef: '$A$1:$B$5' }))],
	);
	assert.equal(absoluteRangeValidated.chart.sourceRangeRef, 'A1:B5');
	const [inferredComboValidated] = validateNativeEdits(
		'C:\\work\\book.xlsx',
		[operation(chart({
			chartType: 51,
			sourceRangeRef: undefined,
			series: [
				{ id: 'series:column', valuesRange: 'B2:B5' },
				{ id: 'series:line', chartType: 4, valuesRange: 'C2:C5' },
			],
		}))],
	);
	assert.equal(inferredComboValidated.chart.chartType, -4152);
	assert.deepEqual(inferredComboValidated.chart.series.map(series => series.chartType), [51, 4]);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [{
			kind: 'createChart',
			sheetName: 'Données',
			chart: chart({ legend: { visible: true, position: 'custom' } }),
		}]),
		/cannot create a custom legend layout/,
	);
	assert.doesNotThrow(() => validateNativeEdits(
		'C:\\work\\book.xlsx',
		[operation(chart({ legend: { visible: true, position: 'custom' } }))],
	));
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
			sourceRangeRef: undefined,
			series: [{ id: 'series:column', valuesRange: 'B2:B5', bubbleSizesRange: 'C2:C5' }],
		}))]),
		/bubbleSizesRange requires a bubble chart type/,
	);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
			sourceRangeRef: undefined,
			series: [{ id: 'series:column', valuesRange: 'B2:B5', smooth: false }],
		}))]),
		/smooth requires a line or scatter chart type/,
	);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
			sourceRangeRef: undefined,
			series: [{ id: 'series:column', valuesRange: 'B2:B5', dataLabels: { showBubbleSize: true } }],
		}))]),
		/showBubbleSize requires a bubble chart type/,
	);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
			sourceRangeRef: undefined,
			series: [{ id: 'series:column', valuesRange: 'B2:B5', dataLabels: { showPercentage: true } }],
		}))]),
		/showPercentage requires a pie or doughnut chart type/,
	);
	assert.doesNotThrow(() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
		sourceRangeRef: undefined,
		series: [{
			id: 'series:column',
			valuesRange: 'B2:B5',
			dataLabels: { showBubbleSize: false, showPercentage: false },
		}],
	}))]));
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
			sourceRangeRef: undefined,
			series: [{
				id: 'series:disabled-position',
				valuesRange: 'B2:B5',
				dataLabels: { showValue: false, position: 'outsideEnd' },
			}],
		}))]),
		/position requires at least one enabled show option/,
	);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
			sourceRangeRef: undefined,
			series: [{ id: 'series:empty-labels', valuesRange: 'B2:B5', dataLabels: {} }],
		}))]),
		/dataLabels must explicitly define at least one show option/,
	);
	assert.doesNotThrow(() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
		chartType: 5,
		sourceRangeRef: undefined,
		series: [{ id: 'series:pie', valuesRange: 'B2:B5', dataLabels: { showPercentage: true } }],
	}))]));
	assert.doesNotThrow(() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
		chartType: 15,
		sourceRangeRef: undefined,
		series: [{ id: 'series:bubble', valuesRange: 'B2:B5', dataLabels: { showBubbleSize: true } }],
	}))]));
	const [clearAlternativeTextOperation] = validateNativeEdits(
		'C:\\work\\book.xlsx',
		[operation(chart({ alternativeText: '' }))],
	);
	assert.equal(clearAlternativeTextOperation.chart.alternativeText, '');
	const [clearAxisTitleOperation] = validateNativeEdits(
		'C:\\work\\book.xlsx',
		[operation(chart({ valueAxis: { visible: true, title: '' } }))],
	);
	assert.equal(clearAxisTitleOperation.chart.valueAxis.title, '');
	const [resetAxisFormatOperation] = validateNativeEdits(
		'C:\\work\\book.xlsx',
		[operation(chart({ valueAxis: { visible: true, numberFormat: '' } }))],
	);
	assert.equal(resetAxisFormatOperation.chart.valueAxis.numberFormat, '');
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
			sourceRangeRef: undefined,
			series: [{ id: 'series:position-only', valuesRange: 'B2:B5', dataLabels: { position: 'outsideEnd' } }],
		}))]),
		/dataLabels must explicitly define at least one show option/,
	);
	assert.doesNotThrow(() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
		sourceRangeRef: undefined,
		series: [{ id: 'series:column-position', valuesRange: 'B2:B5', dataLabels: { showValue: true, position: 'outsideEnd' } }],
	}))]));
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
			sourceRangeRef: undefined,
			series: [{ id: 'series:column-invalid-position', valuesRange: 'B2:B5', dataLabels: { showValue: true, position: 'above' } }],
		}))]),
		/data label position is not supported/,
	);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
			chartType: 52,
			sourceRangeRef: undefined,
			series: [{ id: 'series:stacked-invalid-position', valuesRange: 'B2:B5', dataLabels: { showValue: true, position: 'outsideEnd' } }],
		}))]),
		/data label position is not supported/,
	);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
			chartType: -4120,
			sourceRangeRef: undefined,
			series: [{ id: 'series:doughnut-invalid-position', valuesRange: 'B2:B5', dataLabels: { showValue: true, position: 'center' } }],
		}))]),
		/data label position is not supported/,
	);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
			chartType: -4152,
			sourceRangeRef: undefined,
			series: [
				{ id: 'series:column', chartType: 51, valuesRange: 'B2:B5' },
				{
					id: 'series:bubble',
					chartType: 15,
					xValuesRange: 'A2:A5',
					valuesRange: 'C2:C5',
					bubbleSizesRange: 'D2:D5',
				},
			],
		}))]),
		/cannot mix bubble and non-bubble series/,
	);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
			chartType: -4152,
			sourceRangeRef: undefined,
			series: [
				{ id: 'series:line', chartType: 4, valuesRange: 'B2:B5', smooth: true },
				{ id: 'series:column', chartType: 51, valuesRange: 'C2:C5', smooth: true },
			],
		}))]),
		/series 2 smooth requires a line or scatter chart type/,
	);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
			chartType: -4152,
			sourceRangeRef: undefined,
			categoryAxis: { visible: true, minimumScale: 0 },
			series: [
				{ id: 'series:column', chartType: 51, axisGroup: 'primary', valuesRange: 'B2:B5' },
				{ id: 'series:scatter', chartType: 74, axisGroup: 'secondary', xValuesRange: 'A2:A5', valuesRange: 'C2:C5' },
			],
		}))]),
		/categoryAxis numeric scale settings require a scatter or bubble chart/,
	);

  assert.throws(
    () => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
      anchor: { left: 0, top: 0, width: 19, height: 200 },
    }))]),
    /Native chart width is outside the supported range/,
  );
  assert.throws(
    () => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
      sourceRangeRef: undefined,
      series: [{
        id: 'series:test',
        name: 'Revenue',
        nameRange: 'B1',
        categoryRange: 'A2:A5',
        valuesRange: 'B2:B5',
      }],
    }))]),
    /cannot define both name and nameRange/,
  );
  assert.throws(
    () => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
      sourceRangeRef: undefined,
      series: [{
        id: 'series:test',
        categoryRange: 'A2:A5',
        xValuesRange: 'C2:C5',
        valuesRange: 'B2:B5',
      }],
    }))]),
    /cannot define both categoryRange and xValuesRange/,
  );
  assert.throws(
    () => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
      chartType: -2,
    }))]),
    /not permitted for offline native creation/,
  );
  assert.throws(
    () => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
      chartType: 140,
    }))]),
    /not permitted for offline native creation/,
  );
  assert.throws(
    () => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
      series: [{ id: 'series:test', valuesRange: 'B2:B5' }],
    }))]),
    /cannot define both sourceRangeRef and series/,
  );
  assert.throws(
    () => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
      sourceRangeRef: undefined,
      series: [{ id: 'series:test', nameRange: 'B1:B2', valuesRange: 'B2:B5' }],
    }))]),
    /nameRange must identify exactly one cell/,
  );
  assert.throws(
    () => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
      chartType: -4100,
      secondaryValueAxis: { visible: false },
    }))]),
    /secondary axes are not supported/,
  );
  assert.throws(
    () => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
      chartType: 5,
      categoryAxis: { visible: true },
    }))]),
    /axes are not supported/,
  );
  assert.throws(
    () => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
      chartType: 4,
      gapWidth: 150,
    }))]),
    /gapWidth is not supported/,
  );
  assert.throws(
    () => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
      chartType: -4152,
      sourceRangeRef: undefined,
      series: [{ id: 'series:test', valuesRange: 'B2:B5' }],
    }))]),
	/at least two explicit series with distinct concrete chartTypes/,
  );
	const [homogeneousComboOperation] = validateNativeEdits(
		'C:\\work\\book.xlsx',
		[operation(chart({
			chartType: -4152,
			sourceRangeRef: undefined,
			series: [
				{ id: 'series:one', chartType: 51, valuesRange: 'B2:B5' },
				{ id: 'series:two', chartType: 51, valuesRange: 'C2:C5' },
			],
		}))],
	);
	assert.equal(homogeneousComboOperation.chart.chartType, 51);
	const [singleSeriesOverrideOperation] = validateNativeEdits(
		'C:\\work\\book.xlsx',
		[operation(chart({
			chartType: 51,
			sourceRangeRef: undefined,
			series: [{ id: 'series:line', chartType: 4, valuesRange: 'B2:B5' }],
		}))],
	);
	assert.equal(singleSeriesOverrideOperation.chart.chartType, 4);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
			chartType: -4152,
			sourceRangeRef: undefined,
			series: [
				{ id: 'series:line', chartType: 4, valuesRange: 'B2:B5' },
				{ id: 'series:3d', chartType: -4100, axisGroup: 'secondary', valuesRange: 'C2:C5' },
			],
		}))]),
		/cannot use a secondary axis/,
	);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
			valueAxis: { visible: true, logarithmic: true, minimumScale: 0 },
		}))]),
		/logarithmic scale cannot use a non-positive bound/,
	);
	assert.doesNotThrow(() => validateNativeEdits('C:\\work\\book.xlsx', [operation(chart({
		chartType: -4100,
		gapWidth: 150,
	}))]));
	const unicodeTableOperation = (name, sheetName = 'Données', rangeRef = 'A1:C10') => ({
		kind: 'createTable',
		sheetName,
		table: {
			id: `table:${sheetName}:${rangeRef}`,
			name,
			displayName: name,
			rangeRef,
			headerRow: true,
			totalsRow: false,
			style: { name: 'TableStyleMedium2', showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false },
		},
	});
	const [unicodeTable] = validateNativeEdits(
		'C:\\work\\book.xlsx',
		[unicodeTableOperation('E\u0301quipe_2026')],
	);
	assert.equal(unicodeTable.table.name, 'Équipe_2026');
	assert.equal(unicodeTable.table.displayName, 'Équipe_2026');
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [{
			...unicodeTableOperation('NoHeaders'),
			table: { ...unicodeTableOperation('NoHeaders').table, headerRow: false },
		}]),
		/headerRow=false is disabled/,
	);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [{
			...unicodeTableOperation('NoNativeTotals'),
			table: { ...unicodeTableOperation('NoNativeTotals').table, totalsRow: true },
		}]),
		/totalsRow=true is disabled/,
	);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [{
			...unicodeTableOperation('TooShortTotals', 'Données', 'A1:C2'),
			table: {
				...unicodeTableOperation('TooShortTotals', 'Données', 'A1:C2').table,
				totalsRow: true,
			},
		}]),
		/does not contain enough rows/,
	);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [
			unicodeTableOperation('TooShortData', 'Données', 'A1:C1'),
		]),
		/does not contain enough rows/,
	);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [
			unicodeTableOperation('Équipe_2026', 'Données', 'A1:C10'),
			unicodeTableOperation('E\u0301QUIPE_2026', 'Archive', 'A1:C10'),
		]),
		/requested more than once in the workbook/,
		'canonical Unicode table-name duplicates must fail at the native bridge',
	);
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [{
			kind: 'createTable',
			sheetName: 'Données',
			table: {
				id: 'table:one', name: 'One', displayName: 'One', rangeRef: 'A1:C10',
				headerRow: true, totalsRow: false,
				style: { name: 'TableStyleMedium2', showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false },
			},
		}, {
			kind: 'createTable',
			sheetName: 'Données',
			table: {
				id: 'table:two', name: 'Two', displayName: 'Two', rangeRef: 'C10:E20',
				headerRow: true, totalsRow: false,
				style: { name: 'TableStyleMedium2', showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false },
			},
		}]),
		/Native tables One and Two overlap/,
	);
  assert.throws(
    () => validateNativeEdits('C:\\work\\book.xlsx', [{
      ...operation(chart()),
      preserveSeries: 'yes',
    }]),
    /Native updateChart preserveSeries must be a boolean/,
  );
	assert.throws(
		() => validateNativeEdits('C:\\work\\book.xlsx', [{
			...operation(chart()),
			preserveSeries: false,
			allowSeriesFormattingChange: true,
		}]),
		/allowSeriesFormattingChange requires preserveSeries=true/,
	);

  console.log('Native bridge object validation passed: preservation flags, dimensions and exclusive series fields.');
} finally {
  await rm(buildDirectory, { recursive: true, force: true });
}
