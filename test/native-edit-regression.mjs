import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const root = resolve(import.meta.dirname, '..');
const buildDirectory = await mkdtemp(join(tmpdir(), 'excel-native-diff-test-'));
const outputPath = join(buildDirectory, 'native-edit-diff.mjs');

function sheet(name, cells, styles = [], extra = {}) {
  const rows = {};
  for (const [position, cell] of Object.entries(cells)) {
    const [row, column] = position.split(':').map(Number);
    rows[row] ??= { cells: {} };
    rows[row].cells[column] = cell;
  }
  return { name, rows, styles, ...extra };
}

function table(id, name, rangeRef) {
  return {
    id,
    name,
    displayName: name,
    rangeRef,
    headerRow: true,
    totalsRow: false,
    style: {
      name: 'TableStyleMedium2',
      showFirstColumn: false,
      showLastColumn: false,
      showRowStripes: true,
      showColumnStripes: false,
    },
  };
}

function chart(id, name, overrides = {}) {
  return {
    id,
    name,
    chartType: 51,
    sourceRangeRef: 'A1:C5',
    plotBy: 'columns',
    anchor: { left: 500, top: 20, width: 420, height: 260 },
    title: { visible: true, text: 'Revenue' },
    legend: { visible: true, position: 'right' },
    style: 10,
    roundedCorners: false,
    ...overrides,
  };
}

const highlightStyle = {
  style: {
    fill: {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFC7CE' },
    },
    font: { color: { argb: 'FF9C0006' } },
  },
  displayStyle: {
    bgcolor: '#ffc7ce',
    color: '#9c0006',
    font: { bold: true },
  },
};

const generatedConditionalDefinitions = [
  {
    ref: 'A1:A10',
    rules: [{
      type: 'cellIs',
      operator: 'greaterThan',
      formulae: [10],
      ...highlightStyle,
      priority: 1,
    }],
  },
  {
    ref: 'B1:B10',
    rules: [{
      type: 'containsText',
      operator: 'containsText',
      text: 'Budget',
      formulae: ['Budget'],
      ...highlightStyle,
      priority: 2,
    }],
  },
  {
    ref: 'C1:C10',
    rules: [{
      type: 'colorScale',
      cfvo: [
        { type: 'min' },
        { type: 'percentile', value: 50 },
        { type: 'max' },
      ],
      color: [
        { argb: 'FFF8696B' },
        { argb: 'FFFFEB84' },
        { argb: 'FF63BE7B' },
      ],
      priority: 3,
    }],
  },
  {
    ref: 'D1:D10',
    rules: [{
      type: 'dataBar',
      cfvo: [{ type: 'min' }, { type: 'max' }],
      color: { argb: 'FF5B9BD5' },
      priority: 4,
    }],
  },
  {
    ref: 'E1:E10',
    rules: [{
      type: 'iconSet',
      iconSet: '3TrafficLights1',
      cfvo: [
        { type: 'min' },
        { type: 'percent', value: 33 },
        { type: 'percent', value: 67 },
      ],
      priority: 5,
    }],
  },
];

try {
  await build({
    entryPoints: [
      join(root, 'src', 'react', 'view', 'excel', 'native_edit_diff.ts'),
    ],
    outfile: outputPath,
    bundle: true,
    format: 'esm',
    platform: 'node',
    logLevel: 'silent',
    tsconfig: join(root, 'tsconfig.json'),
  });

  const {
    buildNativeCellEditOperations,
    buildNativeExcelEditPlan,
    initializeNativeEditSheets,
  } = await import(
    `${pathToFileURL(outputPath).href}?cache=${Date.now()}`
  );

  let displayedSheets = [];
  const normalizingSpreadsheet = {
    loadData(inputSheets) {
      displayedSheets = structuredClone(inputSheets).map((inputSheet) => {
        const {
          images: _images,
          pageSetup: _pageSetup,
          ...normalizedSheet
        } = inputSheet;
        return normalizedSheet;
      });
    },
    getData() {
      return displayedSheets;
    },
  };
  const featureRichSource = [
    sheet('Instructions', { '0:0': { text: 'Old' } }, [], {
      images: [{ id: 'existing-shape' }],
      pageSetup: { orientation: 'landscape' },
    }),
  ];
  const normalizedBaseline = initializeNativeEditSheets(
    normalizingSpreadsheet,
    featureRichSource,
  );
  const supportedCellEdit = structuredClone(displayedSheets);
  supportedCellEdit[0].rows[0].cells[0].text = 'New';
  assert.deepEqual(
    buildNativeExcelEditPlan(normalizedBaseline, supportedCellEdit),
    {
      operations: [{
        sheetName: 'Instructions',
        row: 1,
        column: 1,
        value: { kind: 'text', value: 'New' },
      }],
      unsupportedChanges: [],
    },
    'reader-only worksheet features must not become false edits after hydration',
  );

  let restoredFeatureSheets = [];
  const preservingSpreadsheet = {
    loadData(inputSheets) {
      restoredFeatureSheets = structuredClone(inputSheets);
    },
    getData() {
      return restoredFeatureSheets;
    },
  };
  const legacyFeatureSource = [
    sheet('Instructions', {}, [], {
      autofilter: {
        ref: 'C8:C13',
        filters: [{ ci: 2, operator: 'in', value: ['Named Excel Tables'] }],
        sort: { ci: 2, order: 'asc' },
      },
      images: [{ id: 'existing-instruction-image' }],
    }),
    sheet('SalesOrders', {}, [], {
      autofilter: { ref: 'A1:G44', filters: [], sort: null },
    }),
  ];
  const legacyFeatureBaseline = initializeNativeEditSheets(
    preservingSpreadsheet,
    legacyFeatureSource,
  );
  const nativeObjectEdit = structuredClone(restoredFeatureSheets);
  nativeObjectEdit[0].tables = [table('table:instructions:A20:C30', 'HelpTable', 'A20:C30')];
  nativeObjectEdit[1].charts = [chart('chart:sales:Revenue', 'Revenue')];
  const nativeObjectPlan = buildNativeExcelEditPlan(
    legacyFeatureBaseline,
    nativeObjectEdit,
  );
  assert.deepEqual(nativeObjectPlan.unsupportedChanges, []);
  assert.deepEqual(
    nativeObjectPlan.operations.map(operation => operation.kind),
    ['createTable', 'createChart'],
    'native tables/charts must not mutate filters, sorting, images, or other worksheet features',
  );

  assert.deepEqual(
    buildNativeCellEditOperations(
      [sheet('Home', { '0:0': { text: 'Old' } })],
      [sheet('Home', { '0:0': { text: 'New' } })],
    ),
    [{
      sheetName: 'Home',
      row: 1,
      column: 1,
      value: { kind: 'text', value: 'New' },
    }],
  );

  assert.deepEqual(
    buildNativeCellEditOperations(
      [sheet('Home', { '1:2': { text: '1' } })],
      [sheet('Home', { '1:2': { text: '=SUM(A1:A2)' } })],
    )[0].value,
    { kind: 'formula', value: '=SUM(A1:A2)' },
  );

  assert.deepEqual(
    buildNativeCellEditOperations(
      [sheet('Home', { '0:0': { text: '' } })],
      [sheet('Home', { '0:0': { text: '1.0' } })],
    )[0].value,
    { kind: 'number', value: 1 },
  );

  assert.deepEqual(
    buildNativeCellEditOperations(
      [sheet('Home', { '0:0': { text: '' } })],
      [sheet('Home', { '0:0': { text: '00123' } })],
    )[0].value,
    { kind: 'text', value: '00123' },
  );

  assert.deepEqual(
    buildNativeCellEditOperations(
      [sheet('Home', { '0:0': { text: 'Delete me' } })],
      [sheet('Home', {})],
    )[0].value,
    { kind: 'blank' },
  );

  const styleOperation = buildNativeCellEditOperations(
    [
      sheet(
        'Home',
        { '0:0': { text: 'A', style: 0 } },
        [{ align: 'left', font: { bold: false }, bgcolor: '#ffffff' }],
      ),
    ],
    [
      sheet(
        'Home',
        { '0:0': { text: 'A', style: 0 } },
        [{ align: 'center', font: { bold: true }, bgcolor: '#ff0000' }],
      ),
    ],
  )[0];
  assert.deepEqual(styleOperation.style, {
    align: 'center',
    bgcolor: '#ff0000',
    font: { bold: true },
  });

  assert.deepEqual(
    buildNativeCellEditOperations(
      [sheet('Home', { '0:0': { text: 'Same' } })],
      [sheet('Home', { '0:0': { text: 'Same' } })],
    ),
    [],
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', { '0:0': { text: 'Old' } })],
      [sheet('Home', { '0:0': { text: 'New' } })],
    ).unsupportedChanges,
    [],
    'ordinary value and style edits must remain natively saveable',
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], {
        hasNativeCharts: true,
        hasNativeChartParts: true,
        unsupportedNativeChartCount: 2,
      })],
      [sheet('Home', {})],
    ),
    { operations: [], unsupportedChanges: [] },
    'native chart reader metadata must not become a worksheet feature edit',
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {})],
      [sheet('Renamed', {})],
    ).unsupportedChanges,
    ['worksheets'],
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], { cols: { len: 2, 0: { width: 80 } } })],
      [sheet('Home', {}, [], { cols: { len: 2, 0: { width: 120 } } })],
    ),
    {
      operations: [{
        kind: 'columnWidth',
        sheetName: 'Home',
        column: 1,
        widthPx: 120,
      }],
      unsupportedChanges: [],
    },
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', { '0:0': { text: 'A' } })],
      [{
        ...sheet('Home', { '0:0': { text: 'A' } }),
        rows: { len: 1, 0: { height: 36, cells: { 0: { text: 'A' } } } },
      }],
    ),
    {
      operations: [{
        kind: 'rowHeight',
        sheetName: 'Home',
        row: 1,
        heightPx: 36,
      }],
      unsupportedChanges: [],
    },
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], { cols: { 0: { width: 80 } } })],
      [sheet('Home', {})],
    ).unsupportedChanges,
    ['Home:column-dimensions'],
    'implicit dimension resets must remain blocked',
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], { cols: { 0: { width: 80 } } })],
      [sheet('Home', {}, [], { cols: { 0: { width: 2_000 } } })],
    ).unsupportedChanges,
    ['Home:column-dimensions'],
    'column widths outside Excel limits must remain blocked',
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], { cols: { 0: { width: 80 } } })],
      [sheet('Home', {}, [], { cols: { 0: { width: 5 } } })],
    ).unsupportedChanges,
    ['Home:column-dimensions'],
    'a native width of zero must never hide a column implicitly',
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [{
        ...sheet('Home', {}),
        rows: { len: 1, 0: { height: 36, cells: {} } },
      }],
      [sheet('Home', {})],
    ).unsupportedChanges,
    ['Home:row-dimensions'],
    'implicit row-height resets must remain blocked',
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {})],
      [{
        ...sheet('Home', {}),
        rows: { len: 1, 0: { height: 600, cells: {} } },
      }],
    ).unsupportedChanges,
    ['Home:row-dimensions'],
    'row heights outside Excel limits must remain blocked',
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', { '0:0': { text: 'A' } })],
      [sheet('Home', { '0:0': { text: 'A', merge: [0, 1] } })],
    ).unsupportedChanges,
    ['Home:cell-structure'],
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {})],
      [sheet('Home', {}, [], {
        conditionalFormattings: [{ ref: 'A1', rules: [{ type: 'expression' }] }],
      })],
    ).unsupportedChanges,
    ['Home:conditional-formatting'],
  );

  const conditionalPlan = buildNativeExcelEditPlan(
    [sheet('Home', {})],
    [sheet('Home', {}, [], {
      conditionalFormattings: generatedConditionalDefinitions,
    })],
  );
  assert.deepEqual(conditionalPlan.unsupportedChanges, []);
  assert.deepEqual(conditionalPlan.operations, [
    {
      kind: 'addConditionalFormatting',
      sheetName: 'Home',
      rangeRef: 'A1:A10',
      rule: {
        type: 'cellIs',
        operator: 'greaterThan',
        operand: 10,
        fillColor: '#ffc7ce',
        fontColor: '#9c0006',
        bold: true,
      },
    },
    {
      kind: 'addConditionalFormatting',
      sheetName: 'Home',
      rangeRef: 'B1:B10',
      rule: {
        type: 'containsText',
        text: 'Budget',
        fillColor: '#ffc7ce',
        fontColor: '#9c0006',
        bold: true,
      },
    },
    {
      kind: 'addConditionalFormatting',
      sheetName: 'Home',
      rangeRef: 'C1:C10',
      rule: {
        type: 'colorScale',
        colors: ['#f8696b', '#ffeb84', '#63be7b'],
      },
    },
    {
      kind: 'addConditionalFormatting',
      sheetName: 'Home',
      rangeRef: 'D1:D10',
      rule: { type: 'dataBar', color: '#5b9bd5' },
    },
    {
      kind: 'addConditionalFormatting',
      sheetName: 'Home',
      rangeRef: 'E1:E10',
      rule: {
        type: 'iconSet',
        iconSet: '3TrafficLights1',
        thresholds: [33, 67],
      },
    },
  ]);

  const excessiveDefinitions = Array.from({ length: 65 }, (_, index) => ({
    ...structuredClone(generatedConditionalDefinitions[0]),
    ref: `A${index + 1}`,
    rules: [{
      ...structuredClone(generatedConditionalDefinitions[0].rules[0]),
      priority: index + 1,
    }],
  }));
  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {})],
      [sheet('Home', {}, [], {
        conditionalFormattings: excessiveDefinitions,
      })],
    ).unsupportedChanges,
    ['Home:conditional-formatting'],
    'one save cannot append an excessive conditional-rule batch',
  );

  const existingDefinition = {
    ref: 'F1:F10',
    rules: [{ type: 'expression', formulae: ['F1>0'], priority: 1 }],
  };
  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], {
        conditionalFormattings: [existingDefinition],
      })],
      [sheet('Home', {}, [], {
        conditionalFormattings: [
          existingDefinition,
          generatedConditionalDefinitions[0],
        ],
      })],
    ).unsupportedChanges,
    [],
    'an opaque existing rule may remain as an exact prefix',
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], {
        conditionalFormattings: [existingDefinition],
      })],
      [sheet('Home', {}, [], { conditionalFormattings: [] })],
    ).operations,
    [{ kind: 'clearConditionalFormatting', sheetName: 'Home' }],
  );

  const modifiedDefinition = structuredClone(generatedConditionalDefinitions[0]);
  modifiedDefinition.rules[0].formulae = [20];
  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], {
        conditionalFormattings: [generatedConditionalDefinitions[0]],
      })],
      [sheet('Home', {}, [], {
        conditionalFormattings: [modifiedDefinition],
      })],
    ).unsupportedChanges,
    ['Home:conditional-formatting'],
    'editing an existing rule must remain blocked',
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], {
        conditionalFormattings: generatedConditionalDefinitions.slice(0, 2),
      })],
      [sheet('Home', {}, [], {
        conditionalFormattings: [
          generatedConditionalDefinitions[1],
          generatedConditionalDefinitions[0],
        ],
      })],
    ).unsupportedChanges,
    ['Home:conditional-formatting'],
    'reordering rules must remain blocked',
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], {
        conditionalFormattings: generatedConditionalDefinitions.slice(0, 2),
      })],
      [sheet('Home', {}, [], {
        conditionalFormattings: generatedConditionalDefinitions.slice(0, 1),
      })],
    ).unsupportedChanges,
    ['Home:conditional-formatting'],
    'deleting one rule must remain blocked',
  );

  const disjointTables = [
    table('table:home:A1:C5', 'TableTop', 'A1:C5'),
    table('table:home:A20:C30', 'TableMiddle', 'A20:C30'),
    table('table:home:A40:C50', 'TableBottom', 'A40:C50'),
  ];
  const disjointTablePlan = buildNativeExcelEditPlan(
    [sheet('Home', {})],
    [sheet('Home', {}, [], { tables: disjointTables })],
  );
  assert.deepEqual(disjointTablePlan.unsupportedChanges, []);
  assert.deepEqual(
    disjointTablePlan.operations,
    disjointTables.map(item => ({
      kind: 'createTable',
      sheetName: 'Home',
      table: item,
    })),
    'every disjoint range must become its own ListObject operation',
  );

  const renamedTable = {
    ...disjointTables[1],
    name: 'TableMiddleRenamed',
    displayName: 'TableMiddleRenamed',
    totalsRow: true,
  };
  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], { tables: disjointTables })],
      [sheet('Home', {}, [], {
        tables: [disjointTables[0], renamedTable, disjointTables[2]],
      })],
    ).operations,
    [{
      kind: 'updateTable',
      sheetName: 'Home',
      name: 'TableMiddle',
      table: renamedTable,
    }],
    'a table rename must use its stable id and current Excel name',
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], { tables: disjointTables })],
      [sheet('Home', {}, [], { tables: disjointTables.slice(1) })],
    ).operations,
    [{ kind: 'deleteTable', sheetName: 'Home', name: 'TableTop' }],
  );

  const largeTable = table('table:large', 'LargeNativeTable', 'A1:Z10000');
  const largeTablePlan = buildNativeExcelEditPlan(
    [sheet('Large', {}, [{ bgcolor: '#ff0000' }])],
    [sheet('Large', {}, [{ bgcolor: '#ff0000' }], { tables: [largeTable] })],
  );
  assert.deepEqual(
    largeTablePlan.operations,
    [{ kind: 'createTable', sheetName: 'Large', table: largeTable }],
    'a large native table must remain one object operation and must not synthesize cell-style edits',
  );

  const originalChart = chart('chart:home:SalesChart', 'SalesChart');
  const updatedChart = chart('chart:home:SalesChart', 'SalesChartRenamed', {
    chartType: 65,
    plotBy: 'rows',
    title: { visible: true, text: 'Revenue trend' },
    legend: { visible: true, position: 'bottom' },
    series: [{
      id: 'series:revenue',
      name: 'Revenue',
      categoryRange: 'A2:A5',
      valuesRange: 'B2:B5',
      chartType: 65,
      axisGroup: 'primary',
      lineColor: '#3366cc',
      markerStyle: 'circle',
      markerSize: 7,
      smooth: true,
    }],
  });
  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], { charts: [originalChart] })],
      [sheet('Home', {}, [], { charts: [updatedChart] })],
    ).operations,
    [{
      kind: 'updateChart',
      sheetName: 'Home',
      name: 'SalesChart',
      chart: updatedChart,
      preserveAnchor: true,
    }],
    'a chart type, series and rename update must remain one stable-id operation',
  );

  const titleOnlyChart = structuredClone(originalChart);
  titleOnlyChart.title = { visible: true, text: 'Title only' };
  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], { charts: [originalChart] })],
      [sheet('Home', {}, [], { charts: [titleOnlyChart] })],
    ).operations,
    [{
      kind: 'updateChart',
      sheetName: 'Home',
      name: 'SalesChart',
      chart: titleOnlyChart,
      preserveAnchor: true,
      preserveSeries: true,
    }],
    'a title-only update must preserve native geometry and unmodelled series features',
  );

	const describedChart = { ...structuredClone(originalChart), alternativeText: 'Sales chart' };
	const clearedDescriptionChart = { ...structuredClone(describedChart), alternativeText: '' };
	const [clearDescriptionOperation] = buildNativeExcelEditPlan(
		[sheet('Home', {}, [], { charts: [describedChart] })],
		[sheet('Home', {}, [], { charts: [clearedDescriptionChart] })],
	).operations;
	assert.equal(clearDescriptionOperation.kind, 'updateChart');
	assert.equal(clearDescriptionOperation.chart.alternativeText, '');
	assert.equal(clearDescriptionOperation.preserveAnchor, true);
	assert.equal(clearDescriptionOperation.preserveSeries, true);

	const titledAxisChart = {
		...structuredClone(originalChart),
		valueAxis: { visible: true, title: 'Revenue' },
	};
	const clearedAxisTitleChart = {
		...structuredClone(originalChart),
		valueAxis: { visible: true, title: '' },
	};
	const [clearAxisTitleOperation] = buildNativeExcelEditPlan(
		[sheet('Home', {}, [], { charts: [titledAxisChart] })],
		[sheet('Home', {}, [], { charts: [clearedAxisTitleChart] })],
	).operations;
	assert.equal(clearAxisTitleOperation.kind, 'updateChart');
	assert.equal(clearAxisTitleOperation.chart.valueAxis.title, '');
	assert.equal(clearAxisTitleOperation.preserveSeries, true);

	const formattedAxisChart = {
		...structuredClone(originalChart),
		valueAxis: { visible: true, numberFormat: '0.00' },
	};
	const sourceLinkedAxisChart = {
		...structuredClone(originalChart),
		valueAxis: { visible: true, numberFormat: '' },
	};
	const [resetAxisFormatOperation] = buildNativeExcelEditPlan(
		[sheet('Home', {}, [], { charts: [formattedAxisChart] })],
		[sheet('Home', {}, [], { charts: [sourceLinkedAxisChart] })],
	).operations;
	assert.equal(resetAxisFormatOperation.kind, 'updateChart');
	assert.equal(resetAxisFormatOperation.chart.valueAxis.numberFormat, '');
	assert.equal(resetAxisFormatOperation.preserveSeries, true);

  const importedSeriesChartWithSource = chart('chart:home:ImportedChart', 'ImportedChart', {
    series: [{
      id: 'series:imported',
      nameRange: 'B1',
      categoryRange: 'A2:A5',
      valuesRange: 'B2:B5',
      chartType: 51,
      axisGroup: 'primary',
    }],
  });
  const { sourceRangeRef: _discardedSourceRange, ...importedSeriesChart } = importedSeriesChartWithSource;
  const importedTitleOnlyChart = structuredClone(importedSeriesChart);
  importedTitleOnlyChart.title = { visible: true, text: 'Imported title only' };
  const [importedTitleOperation] = buildNativeExcelEditPlan(
    [sheet('Home', {}, [], { charts: [importedSeriesChart] })],
    [sheet('Home', {}, [], { charts: [importedTitleOnlyChart] })],
  ).operations;
  assert.equal('sourceRangeRef' in importedTitleOperation.chart, false);
  assert.equal(importedTitleOperation.preserveSeries, true);
  assert.equal(importedTitleOperation.preserveAnchor, true);

  const plotByOnlyChart = structuredClone(originalChart);
  plotByOnlyChart.plotBy = 'rows';
  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], { charts: [originalChart] })],
      [sheet('Home', {}, [], { charts: [plotByOnlyChart] })],
    ).operations,
    [{
      kind: 'updateChart',
      sheetName: 'Home',
      name: 'SalesChart',
      chart: plotByOnlyChart,
      preserveAnchor: true,
    }],
    'a plotBy change must rebuild the source-derived series instead of preserving them',
  );

	const chartTypeChanged = { ...structuredClone(originalChart), chartType: 65 };
	const [chartTypeOperation] = buildNativeExcelEditPlan(
		[sheet('Home', {}, [], { charts: [originalChart] })],
		[sheet('Home', {}, [], { charts: [chartTypeChanged] })],
	).operations;
	assert.equal(chartTypeOperation.kind, 'updateChart');
	assert.equal(chartTypeOperation.preserveAnchor, true);
	assert.equal(chartTypeOperation.preserveSeries, undefined);

	const styleChanged = { ...structuredClone(originalChart), style: 12 };
	const [chartStyleOperation] = buildNativeExcelEditPlan(
		[sheet('Home', {}, [], { charts: [originalChart] })],
		[sheet('Home', {}, [], { charts: [styleChanged] })],
	).operations;
	assert.equal(chartStyleOperation.kind, 'updateChart');
	assert.equal(chartStyleOperation.preserveAnchor, true);
	assert.equal(chartStyleOperation.preserveSeries, true);
	assert.equal(chartStyleOperation.allowSeriesFormattingChange, true);
	assert.equal(chartStyleOperation.chart.style, 12);

	const presentationOverrides = {
		...structuredClone(originalChart),
		style: 12,
		gapWidth: 220,
		overlap: 25,
	};
	const resetPresentationChart = {
		...structuredClone(originalChart),
		style: 2,
		gapWidth: 150,
		overlap: 0,
	};
	const [resetPresentationOperation] = buildNativeExcelEditPlan(
		[sheet('Home', {}, [], { charts: [presentationOverrides] })],
		[sheet('Home', {}, [], { charts: [resetPresentationChart] })],
	).operations;
	assert.equal(resetPresentationOperation.kind, 'updateChart');
	assert.equal(resetPresentationOperation.chart.style, 2);
	assert.equal(resetPresentationOperation.chart.gapWidth, 150);
	assert.equal(resetPresentationOperation.chart.overlap, 0);
	assert.equal(resetPresentationOperation.preserveSeries, true);
	assert.equal(resetPresentationOperation.allowSeriesFormattingChange, true);
  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], { charts: [originalChart] })],
      [sheet('Home', {}, [], { charts: [] })],
    ).operations,
    [{ kind: 'deleteChart', sheetName: 'Home', name: 'SalesChart' }],
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {}, [], { tables: disjointTables, charts: [originalChart] })],
      [sheet('Home', {}, [], { tables: disjointTables, charts: [originalChart] })],
    ),
    { operations: [], unsupportedChanges: [] },
    'unchanged workbook objects must not become worksheet-feature changes',
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', {})],
      [sheet('Home', {}, [], {
        tables: [
          table('duplicate-id', 'TableOne', 'A1:C5'),
          table('duplicate-id', 'TableTwo', 'A20:C30'),
        ],
      })],
    ).unsupportedChanges,
    ['Home:tables'],
    'duplicate stable ids must be refused instead of dropping an object',
  );

  console.log('Native edit diff regression tests passed.');
} finally {
  await rm(buildDirectory, { recursive: true, force: true });
}
