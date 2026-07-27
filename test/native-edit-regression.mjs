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
  } = await import(
    `${pathToFileURL(outputPath).href}?cache=${Date.now()}`
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

  console.log('Native edit diff regression tests passed.');
} finally {
  await rm(buildDirectory, { recursive: true, force: true });
}
