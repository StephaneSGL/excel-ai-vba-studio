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
    ).unsupportedChanges,
    ['Home:column-dimensions'],
  );

  assert.deepEqual(
    buildNativeExcelEditPlan(
      [sheet('Home', { '0:0': { text: 'A' } })],
      [{
        ...sheet('Home', { '0:0': { text: 'A' } }),
        rows: { len: 1, 0: { height: 36, cells: { 0: { text: 'A' } } } },
      }],
    ).unsupportedChanges,
    ['Home:row-dimensions'],
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
    ['Home:worksheet-features'],
  );

  console.log('Native edit diff regression tests passed.');
} finally {
  await rm(buildDirectory, { recursive: true, force: true });
}
