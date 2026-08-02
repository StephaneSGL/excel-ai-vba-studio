import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const root = resolve(import.meta.dirname, '..');
const buildDirectory = await mkdtemp(join(tmpdir(), 'excel-chart-catalog-test-'));
const outputPath = join(buildDirectory, 'chart-catalog.mjs');

try {
  await build({
    entryPoints: [join(root, 'src', 'common', 'excelWorkbookObjects.ts')],
    outfile: outputPath,
    bundle: true,
    format: 'esm',
    platform: 'node',
    logLevel: 'silent',
    tsconfig: join(root, 'tsconfig.json'),
  });
  const {
    EXCEL_CHART_TYPES,
    buildExcelTableStyleCatalog,
    canonicalChartTypeForSeries,
    chartSeriesSupportedDataLabelPositions,
  } = await import(
    `${pathToFileURL(outputPath).href}?cache=${Date.now()}`
  );

  assert.equal(EXCEL_CHART_TYPES.length, 103);
  assert.equal(new Set(EXCEL_CHART_TYPES.map(option => option.id)).size, 103);
  assert.equal(
    new Set(EXCEL_CHART_TYPES.map(option => option.constant)).size,
    103,
  );
  assert.deepEqual(
    Object.fromEntries(
      EXCEL_CHART_TYPES
        .filter(option => [
          'xlColumnClustered',
          'xlLine',
          'xlCombo',
          'xlXYScatter',
          'xlRegionMap',
          'xlSuggestedChart',
        ].includes(option.constant))
        .map(option => [option.constant, option.value]),
    ),
    {
      xlColumnClustered: 51,
      xlLine: 4,
      xlCombo: -4152,
      xlXYScatter: -4169,
      xlRegionMap: 140,
      xlSuggestedChart: -2,
    },
  );

  const tableStyles = buildExcelTableStyleCatalog();
  assert.equal(tableStyles.length, 60);
  assert.equal(new Set(tableStyles).size, 60);
  assert.ok(tableStyles.includes('TableStyleLight21'));
  assert.ok(tableStyles.includes('TableStyleMedium28'));
  assert.ok(tableStyles.includes('TableStyleDark11'));

  assert.deepEqual(
    chartSeriesSupportedDataLabelPositions(51),
    ['center', 'insideBase', 'insideEnd', 'outsideEnd'],
  );
  assert.deepEqual(
    chartSeriesSupportedDataLabelPositions(52),
    ['center', 'insideBase', 'insideEnd'],
  );
  assert.deepEqual(
    chartSeriesSupportedDataLabelPositions(4),
    ['above', 'below', 'center', 'left', 'right'],
  );
  assert.deepEqual(
    chartSeriesSupportedDataLabelPositions(5),
    ['bestFit', 'center', 'insideEnd', 'outsideEnd'],
  );
  assert.deepEqual(chartSeriesSupportedDataLabelPositions(-4120), []);
  assert.deepEqual(chartSeriesSupportedDataLabelPositions(-4100), []);
	assert.equal(canonicalChartTypeForSeries(51, []), 51);
	assert.equal(canonicalChartTypeForSeries(51, [4]), 4);
	assert.equal(canonicalChartTypeForSeries(-4152, [51, 51]), 51);
	assert.equal(canonicalChartTypeForSeries(51, [51, 4]), -4152);

  console.log('Complete chart and table-style catalogs validated.');
} finally {
  await rm(buildDirectory, { recursive: true, force: true });
}
