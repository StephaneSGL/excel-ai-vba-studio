import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const root = resolve(import.meta.dirname, '..');
const buildDirectory = await mkdtemp(join(tmpdir(), 'excel-chart-model-test-'));
const outputPath = join(buildDirectory, 'chart-designer-model.mjs');

try {
  await build({
    entryPoints: [
      join(root, 'src', 'react', 'view', 'excel', 'chart-designer-model.ts'),
    ],
    outfile: outputPath,
    bundle: true,
    format: 'esm',
    platform: 'node',
    logLevel: 'silent',
    tsconfig: join(root, 'tsconfig.json'),
  });
  const model = await import(
    `${pathToFileURL(outputPath).href}?cache=${Date.now()}`
  );

  assert.equal(
    model.CHART_TYPE_GROUPS.flatMap(group => group.options).length,
    103,
  );
  const draft = model.createChartDraft([], 'A1:C12', 51);
  draft.name = 'Sales chart';
  draft.title = { visible: true, text: 'Sales by month' };
  draft.series = [{
    ...model.createSeriesDraft(),
    id: 'sales-series',
    name: 'Sales',
    categoryRange: 'A2:A12',
    valuesRange: 'B2:B12',
    chartType: 4,
    axisGroup: 'secondary',
    color: '#4472c4',
    lineColor: '#4472c4',
    markerStyle: 'circle',
    markerSize: 7,
    smooth: true,
    dataLabels: { showValue: true, position: 'above' },
  }];
  draft.valueAxis = {
    visible: true,
    minimumScale: 0,
    maximumScale: 100,
    majorUnit: 10,
    majorGridlines: true,
  };
  const valid = model.validateAndNormalizeChart(draft, []);
  assert.deepEqual(valid.errors, []);
  assert.equal(valid.chart.series[0].axisGroup, 'secondary');
  assert.equal(valid.chart.valueAxis.maximumScale, 100);
	assert.equal(valid.chart.valueAxis.numberFormat, '', 'an unspecified UI format must serialize as source-linked');
	assert.equal(valid.chart.alternativeText, '', 'the UI model must emit an explicit empty alternative text');

	const describedDraft = model.createChartDraft([], 'A1:B12', 51);
	describedDraft.alternativeText = '  Accessible sales chart  ';
	assert.equal(
		model.validateAndNormalizeChart(describedDraft, []).chart.alternativeText,
		'Accessible sales chart',
	);
	describedDraft.alternativeText = '   ';
	assert.equal(
		model.validateAndNormalizeChart(describedDraft, []).chart.alternativeText,
		'',
		'clearing alternative text must remain an explicit native write',
	);
	describedDraft.valueAxis.title = '  Revenue  ';
	assert.equal(
		model.validateAndNormalizeChart(describedDraft, []).chart.valueAxis.title,
		'Revenue',
	);
	describedDraft.valueAxis.title = '   ';
	assert.equal(
		model.validateAndNormalizeChart(describedDraft, []).chart.valueAxis.title,
		'',
		'clearing an axis title must remain an explicit native write',
	);

	const importedWithoutPresentationOverrides = model.createChartDraft([], 'A1:B12', 51);
	delete importedWithoutPresentationOverrides.style;
	delete importedWithoutPresentationOverrides.gapWidth;
	delete importedWithoutPresentationOverrides.overlap;
	const preservedImportedPresentation = model.validateAndNormalizeChart(
		importedWithoutPresentationOverrides,
		[],
	).chart;
	assert.equal('style' in preservedImportedPresentation, false);
	assert.equal('gapWidth' in preservedImportedPresentation, false);
	assert.equal('overlap' in preservedImportedPresentation, false);

	importedWithoutPresentationOverrides.style = 2;
	importedWithoutPresentationOverrides.gapWidth = 150;
	importedWithoutPresentationOverrides.overlap = 0;
	const resetPresentation = model.validateAndNormalizeChart(
		importedWithoutPresentationOverrides,
		[],
	).chart;
	assert.equal(resetPresentation.style, 2);
	assert.equal(resetPresentation.gapWidth, 150);
	assert.equal(resetPresentation.overlap, 0);

  const sourceOnlyDraft = model.createChartDraft([], 'A1:C12', 51);
  const sourceOnly = model.validateAndNormalizeChart(sourceOnlyDraft, []);
  assert.deepEqual(sourceOnly.errors, []);
  assert.equal('series' in sourceOnly.chart, false, 'source-only UI charts must not emit an empty explicit-series replacement');

	const absoluteDraft = model.createChartDraft([], '$A$1:$C$12', 51);
	const absoluteChart = model.validateAndNormalizeChart(absoluteDraft, []);
	assert.deepEqual(absoluteChart.errors, []);
	assert.equal(absoluteChart.chart.sourceRangeRef, '$A$1:$C$12');

	const homogeneousCombo = model.createChartDraft([], '', -4152);
	homogeneousCombo.series = [
		{ id: 'combo:one', chartType: 51, valuesRange: 'B2:B12' },
		{ id: 'combo:two', chartType: 51, valuesRange: 'C2:C12' },
	];
	const normalizedHomogeneousCombo = model.validateAndNormalizeChart(homogeneousCombo, []);
	assert.deepEqual(normalizedHomogeneousCombo.errors, []);
	assert.equal(normalizedHomogeneousCombo.chart.chartType, 51);

	const singleSeriesTypeOverride = model.createChartDraft([], '', 51);
	singleSeriesTypeOverride.series = [
		{ id: 'single-line', chartType: 4, valuesRange: 'B2:B12' },
	];
	const normalizedSingleSeriesType = model.validateAndNormalizeChart(singleSeriesTypeOverride, []);
	assert.deepEqual(normalizedSingleSeriesType.errors, []);
	assert.equal(normalizedSingleSeriesType.chart.chartType, 4);

	const inferredCombo = model.createChartDraft([], '', 51);
	inferredCombo.series = [
		{ id: 'inferred:column', valuesRange: 'B2:B12' },
		{ id: 'inferred:line', chartType: 4, valuesRange: 'C2:C12' },
	];
	const normalizedInferredCombo = model.validateAndNormalizeChart(inferredCombo, []);
	assert.deepEqual(normalizedInferredCombo.errors, []);
	assert.equal(normalizedInferredCombo.chart.chartType, -4152);
	assert.deepEqual(normalizedInferredCombo.chart.series.map(series => series.chartType), [51, 4]);

	const invalidBubbleSizes = model.createChartDraft([], '', 51);
	invalidBubbleSizes.series = [{
		id: 'not-bubble',
		valuesRange: 'B2:B12',
		bubbleSizesRange: 'C2:C12',
	}];
	assert.match(
		model.validateAndNormalizeChart(invalidBubbleSizes, []).errors.join('\n'),
		/type de graphique à bulles/,
	);

	const invalidSeriesOptions = model.createChartDraft([], '', 51);
	invalidSeriesOptions.series = [{
		id: 'column-options',
		valuesRange: 'B2:B12',
		smooth: false,
		dataLabels: { showBubbleSize: true, showPercentage: true },
	}];
	const invalidSeriesOptionErrors = model.validateAndNormalizeChart(invalidSeriesOptions, []).errors.join('\n');
	assert.match(invalidSeriesOptionErrors, /lissage.*ligne ou nuage de points/);
	assert.match(invalidSeriesOptionErrors, /taille de bulle.*type bulle/);
	assert.match(invalidSeriesOptionErrors, /pourcentage.*secteur ou anneau/);

	const falseUnsupportedLabels = model.createChartDraft([], '', 51);
	falseUnsupportedLabels.series = [{
		id: 'column-false-labels',
		valuesRange: 'B2:B12',
		dataLabels: { showBubbleSize: false, showPercentage: false },
	}];
	const normalizedFalseLabels = model.validateAndNormalizeChart(falseUnsupportedLabels, []);
	assert.deepEqual(normalizedFalseLabels.errors, []);
	assert.deepEqual(normalizedFalseLabels.chart.series[0].dataLabels, {
		showPercentage: false,
		showBubbleSize: false,
	});

	const disabledLabelsWithPosition = structuredClone(falseUnsupportedLabels);
	disabledLabelsWithPosition.series[0].dataLabels.position = 'outsideEnd';
	assert.match(
		model.validateAndNormalizeChart(disabledLabelsWithPosition, []).errors.join('\n'),
		/position d’étiquette exige au moins une option d’affichage activée/,
	);

	const emptyLabels = model.createChartDraft([], '', 51);
	emptyLabels.series = [{ id: 'empty-labels', valuesRange: 'B2:B12', dataLabels: {} }];
	const normalizedEmptyLabels = model.validateAndNormalizeChart(emptyLabels, []);
	assert.match(
		normalizedEmptyLabels.errors.join('\n'),
		/au moins une option d’affichage explicite/,
	);

	const positionOnlyLabels = model.createChartDraft([], '', 51);
	positionOnlyLabels.series = [{
		id: 'position-only-labels',
		valuesRange: 'B2:B12',
		dataLabels: { position: 'outsideEnd' },
	}];
	assert.match(
		model.validateAndNormalizeChart(positionOnlyLabels, []).errors.join('\n'),
		/au moins une option d’affichage explicite/,
	);

	const validPiePercentage = model.createChartDraft([], '', 5);
	validPiePercentage.series = [{
		id: 'pie-percentage',
		valuesRange: 'B2:B12',
		dataLabels: { showPercentage: true },
	}];
	assert.deepEqual(model.validateAndNormalizeChart(validPiePercentage, []).errors, []);

	const validBubbleLabel = model.createChartDraft([], '', 15);
	validBubbleLabel.series = [{
		id: 'bubble-label',
		xValuesRange: 'A2:A12',
		valuesRange: 'B2:B12',
		bubbleSizesRange: 'C2:C12',
		dataLabels: { showBubbleSize: true },
	}];
	assert.deepEqual(model.validateAndNormalizeChart(validBubbleLabel, []).errors, []);

	const validColumnLabelPosition = model.createChartDraft([], '', 51);
	validColumnLabelPosition.series = [{
		id: 'column-label-position',
		valuesRange: 'B2:B12',
		dataLabels: { showValue: true, position: 'outsideEnd' },
	}];
	assert.deepEqual(model.validateAndNormalizeChart(validColumnLabelPosition, []).errors, []);

	const customNumberFormat = model.createChartDraft([], 'A1:B12', 51);
	customNumberFormat.valueAxis = { visible: true, numberFormat: '  0.00  ' };
	assert.equal(
		model.validateAndNormalizeChart(customNumberFormat, []).chart.valueAxis.numberFormat,
		'0.00',
	);
	customNumberFormat.valueAxis.numberFormat = '   ';
	assert.equal(
		model.validateAndNormalizeChart(customNumberFormat, []).chart.valueAxis.numberFormat,
		'',
		'clearing the axis format must request source-linked formatting',
	);

	const invalidColumnLabelPosition = structuredClone(validColumnLabelPosition);
	invalidColumnLabelPosition.series[0].dataLabels.position = 'above';
	assert.match(
		model.validateAndNormalizeChart(invalidColumnLabelPosition, []).errors.join('\n'),
		/position d’étiquette above.*pas prise en charge/,
	);

	const invalidStackedLabelPosition = structuredClone(validColumnLabelPosition);
	invalidStackedLabelPosition.chartType = 52;
	assert.match(
		model.validateAndNormalizeChart(invalidStackedLabelPosition, []).errors.join('\n'),
		/position d’étiquette outsideEnd.*pas prise en charge/,
	);

	const invalidDoughnutLabelPosition = model.createChartDraft([], '', -4120);
	invalidDoughnutLabelPosition.series = [{
		id: 'doughnut-label-position',
		valuesRange: 'B2:B12',
		dataLabels: { showPercentage: true, position: 'center' },
	}];
	assert.match(
		model.validateAndNormalizeChart(invalidDoughnutLabelPosition, []).errors.join('\n'),
		/position d’étiquette center.*pas prise en charge/,
	);

	const mixedBubbleCombo = model.createChartDraft([], '', -4152);
	mixedBubbleCombo.series = [
		{ id: 'mixed-column', chartType: 51, valuesRange: 'B2:B12' },
		{
			id: 'mixed-bubble',
			chartType: 15,
			xValuesRange: 'A2:A12',
			valuesRange: 'C2:C12',
			bubbleSizesRange: 'D2:D12',
		},
	];
	assert.match(
		model.validateAndNormalizeChart(mixedBubbleCombo, []).errors.join('\n'),
		/mélanger des séries à bulles.*autre type/,
	);

	const newCustomLegend = model.createChartDraft([], 'A1:B12', 51);
	newCustomLegend.legend = { visible: true, position: 'custom' };
	assert.match(
		model.validateAndNormalizeChart(newCustomLegend, []).errors.join('\n'),
		/disposition manuelle déjà créée dans Excel/,
	);
	assert.deepEqual(
		model.validateAndNormalizeChart(newCustomLegend, [structuredClone(newCustomLegend)]).errors,
		[],
		'an imported custom legend may be preserved on update',
	);

	const effectiveComboOptions = model.createChartDraft([], '', -4152);
	effectiveComboOptions.series = [
		{ id: 'combo-line', chartType: 4, valuesRange: 'B2:B12', smooth: true },
		{ id: 'combo-column', chartType: 51, valuesRange: 'C2:C12', smooth: true },
	];
	assert.match(
		model.validateAndNormalizeChart(effectiveComboOptions, []).errors.join('\n'),
		/Série 2.*lissage/,
		'combo validation must use each series effective chart type',
	);

	const mixedAxes = model.createChartDraft([], '', -4152);
	mixedAxes.series = [
		{ id: 'primary-column', chartType: 51, axisGroup: 'primary', valuesRange: 'B2:B12' },
		{ id: 'secondary-scatter', chartType: 74, axisGroup: 'secondary', xValuesRange: 'A2:A12', valuesRange: 'C2:C12' },
	];
	mixedAxes.categoryAxis = { visible: true, minimumScale: 0 };
	mixedAxes.secondaryCategoryAxis = { visible: true, minimumScale: 0 };
	assert.match(
		model.validateAndNormalizeChart(mixedAxes, []).errors.join('\n'),
		/Axe des catégories.*bornes numériques/,
	);
	delete mixedAxes.categoryAxis.minimumScale;
	assert.deepEqual(model.validateAndNormalizeChart(mixedAxes, []).errors, []);

  for (const chartType of [-4100, 83]) {
    const threeDimensionalDraft = model.createChartDraft([], 'A1:C12', chartType);
    assert.equal('secondaryCategoryAxis' in threeDimensionalDraft, false);
    assert.equal('secondaryValueAxis' in threeDimensionalDraft, false);
    const threeDimensional = model.validateAndNormalizeChart(threeDimensionalDraft, []);
    assert.deepEqual(threeDimensional.errors, []);
    assert.equal('secondaryCategoryAxis' in threeDimensional.chart, false);
    assert.equal('secondaryValueAxis' in threeDimensional.chart, false);
  }

  const pieDraft = model.createChartDraft([], 'A1:B12', 5);
  assert.equal('categoryAxis' in pieDraft, false);
  assert.equal('valueAxis' in pieDraft, false);

  const duplicate = model.validateAndNormalizeChart(
    draft,
    [{ ...draft, id: 'another-chart' }],
  );
  assert.match(duplicate.errors.join('\n'), /porte déjà ce nom/);

  const invalidExternal = structuredClone(draft);
  invalidExternal.sourceRangeRef = '[other.xlsx]Sheet1!A1:C12';
  assert.match(
    model.validateAndNormalizeChart(invalidExternal, []).errors.join('\n'),
    /plage Excel locale valide/,
  );

  const invalidOtherSheet = structuredClone(draft);
  invalidOtherSheet.sourceRangeRef = "'Other sheet'!A1:C12";
  assert.match(
    model.validateAndNormalizeChart(invalidOtherSheet, []).errors.join('\n'),
    /plage Excel locale valide/,
  );

  const invalidWholeColumns = structuredClone(draft);
  invalidWholeColumns.sourceRangeRef = 'A:C';
  assert.match(
    model.validateAndNormalizeChart(invalidWholeColumns, []).errors.join('\n'),
    /plage Excel locale valide/,
  );

  const invalidAxis = structuredClone(draft);
  invalidAxis.valueAxis = {
    visible: true,
    logarithmic: true,
    minimumScale: 0,
    maximumScale: 100,
  };
  assert.match(
    model.validateAndNormalizeChart(invalidAxis, []).errors.join('\n'),
    /bornes positives/,
  );

  const undersizedChart = structuredClone(draft);
  undersizedChart.anchor.width = 19.5;
  assert.match(
    model.validateAndNormalizeChart(undersizedChart, []).errors.join('\n'),
    /supérieures ou égales à 20 points/,
  );

  const fractionalGap = structuredClone(draft);
  fractionalGap.gapWidth = 120.5;
  assert.match(
    model.validateAndNormalizeChart(fractionalGap, []).errors.join('\n'),
    /largeur d’intervalle doit être un entier/,
  );

  const fractionalOverlap = structuredClone(draft);
  fractionalOverlap.overlap = -4.5;
  assert.match(
    model.validateAndNormalizeChart(fractionalOverlap, []).errors.join('\n'),
    /chevauchement doit être un entier/i,
  );

  const fractionalMarker = structuredClone(draft);
  fractionalMarker.series[0].markerSize = 7.5;
  assert.match(
    model.validateAndNormalizeChart(fractionalMarker, []).errors.join('\n'),
    /taille de marqueur doit être un entier/,
  );

  const oversizedSeriesName = structuredClone(draft);
  oversizedSeriesName.series[0].name = 'S'.repeat(256);
  assert.match(
    model.validateAndNormalizeChart(oversizedSeriesName, []).errors.join('\n'),
    /nom libre ne peut pas dépasser 255 caractères/,
  );

  const oversizedAxisText = structuredClone(draft);
  oversizedAxisText.valueAxis.title = 'T'.repeat(1001);
  oversizedAxisText.valueAxis.numberFormat = '0'.repeat(256);
  const oversizedAxisErrors = model.validateAndNormalizeChart(oversizedAxisText, []).errors.join('\n');
  assert.match(oversizedAxisErrors, /titre ne peut pas dépasser 1 000 caractères/);
  assert.match(oversizedAxisErrors, /format numérique ne peut pas dépasser 255 caractères/);

  const invalidColumn = structuredClone(draft);
  invalidColumn.sourceRangeRef = 'XFE1:XFE2';
  assert.match(
    model.validateAndNormalizeChart(invalidColumn, []).errors.join('\n'),
    /plage Excel locale valide/,
  );

  const oversizedRange = structuredClone(draft);
  oversizedRange.sourceRangeRef = 'A1:B600000';
  assert.match(
    model.validateAndNormalizeChart(oversizedRange, []).errors.join('\n'),
    /limite de 1[\s\u202f]?000[\s\u202f]?000 cellules/,
  );

  const ambiguousSeries = structuredClone(draft);
  ambiguousSeries.series[0].xValuesRange = 'A2:A12';
  assert.match(
    model.validateAndNormalizeChart(ambiguousSeries, []).errors.join('\n'),
    /catégories et valeurs X/,
  );

  const multiCellSeriesName = structuredClone(draft);
  multiCellSeriesName.series[0].name = '';
  multiCellSeriesName.series[0].nameRange = 'B1:B2';
  assert.match(
    model.validateAndNormalizeChart(multiCellSeriesName, []).errors.join('\n'),
    /doit désigner une seule cellule/,
  );

  const suggestedSeries = structuredClone(draft);
  suggestedSeries.series[0].chartType = -2;
  assert.match(
    model.validateAndNormalizeChart(suggestedSeries, []).errors.join('\n'),
    /type de graphique local autorisé/,
  );

  const regionMap = model.createChartDraft([], 'A1:B12', 140);
  assert.match(
    model.validateAndNormalizeChart(regionMap, []).errors.join('\n'),
    /Bing Maps/,
  );

  const formulaSeriesName = structuredClone(draft);
  formulaSeriesName.series[0].name = "='[other.xlsx]Sheet1'!A1";
  assert.match(
    model.validateAndNormalizeChart(formulaSeriesName, []).errors.join('\n'),
    /texte littéral, pas une formule/,
  );
  for (const maliciousName of ['+SUM(A1:A2)', '-1+2', '@A1']) {
    const formulaLikeSeriesName = structuredClone(draft);
    formulaLikeSeriesName.series[0].name = maliciousName;
    assert.match(
      model.validateAndNormalizeChart(formulaLikeSeriesName, []).errors.join('\n'),
      /texte littéral, pas une formule/,
    );
  }

  const immutable = structuredClone(draft);
  const normalized = model.validateAndNormalizeChart(immutable, []).chart;
  normalized.title.text = 'Changed result';
  assert.equal(immutable.title.text, 'Sales by month');

  console.log('Chart designer model validation passed.');
} finally {
  await rm(buildDirectory, { recursive: true, force: true });
}
