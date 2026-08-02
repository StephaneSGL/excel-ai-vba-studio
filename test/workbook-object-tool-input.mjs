import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const root = resolve(import.meta.dirname, '..');
const buildDirectory = await mkdtemp(join(tmpdir(), 'excel-object-tool-test-'));
const outputPath = join(buildDirectory, 'workbook-object-tool.mjs');

try {
  await build({
    entryPoints: [
      join(root, 'src', 'excelAiVbaStudio', 'workbookObjectTool.ts'),
    ],
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
          () => ({ contents: 'export class CancellationError extends Error {}' }),
        );
      },
    }],
  });

  const toolModule = await import(
    `${pathToFileURL(outputPath).href}?cache=${Date.now()}`
  );
  const parseWorkbookObjectToolInput = value => toolModule.parseWorkbookObjectToolInput({
    workbookPath: 'C:\\work\\contacts.xlsx',
    ...value,
  });

  assert.throws(
    () => toolModule.parseWorkbookObjectToolInput({ operations: [] }),
    /workbookPath est obligatoire/,
  );
  assert.throws(
    () => toolModule.parseWorkbookObjectToolInput({
      workbookPath: 'contacts.xlsx',
      operations: [{
        kind: 'deleteWorksheetChart',
        sheetName: 'People',
        name: 'OldChart',
      }],
    }),
    /chemin absolu explicite/,
  );

  const threeTables = parseWorkbookObjectToolInput({
    workbookPath: 'C:\\work\\contacts.xlsx',
    operations: [
      ['PeopleTop', 'A1:C5'],
      ['PeopleMiddle', 'A20:C30'],
      ['PeopleBottom', 'A40:C50'],
    ].map(([name, rangeRef]) => ({
      kind: 'createWorksheetTable',
      sheetName: 'People',
      table: { name, rangeRef },
    })),
  });
  assert.deepEqual(
    threeTables.operations.map(operation => operation.table.rangeRef),
    ['A1:C5', 'A20:C30', 'A40:C50'],
  );
  assert.ok(
    threeTables.operations.every(
      operation => operation.table.style.name === 'TableStyleMedium2'
        && operation.table.style.showRowStripes === true,
    ),
  );
	assert.throws(
		() => parseWorkbookObjectToolInput({
			operations: [{
				kind: 'createWorksheetTable',
				sheetName: 'People',
				table: { name: 'NoHeaders', rangeRef: 'A1:C5', headerRow: false },
			}],
		}),
		/headerRow=false est refusé à la création/,
	);
	assert.throws(
		() => parseWorkbookObjectToolInput({
			operations: [{
				kind: 'createWorksheetTable',
				sheetName: 'People',
				table: { name: 'NoNativeTotals', rangeRef: 'A1:C5', totalsRow: true },
			}],
		}),
		/totalsRow=true est refusé à la création/,
	);
	assert.throws(
		() => parseWorkbookObjectToolInput({
			operations: [{
				kind: 'createWorksheetTable',
				sheetName: 'People',
				table: { name: 'TooShortTotals', rangeRef: 'A1:C2', totalsRow: true },
			}],
		}),
		/pas assez de lignes.*totaux/,
	);
	assert.throws(
		() => parseWorkbookObjectToolInput({
			operations: [{
				kind: 'createWorksheetTable',
				sheetName: 'People',
				table: { name: 'TooShortData', rangeRef: 'A1:C1' },
			}],
		}),
		/pas assez de lignes.*données/,
	);

  const chart = parseWorkbookObjectToolInput({
    workbookPath: 'C:\\work\\contacts.xlsx',
    operations: [{
      kind: 'createWorksheetChart',
      sheetName: 'People',
      chart: {
        name: 'Age by city',
        chartType: -4152,
        plotBy: 'columns',
        anchor: { left: 400, top: 20, width: 640, height: 360 },
        title: { visible: true, text: 'Age by city' },
        legend: { visible: true, position: 'bottom' },
        valueAxis: {
          title: 'Age',
		  numberFormat: '',
          minimumScale: 0,
          maximumScale: 100,
          majorGridlines: true,
        },
        series: [{
          name: 'Average age',
          categoryRange: 'A2:A50',
          valuesRange: 'C2:C50',
          chartType: 4,
          axisGroup: 'secondary',
          color: '#4472C4',
          lineColor: '#4472C4',
          lineWidth: 2,
          dashStyle: 'solid',
          markerStyle: 'circle',
          markerSize: 7,
          smooth: true,
          visible: true,
          dataLabels: { showValue: true, position: 'above' },
        }, {
          id: 'series:volume',
          name: 'Volume',
          categoryRange: 'A2:A50',
          valuesRange: 'D2:D50',
          chartType: 51,
          axisGroup: 'primary',
        }],
        style: 10,
        alternativeText: 'Average age by city',
      },
    }],
  }).operations[0].chart;
  assert.equal('sourceRangeRef' in chart, false);
  assert.equal(chart.series[0].axisGroup, 'secondary');
	assert.equal(chart.series[1].chartType, 51);
  assert.equal(chart.series[0].color, '#4472c4');
	assert.equal(chart.valueAxis.maximumScale, 100);
	assert.equal(chart.valueAxis.numberFormat, '', 'the tool must preserve an explicit source-linked reset');
	assert.throws(
		() => parseWorkbookObjectToolInput({
			operations: [{
				kind: 'createWorksheetChart',
				sheetName: 'People',
				chart: {
					name: 'Invented custom legend',
					chartType: 51,
					sourceRangeRef: 'A1:B10',
					anchor: { left: 0, top: 0, width: 320, height: 200 },
					legend: { visible: true, position: 'custom' },
				},
			}],
		}),
		/peut seulement préserver une disposition manuelle existante/,
	);

	const parseSeriesOptionChart = (chartType, series) => parseWorkbookObjectToolInput({
		operations: [{
			kind: 'createWorksheetChart',
			sheetName: 'People',
			chart: {
				name: `Options ${chartType}`,
				chartType,
				anchor: { left: 0, top: 0, width: 320, height: 200 },
				series,
			},
		}],
	});
	assert.throws(
		() => parseSeriesOptionChart(51, [{ id: 'column', valuesRange: 'B2:B10', smooth: false }]),
		/smooth exige une série ligne ou nuage de points/,
		'even smooth:false must not cross the tool boundary for unsupported series',
	);
	assert.doesNotThrow(() => parseSeriesOptionChart(51, [{
		id: 'column-false-labels',
		valuesRange: 'B2:B10',
		dataLabels: { showBubbleSize: false, showPercentage: false },
	}]));
	assert.throws(
		() => parseSeriesOptionChart(51, [{
			id: 'disabled-label-position',
			valuesRange: 'B2:B10',
			dataLabels: { showValue: false, position: 'outsideEnd' },
		}]),
		/position exige au moins une option d’affichage activée/,
	);
	assert.throws(
		() => parseSeriesOptionChart(51, [{
			id: 'column-empty-labels', valuesRange: 'B2:B10', dataLabels: {},
		}]),
		/doit définir explicitement au moins une option d’affichage/,
	);
	assert.throws(
		() => parseSeriesOptionChart(51, [{
			id: 'column-bad-labels',
			valuesRange: 'B2:B10',
			dataLabels: { showBubbleSize: true },
		}]),
		/showBubbleSize exige une série de type bulle/,
	);
	assert.throws(
		() => parseSeriesOptionChart(51, [{
			id: 'column-bad-percentage',
			valuesRange: 'B2:B10',
			dataLabels: { showPercentage: true },
		}]),
		/showPercentage exige une série de type secteur ou anneau/,
	);
	assert.doesNotThrow(() => parseSeriesOptionChart(5, [{
		id: 'pie-label', valuesRange: 'B2:B10', dataLabels: { showPercentage: true },
	}]));
	assert.doesNotThrow(() => parseSeriesOptionChart(15, [{
		id: 'bubble-label', valuesRange: 'B2:B10', dataLabels: { showBubbleSize: true },
	}]));
	assert.throws(
		() => parseSeriesOptionChart(51, [{
			id: 'position-only', valuesRange: 'B2:B10', dataLabels: { position: 'outsideEnd' },
		}]),
		/doit définir explicitement au moins une option d’affichage/,
	);
	assert.doesNotThrow(() => parseSeriesOptionChart(51, [{
		id: 'column-position', valuesRange: 'B2:B10', dataLabels: { showValue: true, position: 'outsideEnd' },
	}]));
	assert.throws(
		() => parseSeriesOptionChart(51, [{
			id: 'column-invalid-position', valuesRange: 'B2:B10', dataLabels: { showValue: true, position: 'above' },
		}]),
		/dataLabels.position n’est pas prise en charge/,
	);
	assert.throws(
		() => parseSeriesOptionChart(52, [{
			id: 'stacked-invalid-position', valuesRange: 'B2:B10', dataLabels: { showValue: true, position: 'outsideEnd' },
		}]),
		/dataLabels.position n’est pas prise en charge/,
	);
	assert.throws(
		() => parseSeriesOptionChart(-4120, [{
			id: 'doughnut-invalid-position', valuesRange: 'B2:B10', dataLabels: { showValue: true, position: 'center' },
		}]),
		/dataLabels.position n’est pas prise en charge/,
	);
	assert.throws(
		() => parseSeriesOptionChart(-4152, [
			{ id: 'column', chartType: 51, valuesRange: 'B2:B10' },
			{
				id: 'bubble',
				chartType: 15,
				xValuesRange: 'A2:A10',
				valuesRange: 'C2:C10',
				bubbleSizesRange: 'D2:D10',
			},
		]),
		/mélanger des séries à bulles et non-bulles/,
	);
	assert.throws(
		() => parseSeriesOptionChart(-4152, [
			{ id: 'line', chartType: 4, valuesRange: 'B2:B10', smooth: true },
			{ id: 'column', chartType: 51, valuesRange: 'C2:C10', smooth: true },
		]),
		/series\[1\]\.smooth/,
		'combo tool validation must use the concrete type of each series',
	);

  assert.equal(
    parseWorkbookObjectToolInput({
      operations: [{
        kind: 'createWorksheetTable',
        sheetName: 'People',
        table: { name: 'Équipe_2026', rangeRef: 'A1:C5' },
      }],
    }).operations[0].table.name,
    'Équipe_2026',
  );
  assert.equal(
    parseWorkbookObjectToolInput({
      operations: [{
        kind: 'createWorksheetTable',
        sheetName: 'People',
        table: { name: 'E\u0301quipe_2026', rangeRef: 'A1:C5' },
      }],
    }).operations[0].table.name,
    'Équipe_2026',
    'the tool boundary must normalize Unicode table names to NFC',
  );
  assert.throws(
    () => parseWorkbookObjectToolInput({
      operations: [{
        kind: 'createWorksheetTable',
        sheetName: 'People',
        table: { name: 'XFD1048576', rangeRef: 'A1:C5' },
      }],
    }),
    /nom de tableau Excel valide/,
  );
  assert.doesNotThrow(() => parseWorkbookObjectToolInput({
    operations: [{
      kind: 'createWorksheetTable',
      sheetName: 'People',
      table: { name: 'XFE1048576', rangeRef: 'A1:C5' },
    }],
  }));
  assert.throws(
    () => parseWorkbookObjectToolInput({
      operations: [{
        kind: 'createWorksheetTable',
        sheetName: 'People',
        table: { name: 'Équipe_2026', rangeRef: 'A1:C5' },
      }, {
        kind: 'createWorksheetTable',
        sheetName: 'Archive',
        table: { name: 'E\u0301QUIPE_2026', rangeRef: 'A1:C5' },
      }],
    }),
    /demandé plusieurs fois dans le classeur/,
    'canonical Unicode duplicates must be rejected workbook-wide',
  );

  const sourceOnlyChart = parseWorkbookObjectToolInput({
    workbookPath: 'C:\\work\\contacts.xlsx',
    operations: [{
      kind: 'createWorksheetChart',
      sheetName: 'People',
      chart: {
        name: 'Source only',
        chartType: 51,
        sourceRangeRef: 'A1:B10',
        anchor: { left: 0, top: 0, width: 320, height: 200 },
      },
    }],
  }).operations[0].chart;
  assert.equal('series' in sourceOnlyChart, false);

	const inferredCombo = parseWorkbookObjectToolInput({
		operations: [{
			kind: 'createWorksheetChart',
			sheetName: 'People',
			chart: {
				name: 'Inferred combo',
				chartType: 51,
				anchor: { left: 0, top: 0, width: 320, height: 200 },
				series: [
					{ id: 'series:column', valuesRange: 'B2:B10' },
					{ id: 'series:line', chartType: 4, valuesRange: 'C2:C10' },
				],
			},
		}],
	}).operations[0].chart;
	assert.equal(inferredCombo.chartType, -4152);
	assert.deepEqual(inferredCombo.series.map(series => series.chartType), [51, 4]);

	assert.throws(
		() => parseWorkbookObjectToolInput({
			operations: [{
				kind: 'createWorksheetChart',
				sheetName: 'People',
				chart: {
					name: 'Invalid bubble sizes',
					chartType: 51,
					anchor: { left: 0, top: 0, width: 320, height: 200 },
					series: [{ id: 'series:column', valuesRange: 'B2:B10', bubbleSizesRange: 'C2:C10' }],
				},
			}],
		}),
		/exige une série de type bulle/,
	);
	assert.throws(
		() => parseWorkbookObjectToolInput({
			operations: [{
				kind: 'createWorksheetChart',
				sheetName: 'People',
				chart: {
					name: 'Invalid primary category scale',
					chartType: -4152,
					anchor: { left: 0, top: 0, width: 320, height: 200 },
					categoryAxis: { visible: true, minimumScale: 0 },
					series: [
						{ id: 'series:column', chartType: 51, axisGroup: 'primary', valuesRange: 'B2:B10' },
						{ id: 'series:scatter', chartType: 74, axisGroup: 'secondary', xValuesRange: 'A2:A10', valuesRange: 'C2:C10' },
					],
				},
			}],
		}),
		/categoryAxis ne peut pas définir d’échelle de valeurs/,
	);

  assert.throws(
    () => parseWorkbookObjectToolInput({
      operations: [{
        kind: 'createWorksheetChart',
        sheetName: 'People',
        chart: {
          name: 'External',
          chartType: 51,
          sourceRangeRef: "'[other.xlsx]Sheet1'!A1:B10",
          anchor: { left: 0, top: 0, width: 320, height: 200 },
        },
      }],
    }),
    /plage A1 locale simple/,
  );
  assert.throws(
    () => parseWorkbookObjectToolInput({
      operations: [{
        kind: 'createWorksheetChart',
        sheetName: 'People',
        chart: {
          name: 'Unknown type',
          chartType: 9999,
          anchor: { left: 0, top: 0, width: 320, height: 200 },
        },
      }],
    }),
    /XlChartType publié/,
  );
  assert.throws(
    () => parseWorkbookObjectToolInput({
      operations: [{
        kind: 'createWorksheetChart',
        sheetName: 'People',
        chart: {
          name: 'Offline map',
          chartType: 140,
          sourceRangeRef: 'A1:B10',
          anchor: { left: 0, top: 0, width: 320, height: 200 },
        },
      }],
    }),
    /création locale sans accès réseau/,
  );
  assert.throws(
    () => parseWorkbookObjectToolInput({
      operations: [{
        kind: 'createWorksheetChart',
        sheetName: 'People',
        chart: {
          name: 'Formula series name',
          chartType: 51,
          anchor: { left: 0, top: 0, width: 320, height: 200 },
          series: [{
            name: "='[other.xlsx]Sheet1'!A1",
            valuesRange: 'B2:B10',
          }],
        },
      }],
    }),
    /texte littéral.*formule/,
  );
  for (const maliciousName of ['+SUM(A1:A2)', '-1+2', '@A1']) {
    assert.throws(
      () => parseWorkbookObjectToolInput({
        operations: [{
          kind: 'createWorksheetChart',
          sheetName: 'People',
          chart: {
            name: 'Formula-like series name',
            chartType: 51,
            anchor: { left: 0, top: 0, width: 320, height: 200 },
            series: [{ name: maliciousName, valuesRange: 'B2:B10' }],
          },
        }],
      }),
      /texte littéral.*formule/,
    );
  }
  assert.throws(
    () => parseWorkbookObjectToolInput({
      operations: [
        {
          kind: 'deleteWorksheetTable',
          sheetName: 'People',
          name: 'PeopleTop',
        },
        {
          kind: 'deleteWorksheetTable',
          sheetName: 'People',
          name: 'PeopleTop',
        },
      ],
    }),
    /même objet est ciblé plusieurs fois/,
  );

	assert.throws(
		() => parseWorkbookObjectToolInput({
			operations: [
				{ kind: 'createWorksheetTable', sheetName: 'People', table: { name: 'Top', rangeRef: 'A1:C10' } },
				{ kind: 'createWorksheetTable', sheetName: 'People', table: { name: 'Overlap', rangeRef: 'C10:E20' } },
			],
		}),
		/tableaux Top et Overlap se chevauchent/,
	);

	assert.throws(
		() => parseWorkbookObjectToolInput({
			operations: [{
				kind: 'createWorksheetChart',
				sheetName: 'People',
				chart: {
					name: 'Duplicate series ids',
					chartType: 51,
					anchor: { left: 0, top: 0, width: 320, height: 200 },
					series: [
						{ id: 'series:duplicate', valuesRange: 'B2:B10' },
						{ id: 'series:duplicate', valuesRange: 'C2:C10' },
					],
				},
			}],
		}),
		/plusieurs fois l’identifiant/,
	);

	const homogeneousCombo = parseWorkbookObjectToolInput({
		operations: [{
			kind: 'createWorksheetChart',
			sheetName: 'People',
			chart: {
				name: 'Homogeneous combo',
				chartType: -4152,
				anchor: { left: 0, top: 0, width: 320, height: 200 },
				series: [
					{ id: 'series:one', chartType: 51, valuesRange: 'B2:B10' },
					{ id: 'series:two', chartType: 51, valuesRange: 'C2:C10' },
				],
			},
		}],
	}).operations[0].chart;
	assert.equal(homogeneousCombo.chartType, 51);

	const singleSeriesOverride = parseSeriesOptionChart(51, [{
		id: 'series:single-line', chartType: 4, valuesRange: 'B2:B10',
	}]).operations[0].chart;
	assert.equal(singleSeriesOverride.chartType, 4);

	assert.throws(
		() => parseWorkbookObjectToolInput({
			operations: [{
				kind: 'createWorksheetChart',
				sheetName: 'People',
				chart: {
					name: 'Invalid mixed secondary axis',
					chartType: -4152,
					anchor: { left: 0, top: 0, width: 320, height: 200 },
					series: [
						{ id: 'series:line', chartType: 4, valuesRange: 'B2:B10' },
						{ id: 'series:3d', chartType: -4100, axisGroup: 'secondary', valuesRange: 'C2:C10' },
					],
				},
			}],
		}),
		/ne peut pas affecter de série à un axe secondaire/,
	);

  assert.throws(
    () => parseWorkbookObjectToolInput({
      operations: [
        {
          kind: 'createWorksheetTable',
          sheetName: 'First',
          table: { name: 'GlobalName', rangeRef: 'A1:C5' },
        },
        {
          kind: 'createWorksheetTable',
          sheetName: 'Second',
          table: { name: 'globalname', rangeRef: 'A20:C30' },
        },
      ],
    }),
    /nom de tableau .* demandé plusieurs fois/i,
  );
  for (const reservedTableName of ['R', 'C']) {
    assert.throws(
      () => parseWorkbookObjectToolInput({
        operations: [{
          kind: 'createWorksheetTable',
          sheetName: 'People',
          table: { name: reservedTableName, rangeRef: 'A1:C5' },
        }],
      }),
      /nom de tableau Excel valide/,
    );
  }
  assert.throws(
    () => parseWorkbookObjectToolInput({
      operations: [
        {
          kind: 'updateWorksheetChart',
          sheetName: 'People',
          name: 'Existing chart',
          chart: {
            name: 'Duplicate chart',
            chartType: 51,
            sourceRangeRef: 'A1:B10',
            anchor: { left: 0, top: 0, width: 320, height: 200 },
          },
        },
        {
          kind: 'createWorksheetChart',
          sheetName: 'People',
          chart: {
            name: 'Duplicate chart',
            chartType: 51,
            sourceRangeRef: 'A1:B10',
            anchor: { left: 340, top: 0, width: 320, height: 200 },
          },
        },
      ],
    }),
    /nom de graphique .* demandé plusieurs fois/i,
  );

  assert.throws(
    () => parseWorkbookObjectToolInput({
      operations: [{
        kind: 'createWorksheetTable',
        sheetName: 'People',
        table: { name: 'TooLarge', rangeRef: 'A1:B600000' },
      }],
    }),
    /limite de .*cellules/,
  );

  assert.throws(
    () => parseWorkbookObjectToolInput({
      operations: [{
        kind: 'createWorksheetChart',
        sheetName: 'People',
        chart: {
          name: 'Ambiguous series',
          chartType: 4,
          anchor: { left: 0, top: 0, width: 320, height: 200 },
          series: [{
            categoryRange: 'A2:A10',
            xValuesRange: 'B2:B10',
            valuesRange: 'C2:C10',
          }],
        },
      }],
    }),
    /categoryRange et xValuesRange/,
  );

  console.log('Workbook object tool input validation passed.');
} finally {
  await rm(buildDirectory, { recursive: true, force: true });
}
