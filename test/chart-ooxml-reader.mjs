import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import JSZip from 'jszip';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const temporaryDirectory = mkdtempSync(join(tmpdir(), 'chart-ooxml-reader-'));
const bundle = join(temporaryDirectory, 'chart-ooxml-reader.cjs');

const packageRelationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
const officeRelationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const spreadsheet = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const chart = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const drawing = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const drawingMain = 'http://schemas.openxmlformats.org/drawingml/2006/main';

function relationshipsXml(items) {
  return [
    `<Relationships xmlns="${packageRelationships}">`,
    ...items.map(item => `<Relationship Id="${item.id}" Type="${item.type}" Target="${item.target}"${item.external ? ' TargetMode="External"' : ''}/>`),
    '</Relationships>',
  ].join('');
}

function worksheetXml(drawingRelationshipId) {
  return `<worksheet xmlns="${spreadsheet}" xmlns:r="${officeRelationships}"><sheetData/><drawing r:id="${drawingRelationshipId}"/></worksheet>`;
}

const columnChartXml = [
  `<c:chartSpace xmlns:c="${chart}" xmlns:a="${drawingMain}">`,
  '<c:style val="10"/>',
  '<c:chart>',
  '<c:title><c:tx><c:rich><a:p><a:r><a:t>Quarterly sales</a:t></a:r></a:p></c:rich></c:tx></c:title>',
  '<c:plotArea>',
  '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>',
  '<c:ser><c:idx val="0"/><c:order val="0"/>',
  '<c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:pt idx="0"><c:v>North</c:v></c:pt></c:strCache></c:strRef></c:tx>',
  '<c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f></c:strRef></c:cat>',
  '<c:val><c:numRef><c:f>Data!$B$2:$B$5</c:f></c:numRef></c:val>',
  '<c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>',
	'<c:smooth val="0"/>',
  '</c:ser>',
  '<c:ser><c:idx val="1"/><c:order val="1"/>',
  '<c:tx><c:strRef><c:f>Data!$C$1</c:f><c:strCache><c:pt idx="0"><c:v>South</c:v></c:pt></c:strCache></c:strRef></c:tx>',
  '<c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f></c:strRef></c:cat>',
  '<c:val><c:numRef><c:f>Data!$C$2:$C$5</c:f></c:numRef></c:val>',
  '</c:ser>',
	'<c:dLbls><c:dLblPos val="t"/><c:showVal val="1"/><c:showPercent val="1"/><c:showBubbleSize val="1"/></c:dLbls>',
  '<c:gapWidth val="120"/><c:overlap val="-5"/><c:axId val="10"/><c:axId val="20"/>',
  '</c:barChart>',
  '<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/>',
  '<c:title><c:tx><c:rich><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title>',
  '<c:crossAx val="20"/></c:catAx>',
  '<c:valAx><c:axId val="20"/><c:scaling><c:min val="0"/><c:max val="500"/></c:scaling><c:axPos val="l"/>',
  '<c:majorGridlines/><c:numFmt formatCode="0"/><c:majorUnit val="100"/><c:crossAx val="10"/></c:valAx>',
  '</c:plotArea>',
  '<c:legend><c:legendPos val="r"/></c:legend>',
  '</c:chart>',
  '</c:chartSpace>',
].join('');

const secondaryLineChartXml = [
  `<c:chartSpace xmlns:c="${chart}" xmlns:a="${drawingMain}">`,
  '<c:style val="4"/><c:chart><c:plotArea>',
  '<c:lineChart><c:grouping val="standard"/>',
  '<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:pt idx="0"><c:v>Volume</c:v></c:pt></c:strCache></c:strRef></c:tx>',
  '<c:marker><c:symbol val="circle"/><c:size val="8"/></c:marker>',
  '<c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$5</c:f></c:numRef></c:val>',
  '<c:smooth val="1"/></c:ser><c:dLbls><c:dLblPos val="bestFit"/></c:dLbls><c:axId val="100"/><c:axId val="200"/></c:lineChart>',
  '<c:lineChart><c:grouping val="standard"/>',
  '<c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:strRef><c:f>Data!$C$1</c:f><c:strCache><c:pt idx="0"><c:v>Ratio</c:v></c:pt></c:strCache></c:strRef></c:tx>',
  '<c:marker><c:symbol val="circle"/></c:marker>',
  '<c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$C$2:$C$5</c:f></c:numRef></c:val>',
  '</c:ser><c:axId val="300"/><c:axId val="400"/></c:lineChart>',
  '<c:catAx><c:axId val="100"/><c:scaling/><c:axPos val="b"/><c:crossAx val="200"/></c:catAx>',
  '<c:valAx><c:axId val="200"/><c:scaling><c:min val="0"/><c:max val="100"/></c:scaling><c:axPos val="l"/><c:crossAx val="100"/></c:valAx>',
  '<c:catAx><c:axId val="300"/><c:scaling/><c:delete val="1"/><c:axPos val="t"/><c:crossAx val="400"/></c:catAx>',
  '<c:valAx><c:axId val="400"/><c:scaling><c:min val="0"/><c:max val="1"/></c:scaling><c:axPos val="r"/>',
  '<c:title><c:tx><c:rich><a:p><a:r><a:t>Ratio axis</a:t></a:r></a:p></c:rich></c:tx></c:title>',
  '<c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorUnit val="0.2"/><c:crossAx val="300"/></c:valAx>',
  '</c:plotArea><c:legend><c:legendPos val="b"/><c:layout><c:manualLayout><c:x val="0.2"/><c:y val="0.8"/></c:manualLayout></c:layout></c:legend></c:chart></c:chartSpace>',
].join('');

const pivotChartXml = columnChartXml.replace(
  '<c:chart>',
  '<c:pivotSource><c:name>PivotTable1</c:name><c:fmtId val="0"/></c:pivotSource><c:chart>',
);

const partiallyUnreadableChartXml = columnChartXml.replace(
  '<c:gapWidth val="120"/>',
  [
    '<c:ser><c:idx val="2"/><c:order val="2"/>',
    '<c:tx><c:v>Unreadable series</c:v></c:tx>',
    '<c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f></c:strRef></c:cat>',
    '<c:val><c:numRef/></c:val>',
    '</c:ser>',
    '<c:gapWidth val="120"/>',
  ].join(''),
);

const classicDrawingXml = [
  `<xdr:wsDr xmlns:xdr="${drawing}" xmlns:a="${drawingMain}" xmlns:c="${chart}" xmlns:r="${officeRelationships}">`,
  '<xdr:twoCellAnchor>',
  '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>',
  '<xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>',
  '<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Sales chart" descr="Regional quarterly sales"/></xdr:nvGraphicFramePr>',
  '<a:graphic><a:graphicData><c:chart r:id="rIdChart1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>',
  '</xdr:twoCellAnchor>',
  '<xdr:absoluteAnchor><xdr:pos x="635000" y="3810000"/><xdr:ext cx="7620000" cy="3810000"/>',
  '<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Ratio chart"/></xdr:nvGraphicFramePr>',
  '<a:graphic><a:graphicData><c:chart r:id="rIdChart2"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>',
  '</xdr:absoluteAnchor>',
  '<xdr:absoluteAnchor><xdr:pos x="635000" y="8000000"/><xdr:ext cx="5000000" cy="2500000"/>',
  '<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="5" name="Pivot chart"/></xdr:nvGraphicFramePr>',
  '<a:graphic><a:graphicData><c:chart r:id="rIdPivotChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>',
  '</xdr:absoluteAnchor>',
  '<xdr:absoluteAnchor><xdr:pos x="635000" y="11000000"/><xdr:ext cx="5000000" cy="2500000"/>',
  '<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="6" name="Partially readable chart"/></xdr:nvGraphicFramePr>',
  '<a:graphic><a:graphicData><c:chart r:id="rIdPartialChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>',
  '</xdr:absoluteAnchor>',
  '</xdr:wsDr>',
].join('');

const chartExDrawingXml = [
  `<xdr:wsDr xmlns:xdr="${drawing}" xmlns:a="${drawingMain}" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="${officeRelationships}">`,
  '<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>',
  '<xdr:ext cx="3810000" cy="2540000"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="4" name="Histogram"/></xdr:nvGraphicFramePr>',
  '<a:graphic><a:graphicData><cx:chart r:id="rIdChartEx"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>',
  '</xdr:oneCellAnchor></xdr:wsDr>',
].join('');

async function buildFixture() {
  const zip = new JSZip();
  zip.file('xl/workbook.xml', [
    `<workbook xmlns="${spreadsheet}" xmlns:r="${officeRelationships}"><sheets>`,
    '<sheet name="Data" sheetId="7" r:id="rIdData"/>',
    '<sheet name="Modern" sheetId="42" r:id="rIdModern"/>',
    '<sheet name="External" sheetId="99" r:id="rIdExternal"/>',
    '</sheets></workbook>',
  ].join(''));
  zip.file('xl/_rels/workbook.xml.rels', relationshipsXml([
    { id: 'rIdData', type: `${officeRelationships}/worksheet`, target: 'worksheets/custom-data.xml' },
    { id: 'rIdModern', type: `${officeRelationships}/worksheet`, target: 'worksheets/not-sheet2.xml' },
    { id: 'rIdExternal', type: `${officeRelationships}/worksheet`, target: 'https://example.invalid/external.xml', external: true },
  ]));
  zip.file('xl/worksheets/custom-data.xml', worksheetXml('rIdDrawing1'));
  zip.file('xl/worksheets/_rels/custom-data.xml.rels', relationshipsXml([
    { id: 'rIdDrawing1', type: `${officeRelationships}/drawing`, target: '../drawings/drawing1.xml' },
  ]));
  zip.file('xl/drawings/drawing1.xml', classicDrawingXml);
  zip.file('xl/drawings/_rels/drawing1.xml.rels', relationshipsXml([
    { id: 'rIdChart1', type: `${officeRelationships}/chart`, target: '../charts/chart1.xml' },
    { id: 'rIdChart2', type: `${officeRelationships}/chart`, target: '../charts/chart2.xml' },
    { id: 'rIdPivotChart', type: `${officeRelationships}/chart`, target: '../charts/chart3.xml' },
    { id: 'rIdPartialChart', type: `${officeRelationships}/chart`, target: '../charts/chart4.xml' },
  ]));
  zip.file('xl/charts/chart1.xml', columnChartXml);
  zip.file('xl/charts/chart2.xml', secondaryLineChartXml);
  zip.file('xl/charts/chart3.xml', pivotChartXml);
  zip.file('xl/charts/chart4.xml', partiallyUnreadableChartXml);

  zip.file('xl/worksheets/not-sheet2.xml', worksheetXml('rIdDrawing2'));
  zip.file('xl/worksheets/_rels/not-sheet2.xml.rels', relationshipsXml([
    { id: 'rIdDrawing2', type: `${officeRelationships}/drawing`, target: '../drawings/drawing2.xml' },
  ]));
  zip.file('xl/drawings/drawing2.xml', chartExDrawingXml);
  zip.file('xl/drawings/_rels/drawing2.xml.rels', relationshipsXml([
    { id: 'rIdChartEx', type: 'http://schemas.microsoft.com/office/2014/relationships/chartEx', target: '../charts/chartEx1.xml' },
  ]));
  return JSZip.loadAsync(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
}

try {
  await build({
    stdin: {
      contents: "export * from './src/react/view/excel/chart-ooxml-reader.ts';",
      resolveDir: root,
      sourcefile: 'chart-ooxml-reader-entry.ts',
      loader: 'ts',
    },
    bundle: true,
    format: 'cjs',
    logLevel: 'silent',
    outfile: bundle,
    platform: 'node',
    target: 'node22',
  });
  const { readOoxmlChartInventory } = await import(pathToFileURL(bundle).href);
  const zip = await buildFixture();
  const first = await readOoxmlChartInventory(zip);
  const second = await readOoxmlChartInventory(zip);

  assert.deepEqual(first.sheets.map(sheet => [sheet.sheetName, sheet.worksheetPart]), [
    ['Data', 'xl/worksheets/custom-data.xml'],
    ['Modern', 'xl/worksheets/not-sheet2.xml'],
  ], 'workbook relationships, not sheetId filenames, must select worksheet parts');

  const data = first.sheets[0];
  assert.equal(data.hasChartParts, true);
  assert.equal(data.unsupportedChartCount, 2, 'PivotChart and partially unreadable charts must remain native and non-editable');
  assert.equal(data.charts.length, 2);
  assert.equal(
    data.charts.some(item => item.name === 'Partially readable chart'),
    false,
    'one unreadable series must prevent partial hydration of the entire chart',
  );

  const columns = data.charts[0];
  assert.equal(columns.name, 'Sales chart');
  assert.equal(columns.alternativeText, 'Regional quarterly sales');
  assert.equal(columns.chartType, 51);
  assert.equal(columns.sourceRangeRef, undefined, 'hydrated charts must use explicit series, never sourceRangeRef and series together');
  assert.equal(columns.plotBy, 'columns');
  assert.equal(columns.title.text, 'Quarterly sales');
  assert.deepEqual(columns.legend, { visible: true, position: 'right' });
  assert.equal(columns.style, 10);
  assert.equal(columns.gapWidth, 120);
  assert.equal(columns.overlap, -5);
  assert.equal(columns.series.length, 2);
	assert.equal('smooth' in columns.series[0], false, 'column OOXML must not hydrate an unsupported smooth:false property');
	assert.equal(columns.series[0].dataLabels.showValue, true);
	assert.equal('position' in columns.series[0].dataLabels, false, 'unsupported column label positions must not cross the native-write boundary');
	assert.equal('showPercentage' in columns.series[0].dataLabels, false, 'unsupported percentage labels must not be hydrated on columns');
	assert.equal('showBubbleSize' in columns.series[0].dataLabels, false, 'unsupported bubble-size labels must not be hydrated on columns');
  assert.deepEqual(columns.series.map(series => [series.name, series.nameRange, series.categoryRange, series.valuesRange]), [
    [undefined, 'B1', 'A2:A5', 'B2:B5'],
    [undefined, 'C1', 'A2:A5', 'C2:C5'],
  ]);
  assert.deepEqual(columns.valueAxis, {
    visible: true,
    minimumScale: 0,
    maximumScale: 500,
    majorUnit: 100,
    minorUnit: null,
    logarithmic: false,
    reverseOrder: false,
    numberFormat: '0',
    majorGridlines: true,
    minorGridlines: false,
  });
  assert.ok(columns.anchor.left > 0 && columns.anchor.width > 0 && columns.anchor.height > 0);

  const line = data.charts[1];
  assert.equal(line.chartType, 65);
  assert.equal(line.series.length, 2);
  assert.equal(line.sourceRangeRef, undefined);
  assert.equal(line.series[0].name, undefined, 'a strRef cache must not become a competing literal name');
  assert.equal(line.series[0].nameRange, 'B1');
  assert.deepEqual(line.legend, { visible: true, position: 'custom' });
  assert.equal(line.series[0].axisGroup, 'primary');
  assert.equal(line.series[0].markerStyle, 'circle');
  assert.equal(line.series[0].smooth, true);
	assert.equal('dataLabels' in line.series[0], false, 'an empty label definition after compatibility filtering must be omitted');
  assert.equal(line.series[1].axisGroup, 'secondary');
  assert.equal(line.secondaryCategoryAxis.visible, false);
  assert.equal(line.secondaryValueAxis.title, 'Ratio axis');
  assert.equal(line.secondaryValueAxis.minimumScale, 0);
  assert.equal(line.secondaryValueAxis.maximumScale, 1);
  assert.equal(line.secondaryValueAxis.majorUnit, 0.2);
  assert.equal(
	line.secondaryValueAxis.numberFormat,
	'',
	'a source-linked OOXML axis format must hydrate as the explicit automatic-reset sentinel',
  );

  assert.deepEqual(
    first.sheets.flatMap(sheet => sheet.charts.map(item => item.id)),
    second.sheets.flatMap(sheet => sheet.charts.map(item => item.id)),
    'OOXML chart IDs must be deterministic',
  );

  const modern = first.sheets[1];
  assert.equal(modern.hasChartParts, true);
  assert.equal(modern.unsupportedChartCount, 1);
  assert.deepEqual(modern.charts, [], 'chartEx must not become a fake classic SheetChartData');
  assert.ok(first.warnings.some(warning => warning.includes('chartEx')));
  assert.ok(first.warnings.some(warning => warning.includes('PivotChart')));
  assert.ok(first.warnings.some(warning => warning.includes('série OOXML illisible')));
  assert.ok(first.warnings.some(warning => warning.includes('externe')));

  let blockedAsyncCalls = 0;
  const budgetBlockedZip = {
    file(path) {
      const entry = zip.file(path);
      if (!entry) return null;
      return {
        dir: entry.dir,
        _data: entry._data,
        async(type) {
          blockedAsyncCalls += 1;
          return entry.async(type);
        },
      };
    },
  };
  const budgetBlocked = await readOoxmlChartInventory(
    budgetBlockedZip,
    undefined,
    { maxTotalXmlCharacters: 1 },
  );
  assert.equal(blockedAsyncCalls, 0, 'declared XML over the remaining budget must never be inflated');
  assert.deepEqual(budgetBlocked.sheets, []);

  const inflatedParts = [];
  const attemptBoundedZip = {
    file(path) {
      const entry = zip.file(path);
      if (!entry) return null;
      return {
        dir: entry.dir,
        _data: entry._data,
        async(type) {
          inflatedParts.push(path);
          return entry.async(type);
        },
      };
    },
  };
  await readOoxmlChartInventory(attemptBoundedZip, undefined, { maxCharts: 1 });
  assert.ok(inflatedParts.includes('xl/charts/chart1.xml'));
  assert.ok(!inflatedParts.includes('xl/charts/chart2.xml'), 'failed or excess chart attempts must consume the global budget');

  console.log('OOXML chart reader passed: classic multi-series, secondary axis, chartEx preservation signal.');
} finally {
  rmSync(temporaryDirectory, { recursive: true, force: true });
}
