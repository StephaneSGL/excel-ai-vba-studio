import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '..');
const buildDirectory = await mkdtemp(join(tmpdir(), 'excel-table-roundtrip-test-'));
const outputPath = join(buildDirectory, 'table-roundtrip.cjs');

function setRow(rows, rowNumber, values) {
  rows[rowNumber - 1] = {
    cells: Object.fromEntries(values.map((text, index) => [index, { text }])),
  };
}

function table(id, name, rangeRef, styleName) {
  return {
    id,
    name,
    displayName: name,
    rangeRef,
    headerRow: true,
    totalsRow: false,
    style: {
      name: styleName,
      showFirstColumn: false,
      showLastColumn: false,
      showRowStripes: true,
      showColumnStripes: false,
    },
  };
}

const rows = { len: 50 };
for (const [start, end, label] of [
  [1, 5, 'First'],
  [20, 30, 'Second'],
  [40, 50, 'Third'],
]) {
  setRow(rows, start, ['Name', 'Value', 'Status']);
  for (let row = start + 1; row <= end; row += 1) {
    setRow(rows, row, [`${label}-${row}`, String(row), row % 2 ? 'Open' : 'Closed']);
  }
}

const threeTableSheet = {
  name: 'Multiple tables',
  rows,
  cols: { len: 3 },
  tables: [
    table('table:1:A1:C5', 'FirstTable', 'A1:C5', 'TableStyleMedium2'),
    table('table:1:A20:C30', 'SecondTable', 'A20:C30', 'TableStyleLight9'),
    table('table:1:A40:C50', 'ThirdTable', 'A40:C50', 'TableStyleDark3'),
  ],
};

globalThis.window = {
  acquireVsCodeApi: undefined,
  addEventListener() {},
};
Object.defineProperty(globalThis, 'navigator', {
  configurable: true,
  value: { userAgent: 'node' },
});
globalThis.localStorage = { getItem() { return null; }, setItem() {} };

try {
  await build({
    stdin: {
      contents: [
        "export { loadSheets } from './src/react/view/excel/excel_reader.ts';",
        "export { buildExcelWorkbookBuffer } from './src/react/view/excel/excel_writer.ts';",
        "export { createSheetTableDefinition } from './src/react/view/excel/x-spreadsheet/index.ts';",
        "export { isWorkbookTableNameAvailableInSheets } from './src/react/view/excel/x-spreadsheet/index.ts';",
        "export { excelTableNameComparisonKey, isValidExcelTableName, normalizeExcelTableName } from './src/common/excelWorkbookObjects.ts';",
      ].join('\n'),
      resolveDir: root,
      sourcefile: 'table-roundtrip-entry.ts',
      loader: 'ts',
    },
    outfile: outputPath,
    bundle: true,
    format: 'cjs',
    platform: 'node',
    target: 'node22',
    logLevel: 'silent',
    tsconfig: join(root, 'tsconfig.json'),
    loader: {
      '.css': 'empty',
      '.less': 'empty',
      '.svg': 'dataurl',
      '.png': 'dataurl',
    },
  });
  const {
    buildExcelWorkbookBuffer,
    createSheetTableDefinition,
    excelTableNameComparisonKey,
    isWorkbookTableNameAvailableInSheets,
    isValidExcelTableName,
    loadSheets,
    normalizeExcelTableName,
  } = await import(
    `${pathToFileURL(outputPath).href}?cache=${Date.now()}`
  );

  const firstDefinition = createSheetTableDefinition('A1:C5', {
    idScope: 'sheet-1',
    sheetName: 'Multiple tables',
  }, { name: 'People' });
  const secondDefinition = createSheetTableDefinition('A20:C30', {
    idScope: 'sheet-1',
    sheetName: 'Multiple tables',
    existingTables: [firstDefinition],
    usedNames: ['People'],
  }, { name: 'People' });
  assert.equal(secondDefinition.rangeRef, 'A20:C30');
  assert.equal(secondDefinition.name, 'People_2');
  assert.throws(
    () => createSheetTableDefinition('A4:C8', {
      idScope: 'sheet-1',
      sheetName: 'Multiple tables',
      existingTables: [firstDefinition],
    }),
    /chevauche/i,
  );
	assert.throws(
		() => createSheetTableDefinition('A1:C2', {
			idScope: 'sheet-1',
			sheetName: 'Totals',
		}, { headerRow: false, totalsRow: true }),
		/pas assez de lignes/,
		'the final range needs an internal header, data and totals row even when headers are hidden',
	);

  const decomposedUnicodeName = 'E\u0301quipe_2026';
  assert.equal(normalizeExcelTableName(decomposedUnicodeName), 'Équipe_2026');
  assert.equal(excelTableNameComparisonKey(decomposedUnicodeName), excelTableNameComparisonKey('ÉQUIPE_2026'));
  assert.equal(isValidExcelTableName('Équipe_2026'), true);
  assert.equal(isValidExcelTableName('XFD1048576'), false, 'the last real A1 cell is not a valid table name');
  assert.equal(isValidExcelTableName('XFE1048576'), true, 'an identifier beyond Excel A1 bounds is not a cell reference');
  assert.equal(isValidExcelTableName('R1C1'), false, 'a real R1C1 reference is not a valid table name');
  assert.equal(isValidExcelTableName('R1048577C1'), true, 'an identifier beyond Excel R1C1 bounds is not a cell reference');
  const canonicalCollision = createSheetTableDefinition('E1:F5', {
    idScope: 'sheet-1',
    sheetName: 'Multiple tables',
    usedNames: ['Équipe_2026'],
  }, { name: decomposedUnicodeName });
  assert.equal(canonicalCollision.name, 'Équipe_2026_2');

  const collidingIdsAcrossSheets = [
    { tables: [table('shared-id', 'GlobalName', 'A1:B5', 'TableStyleMedium2')] },
    { tables: [table('shared-id', 'OtherName', 'A1:B5', 'TableStyleMedium2')] },
  ];
  assert.equal(
    isWorkbookTableNameAvailableInSheets(collidingIdsAcrossSheets, 1, 'globalname', 'shared-id'),
    false,
    'a colliding id on another sheet must never hide a workbook-global name conflict',
  );
  assert.equal(
    isWorkbookTableNameAvailableInSheets(collidingIdsAcrossSheets, 0, 'GlobalName', 'shared-id'),
    true,
    'an edited table may keep its own name on the active sheet',
  );
  assert.equal(
    isWorkbookTableNameAvailableInSheets(collidingIdsAcrossSheets, 0, 'othername', 'shared-id'),
    false,
    'renaming to a table name used on another sheet must be rejected',
  );
  assert.equal(
    isWorkbookTableNameAvailableInSheets(collidingIdsAcrossSheets, 1, 'GLOBALNAME'),
    false,
    'new table names must be checked case-insensitively across the workbook',
  );

  const buffer = await buildExcelWorkbookBuffer([threeTableSheet]);
  const zip = await JSZip.loadAsync(buffer);
  const tableEntries = Object.keys(zip.files)
    .filter(name => /^xl\/tables\/table\d+\.xml$/i.test(name))
    .sort();
  assert.equal(tableEntries.length, 3, 'the XLSX package must contain three table parts');
  const refs = [];
  for (const entryName of tableEntries) {
    const xml = await zip.file(entryName).async('string');
    refs.push(/\bref="([A-Z0-9:]+)"/.exec(xml)?.[1]);
  }
  assert.deepEqual(refs.sort(), ['A1:C5', 'A20:C30', 'A40:C50'].sort());

  const arrayBuffer = buffer.buffer.slice(
    buffer.byteOffset,
    buffer.byteOffset + buffer.byteLength,
  );
  const reloaded = await loadSheets(arrayBuffer, 'xlsx');
  assert.equal(reloaded.sheets.length, 1);
  assert.deepEqual(
    reloaded.sheets[0].tables.map(item => ({
      name: item.name,
      rangeRef: item.rangeRef,
      style: item.style.name,
    })),
    [
      { name: 'FirstTable', rangeRef: 'A1:C5', style: 'TableStyleMedium2' },
      { name: 'SecondTable', rangeRef: 'A20:C30', style: 'TableStyleLight9' },
      { name: 'ThirdTable', rangeRef: 'A40:C50', style: 'TableStyleDark3' },
    ],
  );

  const unicodeRows = { len: 3 };
  setRow(unicodeRows, 1, ['Nom', 'Statut']);
  setRow(unicodeRows, 2, ['Zoé', 'Actif']);
  setRow(unicodeRows, 3, ['André', 'Inactif']);
  const unicodeBuffer = await buildExcelWorkbookBuffer([{
    name: 'Unicode',
    rows: unicodeRows,
    cols: { len: 2 },
    tables: [table('table:unicode:A1:B3', decomposedUnicodeName, 'A1:B3', 'TableStyleMedium2')],
  }]);
  const unicodeZip = await JSZip.loadAsync(unicodeBuffer);
  const unicodeTableEntry = Object.keys(unicodeZip.files).find(name => /^xl\/tables\/table\d+\.xml$/i.test(name));
  assert.ok(unicodeTableEntry, 'the Unicode table must be written as a real OOXML table part');
  const unicodeTableXml = await unicodeZip.file(unicodeTableEntry).async('string');
  assert.match(unicodeTableXml, /name="Équipe_2026"/);
  assert.match(unicodeTableXml, /displayName="Équipe_2026"/);
  const unicodeReloaded = await loadSheets(
    unicodeBuffer.buffer.slice(
      unicodeBuffer.byteOffset,
      unicodeBuffer.byteOffset + unicodeBuffer.byteLength,
    ),
    'xlsx',
  );
  assert.equal(unicodeReloaded.sheets[0].tables[0].name, 'Équipe_2026');
  assert.equal(unicodeReloaded.sheets[0].tables[0].displayName, 'Équipe_2026');

  const overlappingSheet = structuredClone(threeTableSheet);
  overlappingSheet.tables[1].rangeRef = 'A4:C30';
  await assert.rejects(
    () => buildExcelWorkbookBuffer([overlappingSheet]),
    /overlap/i,
    'overlapping native Excel tables must be rejected before writing',
  );

  const duplicateNameSheets = [
    structuredClone(threeTableSheet),
    {
      ...structuredClone(threeTableSheet),
      name: 'Second sheet',
      tables: [table('table:second:A1:C5', 'firsttable', 'A1:C5', 'TableStyleMedium2')],
    },
  ];
  await assert.rejects(
    () => buildExcelWorkbookBuffer(duplicateNameSheets),
    /nom.*table|table.*nom|dupli/i,
    'the writer must reject a workbook-global table-name duplicate injected below the UI',
  );
	const shortTotalsSheet = {
		name: 'Short totals',
		rows: { len: 2 },
		cols: { len: 2 },
		tables: [{
			...table('table:short:A1:B2', 'ShortTotals', 'A1:B2', 'TableStyleMedium2'),
			totalsRow: true,
		}],
	};
	await assert.rejects(
		() => buildExcelWorkbookBuffer([shortTotalsSheet]),
		/does not contain a data row/,
		'the writer must enforce header plus data plus totals in the final range',
	);
  await assert.rejects(
    () => buildExcelWorkbookBuffer([{
      name: 'Unicode one',
      rows: unicodeRows,
      cols: { len: 2 },
      tables: [table('table:unicode-one:A1:B3', 'Équipe_2026', 'A1:B3', 'TableStyleMedium2')],
    }, {
      name: 'Unicode two',
      rows: unicodeRows,
      cols: { len: 2 },
      tables: [table('table:unicode-two:A1:B3', 'E\u0301QUIPE_2026', 'A1:B3', 'TableStyleMedium2')],
    }]),
    /Duplicate Excel table name/,
    'the writer must reject canonically equivalent Unicode names workbook-wide',
  );

  const fixture = resolve(root, 'test/fixtures/Excel-AI-VBA-Studio-Demo-base.xlsx');
  assert.ok(existsSync(fixture), `Missing fixture: ${fixture}`);
  const fixtureBuffer = readFileSync(fixture);
  const fixtureData = await loadSheets(
    fixtureBuffer.buffer.slice(
      fixtureBuffer.byteOffset,
      fixtureBuffer.byteOffset + fixtureBuffer.byteLength,
    ),
    'xlsx',
  );
  assert.equal(
    fixtureData.sheets.reduce((total, sheet) => total + (sheet.tables?.length ?? 0), 0),
    3,
    'the existing three native tables must survive workbook import',
  );
  const fixtureRoundTrip = await buildExcelWorkbookBuffer(fixtureData.sheets);
  const reloadedFixture = await loadSheets(
    fixtureRoundTrip.buffer.slice(
      fixtureRoundTrip.byteOffset,
      fixtureRoundTrip.byteOffset + fixtureRoundTrip.byteLength,
    ),
    'xlsx',
  );
  assert.equal(
    reloadedFixture.sheets.reduce((total, sheet) => total + (sheet.tables?.length ?? 0), 0),
    3,
    'the existing three native tables must survive import and export',
  );

  console.log(
    'Native Excel table round-trip passed: disjoint ranges and Unicode NFC names survive the OOXML engine.',
  );
} finally {
  await rm(buildDirectory, { recursive: true, force: true });
}
