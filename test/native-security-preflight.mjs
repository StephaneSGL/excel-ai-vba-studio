import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { createHash, randomUUID } from 'node:crypto';
import { promisify } from 'node:util';
import {
  copyFile,
  mkdtemp,
  readFile,
  rm,
  writeFile,
} from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import * as XLSX from 'xlsx';

const execFileAsync = promisify(execFile);
const root = resolve(import.meta.dirname, '..');
const scriptPath = join(root, 'scripts', 'office-ai-apply-edits.ps1');
const powerShell = process.platform === 'win32'
  ? join(
      process.env.SystemRoot || 'C:\\Windows',
      'System32',
      'WindowsPowerShell',
      'v1.0',
      'powershell.exe',
    )
  : 'pwsh';

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
}

function payload(expectedWorkbookSha256, operations, version = 2) {
  return {
    version,
    transactionId: randomUUID().toLowerCase(),
    expectedWorkbookSha256,
    operations,
  };
}

function table(name, rangeRef) {
  return {
    id: `table:${name.toLowerCase()}`,
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

function chart(name, sourceRangeRef) {
  return {
    id: `chart:${name.toLowerCase()}`,
    name,
    chartType: 51,
    sourceRangeRef,
    plotBy: 'columns',
    anchor: { left: 10, top: 10, width: 300, height: 200 },
  };
}

async function runNativeEdit(workbookPath, operationsPath, timeout = 30_000) {
  return execFileAsync(powerShell, [
    '-NoLogo',
    '-NoProfile',
    '-NonInteractive',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    scriptPath,
    '-WorkbookPath',
    workbookPath,
    '-OperationsPath',
    operationsPath,
  ], {
    cwd: root,
    windowsHide: true,
    timeout,
    maxBuffer: 1024 * 1024,
  });
}

async function expectPreComRefusal(workbookPath, operationsPath, pattern) {
  await assert.rejects(
    runNativeEdit(workbookPath, operationsPath),
    error => {
      const output = String(error.stdout ?? '') + String(error.stderr ?? '');
      return pattern.test(output) && !/OWNED_EXCEL_PID\|/.test(output);
    },
    'PRECOM_GUARD_ABSENCE: invalid native operations must be rejected before Excel COM starts',
  );
}

if (process.platform !== 'win32') {
  console.log('Native security preflight skipped outside Windows.');
  process.exit(0);
}

const testRoot = await mkdtemp(join(tmpdir(), 'excel-native-security-'));
try {
  const workbookPath = join(testRoot, 'security.xlsm');
  const operationsPath = join(testRoot, 'operations.json');
  await copyFile(
    join(root, 'test', 'fixtures', 'NativeEditingSynthetic.xlsm'),
    workbookPath,
  );
  const workbookBytes = await readFile(workbookPath);
  const workbookSha256 = sha256(workbookBytes);

  await writeFile(
    `${workbookPath}:Zone.Identifier`,
    '[ZoneTransfer]\r\nZoneId=3\r\n',
    'utf8',
  );
  await writeFile(
    operationsPath,
    JSON.stringify(payload(workbookSha256, [{
      kind: 'cell',
      sheetName: 'Resume',
      row: 1,
      column: 1,
      value: { kind: 'text', value: 'new' },
    }])),
    'utf8',
  );
  await expectPreComRefusal(
    workbookPath,
    operationsPath,
    /Internet or Restricted Zone \(ZoneId=3\)/,
  );

  await writeFile(
    `${workbookPath}:Zone.Identifier`,
    '[ZoneTransfer]\r\nZoneId=2\r\nZoneId=3\r\n',
    'utf8',
  );
  await expectPreComRefusal(
    workbookPath,
    operationsPath,
    /Zone\.Identifier is missing or ambiguous/,
  );

  await writeFile(`${workbookPath}:Zone.Identifier`, Buffer.alloc(64 * 1024 + 1, 0x41));
  await expectPreComRefusal(
    workbookPath,
    operationsPath,
    /Zone\.Identifier exceeds the 65536-byte safety limit/,
  );

  await rm(`${workbookPath}:Zone.Identifier`);
  await writeFile(
    operationsPath,
    JSON.stringify(payload(workbookSha256, [{
      kind: 'createTable',
      sheetName: 'Resume',
      table: table('TooLarge', 'A1:A1000001'),
    }])),
    'utf8',
  );
  await expectPreComRefusal(
    workbookPath,
    operationsPath,
    /exceeds the 1000000-cell native Excel safety limit/,
  );

  await writeFile(
    operationsPath,
    JSON.stringify(payload(workbookSha256, [{
      kind: 'createTable',
      sheetName: 'Resume',
      table: { ...table('NoHeaders', 'A1:C5'), headerRow: false },
    }])),
    'utf8',
  );
  await expectPreComRefusal(
    workbookPath,
    operationsPath,
    /headerRow=false is disabled because Excel can move worksheet cells/,
  );

  await writeFile(
    operationsPath,
    JSON.stringify(payload(
      workbookSha256,
      Array.from({ length: 6 }, (_, index) => ({
        kind: 'createChart',
        sheetName: 'Resume',
        chart: chart(`Chart${index + 1}`, 'A1:A1000000'),
      })),
    )),
    'utf8',
  );
  await expectPreComRefusal(
    workbookPath,
    operationsPath,
    /exceed the 5000000-cell transaction safety budget/,
  );

	await writeFile(
		operationsPath,
		JSON.stringify(payload(workbookSha256, [
			{ kind: 'createTable', sheetName: 'Resume', table: table('FirstOverlap', '$A$1:$C$10') },
			{ kind: 'createTable', sheetName: 'Resume', table: table('SecondOverlap', 'C10:E20') },
		])),
		'utf8',
	);
	await expectPreComRefusal(
		workbookPath,
		operationsPath,
		/Native tables FirstOverlap and SecondOverlap overlap/,
	);

	await writeFile(
		operationsPath,
		JSON.stringify(payload(workbookSha256, [{
			kind: 'createChart',
			sheetName: 'Resume',
			chart: {
				...chart('HomogeneousCombo', undefined),
				chartType: -4152,
				series: [
					{ id: 'series:one', chartType: 51, valuesRange: 'B2:B5' },
					{ id: 'series:two', chartType: 51, valuesRange: 'C2:C5' },
				],
			},
		}])),
		'utf8',
	);
	await expectPreComRefusal(
		workbookPath,
		operationsPath,
		/homogeneous explicit series require their concrete type as the top-level chart type/,
	);

	await writeFile(
		operationsPath,
		JSON.stringify(payload(workbookSha256, [{
			kind: 'createChart',
			sheetName: 'Resume',
			chart: {
				...chart('SingleSeriesTypeMismatch', undefined),
				chartType: 51,
				series: [
					{ id: 'series:line', chartType: 4, valuesRange: 'B2:B5' },
				],
			},
		}])),
		'utf8',
	);
	await expectPreComRefusal(
		workbookPath,
		operationsPath,
		/homogeneous explicit series require their concrete type as the top-level chart type/,
	);

	await writeFile(
		operationsPath,
		JSON.stringify(payload(workbookSha256, [{
			kind: 'createChart',
			sheetName: 'Resume',
			chart: {
				...chart('UnnormalizedCombo', undefined),
				chartType: 51,
				series: [
					{ id: 'series:column', chartType: 51, valuesRange: 'B2:B5' },
					{ id: 'series:line', chartType: 4, valuesRange: 'C2:C5' },
				],
			},
		}])),
		'utf8',
	);
	await expectPreComRefusal(
		workbookPath,
		operationsPath,
		/heterogeneous explicit series require xlCombo/,
	);

	await writeFile(
		operationsPath,
		JSON.stringify(payload(workbookSha256, [{
			kind: 'createChart',
			sheetName: 'Resume',
			chart: {
				...chart('InvalidLogAxis', 'A1:B5'),
				valueAxis: { visible: true, logarithmic: true, minimumScale: 0 },
			},
		}])),
		'utf8',
	);
	await expectPreComRefusal(
		workbookPath,
		operationsPath,
		/logarithmic scale cannot use a non-positive bound/,
	);

	await writeFile(
		operationsPath,
		JSON.stringify(payload(workbookSha256, [{
			kind: 'createChart',
			sheetName: 'Resume',
			chart: {
				...chart('InvalidBubbleSizes', undefined),
				series: [{
					id: 'series:column',
					chartType: 51,
					valuesRange: 'B2:B5',
					bubbleSizesRange: 'C2:C5',
				}],
			},
		}])),
		'utf8',
	);
	await expectPreComRefusal(
		workbookPath,
		operationsPath,
		/bubbleSizesRange requires a bubble series type/,
	);

	await writeFile(
		operationsPath,
		JSON.stringify(payload(workbookSha256, [{
			kind: 'createChart',
			sheetName: 'Resume',
			chart: {
				...chart('EmptyDataLabels', undefined),
				series: [{
					id: 'series:column',
					chartType: 51,
					valuesRange: 'B2:B5',
					dataLabels: {},
				}],
			},
		}])),
		'utf8',
	);
	await expectPreComRefusal(
		workbookPath,
		operationsPath,
		/dataLabels must explicitly define at least one show option/,
	);

	await writeFile(
		operationsPath,
		JSON.stringify(payload(workbookSha256, [{
			kind: 'createChart',
			sheetName: 'Resume',
			chart: {
				...chart('DisabledDataLabelPosition', undefined),
				series: [{
					id: 'series:column',
					chartType: 51,
					valuesRange: 'B2:B5',
					dataLabels: { showValue: false, position: 'outsideEnd' },
				}],
			},
		}])),
		'utf8',
	);
	await expectPreComRefusal(
		workbookPath,
		operationsPath,
		/dataLabels.position requires at least one enabled show option/,
	);

	await writeFile(
		operationsPath,
		JSON.stringify(payload(workbookSha256, [{
			kind: 'createChart',
			sheetName: 'Resume',
			chart: {
				...chart('PositionOnlyDataLabels', undefined),
				series: [{
					id: 'series:column',
					chartType: 51,
					valuesRange: 'B2:B5',
					dataLabels: { position: 'outsideEnd' },
				}],
			},
		}])),
		'utf8',
	);
	await expectPreComRefusal(
		workbookPath,
		operationsPath,
		/dataLabels must explicitly define at least one show option/,
	);

	await writeFile(
		operationsPath,
		JSON.stringify(payload(workbookSha256, [{
			kind: 'createChart',
			sheetName: 'Resume',
			chart: {
				...chart('InvalidDataLabelPosition', undefined),
				series: [{
					id: 'series:column',
					chartType: 51,
					valuesRange: 'B2:B5',
					dataLabels: { showValue: true, position: 'above' },
				}],
			},
		}])),
		'utf8',
	);
	await expectPreComRefusal(
		workbookPath,
		operationsPath,
		/data label position is not supported by this chart type/,
	);

	await writeFile(
		operationsPath,
		JSON.stringify(payload(workbookSha256, [{
			kind: 'createChart',
			sheetName: 'Resume',
			chart: {
				...chart('MixedBubbleCombo', undefined),
				chartType: -4152,
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
			},
		}])),
		'utf8',
	);
	await expectPreComRefusal(
		workbookPath,
		operationsPath,
		/cannot mix bubble and non-bubble series/,
	);

	await writeFile(
		operationsPath,
		JSON.stringify(payload(workbookSha256, [{
			kind: 'createChart',
			sheetName: 'Resume',
			chart: {
				...chart('InvalidPrimaryCategoryScale', undefined),
				chartType: -4152,
				categoryAxis: { visible: true, minimumScale: 0 },
				series: [
					{ id: 'series:column', chartType: 51, axisGroup: 'primary', valuesRange: 'B2:B5' },
					{ id: 'series:scatter', chartType: 74, axisGroup: 'secondary', xValuesRange: 'A2:A5', valuesRange: 'C2:C5' },
				],
			},
		}])),
		'utf8',
	);
	await expectPreComRefusal(
		workbookPath,
		operationsPath,
		/categoryAxis numeric scale settings require a scatter or bubble series on the same axis group/,
	);

	await writeFile(
		operationsPath,
		JSON.stringify(payload(workbookSha256, [{
			kind: 'updateChart',
			sheetName: 'Resume',
			name: 'Chart1',
			chart: chart('Chart1', 'A1:A2'),
			preserveSeries: false,
			allowSeriesFormattingChange: true,
		}])),
		'utf8',
	);
	await expectPreComRefusal(
		workbookPath,
		operationsPath,
		/allowSeriesFormattingChange requires preserveSeries=true/,
	);

  await writeFile(
    operationsPath,
    JSON.stringify(payload(workbookSha256, [{
      kind: 'updateChart',
      sheetName: 'Resume',
      name: 'Chart1',
      chart: chart('Chart1', 'A1:A2'),
      preserveAnchor: 'yes',
    }])),
    'utf8',
  );
  await expectPreComRefusal(
    workbookPath,
    operationsPath,
    /Native updateChart preserveAnchor must be a boolean/,
  );

  const trustedZone = Buffer.from('[ZoneTransfer]\r\nZoneId=2\r\n', 'utf8');
  const customStream = Buffer.from('preserve-this-stream', 'utf8');
  await writeFile(`${workbookPath}:Zone.Identifier`, trustedZone);
  await writeFile(`${workbookPath}:ExcelAiVbaTest`, customStream);
  await writeFile(
    operationsPath,
    JSON.stringify(payload(workbookSha256, [{
      kind: 'cell',
      sheetName: 'Resume',
      row: 1,
      column: 1,
      value: { kind: 'text', value: 'new' },
    }])),
    'utf8',
  );
  const result = await runNativeEdit(workbookPath, operationsPath, 180_000);
  assert.match(String(result.stdout), /OWNED_EXCEL_PID\|\d+\|\d{15,19}/);
  const jsonLine = String(result.stdout)
    .split(/\r?\n/)
    .find(line => line.trim().startsWith('{'));
  assert.ok(jsonLine, 'native edit result JSON is missing');
  const parsed = JSON.parse(jsonLine);
  assert.equal(parsed.ok, true);
  assert.deepEqual(await readFile(`${workbookPath}:Zone.Identifier`), trustedZone);
  assert.deepEqual(await readFile(`${workbookPath}:ExcelAiVbaTest`), customStream);
  assert.deepEqual(await readFile(`${parsed.backupPath}:Zone.Identifier`), trustedZone);
  assert.deepEqual(await readFile(`${parsed.backupPath}:ExcelAiVbaTest`), customStream);

  const reopened = XLSX.read(await readFile(workbookPath), { type: 'buffer' });
  assert.equal(reopened.Sheets.Resume.A1.v, 'new');

  console.log('Native security preflight passed: MOTW, ADS, range budgets and PID identity are enforced.');
} finally {
  await rm(testRoot, { recursive: true, force: true });
}
