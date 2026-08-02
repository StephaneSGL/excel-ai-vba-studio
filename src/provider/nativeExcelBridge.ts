import { execFile } from 'child_process';
import { createHash, randomUUID } from 'crypto';
import { createReadStream, promises as fs } from 'fs';
import { tmpdir } from 'os';
import { dirname, extname, isAbsolute, join, resolve } from 'path';
import { TextDecoder } from 'util';
import * as vscode from 'vscode';
import {
    type NativeExcelConditionalFormattingRule,
    type NativeExcelEditOperation,
    type NativeExcelEditPayload,
    type NativeExcelEditResult,
} from '@/common/nativeExcelEdits';
import {
	canonicalChartTypeForSeries,
	chartAxisGroupSupportsCategoryScale,
	chartDataLabelsHaveEnabledShowOption,
	chartDataLabelsHaveExplicitShowOption,
	chartSeriesSupportsBubbleSizes,
	chartSeriesSupportsDataLabelPosition,
	chartSeriesSupportsPercentageDataLabels,
	chartSeriesSupportsSmooth,
	chartSeriesTypesCanCoexist,
    chartTypeSupportsAxes,
    chartTypeSupportsGapWidth,
    chartTypeSupportsOverlap,
    chartTypeSupportsSecondaryAxes,
    excelTableNameComparisonKey,
    EXCEL_CHART_TYPES,
    isChartTypeCreatable,
    isValidExcelTableName,
	minimumExcelTableRangeRows,
    normalizeExcelTableName,
    simpleA1RangesOverlap,
    type SheetChartAxisData,
    type SheetChartData,
    type SheetChartSeriesData,
    type SheetTableData,
} from '@/common/excelWorkbookObjects';
import {
    assertOoxmlHasNoXlmMacroSheetsForAutomation,
    assertOoxmlPackageUnsignedForMutation,
} from '@/common/ooxmlPackageSignature';
import { assertNoReparsePointChain } from '@/excelAiVbaStudio/security';

const MAX_NATIVE_EDIT_OPERATIONS = 10_000;
const MAX_CONDITIONAL_FORMATTING_ADDS_PER_SHEET = 64;
const MAX_NATIVE_EDIT_PAYLOAD_BYTES = 4 * 1024 * 1024;
const NATIVE_EDIT_TIMEOUT_MS = 120_000;
const MAX_WORKBOOK_OBJECTS_PER_SHEET = 512;
const MAX_CHART_SERIES = 255;
const MAX_WORKBOOK_OBJECT_RANGE_CELLS = 1_000_000;
const MAX_WORKBOOK_OBJECT_TRANSACTION_RANGE_CELLS = 5_000_000;
const MAX_ZONE_IDENTIFIER_BYTES = 64 * 1024;
const VALID_CHART_TYPES = new Set(EXCEL_CHART_TYPES.map(option => option.value));

function getExtensionRoot(): string {
    const installed = vscode.extensions.getExtension(
        'steph-tools.excel-ai-vba-studio'
    );
    return installed?.extensionUri.fsPath ?? resolve(__dirname, '..');
}

function getPowerShellPath(): string {
    return process.env.SystemRoot
        ? join(
            process.env.SystemRoot,
            'System32',
            'WindowsPowerShell',
            'v1.0',
            'powershell.exe'
        )
        : 'powershell.exe';
}

interface OwnedExcelProcessIdentity {
    processId: number;
    startTimeUtcTicks: string;
}

async function terminateExactProcess(identity: OwnedExcelProcessIdentity): Promise<void> {
    const { processId, startTimeUtcTicks } = identity;
    if (
        process.platform !== 'win32' ||
        !Number.isSafeInteger(processId) ||
        processId <= 0 ||
        !/^\d{15,19}$/.test(startTimeUtcTicks)
    ) {
        return;
    }
    const verifyAndStopScript = [
        '$expectedPid = [int]$args[0]',
        '$expectedTicks = [long]$args[1]',
        '$candidate = Get-Process -Id $expectedPid -ErrorAction SilentlyContinue',
        'if ($null -eq $candidate) { exit 0 }',
        'try {',
        "  if (-not [string]::Equals($candidate.ProcessName, 'EXCEL', [StringComparison]::OrdinalIgnoreCase)) { exit 0 }",
        '  if ($candidate.StartTime.ToUniversalTime().Ticks -ne $expectedTicks) { exit 0 }',
        '  $candidate.Kill()',
        '  [void]$candidate.WaitForExit(5000)',
        '} finally { $candidate.Dispose() }',
    ].join('; ');
    await new Promise<void>(resolvePromise => {
        execFile(
            getPowerShellPath(),
            [
                '-NoLogo',
                '-NoProfile',
                '-NonInteractive',
                '-ExecutionPolicy',
                'Bypass',
                '-Command',
                verifyAndStopScript,
                String(processId),
                startTimeUtcTicks,
            ],
            { windowsHide: true },
            () => resolvePromise()
        );
    });
}

function parseOwnedExcelProcessId(output: string): OwnedExcelProcessIdentity | undefined {
    const match = output.match(
        /(?:^|\r?\n)OWNED_EXCEL_PID\|(\d+)\|(\d{15,19})(?:\r?\n|$)/
    );
    if (!match) {
        return undefined;
    }
    const processId = Number.parseInt(match[1], 10);
    if (!Number.isSafeInteger(processId) || processId <= 0) return undefined;
    return { processId, startTimeUtcTicks: match[2] };
}

function parseZoneIdentifier(bytes: Uint8Array, workbookPath: string): number {
    let text: string;
    try {
        text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    } catch {
        throw new Error(
            `Native Excel automation refused: Zone.Identifier is not valid UTF-8: ${workbookPath}`
        );
    }
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    if (!text || text.includes('\0') || /\r(?!\n)/.test(text)) {
        throw new Error(
            `Native Excel automation refused: Zone.Identifier is malformed: ${workbookPath}`
        );
    }

    let section = '';
    let zoneTransferSections = 0;
    const zoneIds: number[] = [];
    for (const rawLine of text.split(/\r?\n/)) {
        const line = rawLine.trim();
        if (!line || line.startsWith(';') || line.startsWith('#')) continue;
        const sectionMatch = /^\[([^\[\]\r\n]+)\]$/.exec(line);
        if (sectionMatch) {
            section = sectionMatch[1].trim().toLocaleLowerCase('en-US');
            if (!section) {
                throw new Error(
                    `Native Excel automation refused: Zone.Identifier is malformed: ${workbookPath}`
                );
            }
            if (section === 'zonetransfer') zoneTransferSections += 1;
            continue;
        }
        const assignment = /^([A-Za-z][A-Za-z0-9._-]*)\s*=\s*(.*)$/.exec(line);
        if (!assignment || !section) {
            throw new Error(
                `Native Excel automation refused: Zone.Identifier is malformed: ${workbookPath}`
            );
        }
        if (
            section === 'zonetransfer'
            && assignment[1].toLocaleLowerCase('en-US') === 'zoneid'
        ) {
            if (!/^[0-4]$/.test(assignment[2])) {
                throw new Error(
                    `Native Excel automation refused: Zone.Identifier has an invalid ZoneId: ${workbookPath}`
                );
            }
            zoneIds.push(Number(assignment[2]));
        }
    }
    if (zoneTransferSections !== 1 || zoneIds.length !== 1) {
        throw new Error(
            `Native Excel automation refused: Zone.Identifier is missing or ambiguous: ${workbookPath}`
        );
    }
    return zoneIds[0];
}

async function assertSafeZoneIdentifierForNativeAutomation(
    workbookPath: string
): Promise<void> {
    if (process.platform !== 'win32') return;
    const streamPath = `${workbookPath}:Zone.Identifier`;
    let handle: Awaited<ReturnType<typeof fs.open>>;
    try {
        handle = await fs.open(streamPath, 'r');
    } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'ENOENT') return;
        throw new Error(
            `Native Excel automation refused: Zone.Identifier could not be verified: ${workbookPath}`,
            { cause: error }
        );
    }
    try {
        const chunks: Buffer[] = [];
        let total = 0;
        let position = 0;
        while (true) {
            const chunk = Buffer.allocUnsafe(Math.min(8192, MAX_ZONE_IDENTIFIER_BYTES + 1 - total));
            const { bytesRead } = await handle.read(chunk, 0, chunk.length, position);
            if (bytesRead === 0) break;
            chunks.push(chunk.subarray(0, bytesRead));
            total += bytesRead;
            position += bytesRead;
            if (total > MAX_ZONE_IDENTIFIER_BYTES) {
                throw new Error(
                    `Native Excel automation refused: Zone.Identifier exceeds ${MAX_ZONE_IDENTIFIER_BYTES} bytes: ${workbookPath}`
                );
            }
        }
        const zoneId = parseZoneIdentifier(Buffer.concat(chunks, total), workbookPath);
        if (zoneId === 3 || zoneId === 4) {
            throw new Error(
                `Native Excel automation refused: the workbook is marked as Internet or Restricted Zone (ZoneId=${zoneId}). Trust and unblock it explicitly before editing: ${workbookPath}`
            );
        }
    } finally {
        await handle.close();
    }
}

export async function getFileSha256(filePath: string): Promise<string> {
    const hash = createHash('sha256');
    for await (const chunk of createReadStream(filePath)) {
        hash.update(chunk);
    }
    return hash.digest('hex');
}

function parseNativeEditResult(
    output: string,
    workbookExtension: '.xlsx' | '.xlsm'
): NativeExcelEditResult {
    const lines = output
        .split(/\r?\n/)
        .map(line => line.trim())
        .filter(Boolean);
    const finalLine = lines.at(-1);
    if (!finalLine) {
        throw new Error('Native edit script returned no result.');
    }

    let parsed: unknown;
    try {
        parsed = JSON.parse(finalLine);
    } catch {
        throw new Error(
            `Native edit script returned invalid JSON: ${finalLine.slice(0, 500)}`
        );
    }
    if (
        typeof parsed !== 'object' ||
        parsed === null ||
        !('ok' in parsed) ||
        (parsed as { ok?: unknown }).ok !== true
    ) {
        throw new Error('Native edit script did not report success.');
    }

    const { backupPath, workbookSha256 } = parsed as {
        backupPath?: unknown;
        workbookSha256?: unknown;
    };
    if (
        typeof backupPath !== 'string' ||
        !isAbsolute(backupPath) ||
        extname(backupPath).toLowerCase() !== workbookExtension
    ) {
        throw new Error('Native edit script returned an invalid backup path.');
    }
    if (
        typeof workbookSha256 !== 'string' ||
        !/^[0-9a-f]{64}$/.test(workbookSha256)
    ) {
        throw new Error('Native edit script returned an invalid workbook hash.');
    }
    return { backupPath, workbookSha256 };
}

function isRecord(value: unknown): value is Record<string, unknown> {
    return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function assertAllowedKeys(
    value: Record<string, unknown>,
    allowedKeys: readonly string[],
    label: string
): void {
    const allowed = new Set(allowedKeys);
    if (Object.keys(value).some(key => !allowed.has(key))) {
        throw new Error(`${label} contains an unknown property.`);
    }
}

function validateSheetName(value: unknown): asserts value is string {
    if (
        typeof value !== 'string' ||
        !value.trim() ||
        value.length > 31
    ) {
        throw new Error(
            'Native edit operation contains an invalid worksheet name.'
        );
    }
}

function validateExcelRow(value: unknown): asserts value is number {
    if (
        !Number.isInteger(value) ||
        (value as number) < 1 ||
        (value as number) > 1_048_576
    ) {
        throw new Error('Native edit operation contains an invalid row.');
    }
}

function validateExcelColumn(value: unknown): asserts value is number {
    if (
        !Number.isInteger(value) ||
        (value as number) < 1 ||
        (value as number) > 16_384
    ) {
        throw new Error('Native edit operation contains an invalid column.');
    }
}

function validateColor(value: unknown): asserts value is string {
    if (typeof value !== 'string' || !/^#[0-9a-f]{6}$/.test(value)) {
        throw new Error(
            'Native conditional-formatting color must be lowercase #rrggbb.'
        );
    }
}

function validateShortText(value: unknown, label: string): asserts value is string {
    if (
        typeof value !== 'string' ||
        value.length === 0 ||
        value.length > 255 ||
        value.includes('\0')
    ) {
        throw new Error(`${label} must contain 1-255 characters without NUL.`);
    }
}

function validateConditionalFormattingRule(
    value: unknown
): asserts value is NativeExcelConditionalFormattingRule {
    if (!isRecord(value) || typeof value.type !== 'string') {
        throw new Error('Native conditional-formatting rule is invalid.');
    }
    switch (value.type) {
        case 'cellIs':
            assertAllowedKeys(
                value,
                [
                    'type',
                    'operator',
                    'operand',
                    'fillColor',
                    'fontColor',
                    'bold',
                ],
                'Native cellIs rule'
            );
            if (
                !['greaterThan', 'lessThan', 'equal'].includes(
                    String(value.operator)
                )
            ) {
                throw new Error('Native cellIs operator is invalid.');
            }
            if (typeof value.operand === 'string') {
                validateShortText(value.operand, 'Native cellIs operand');
                if (value.operand.startsWith('=')) {
                    throw new Error(
                        'Native cellIs operand cannot be an Excel formula.'
                    );
                }
            } else if (
                typeof value.operand !== 'number' ||
                !Number.isFinite(value.operand)
            ) {
                throw new Error('Native cellIs operand is invalid.');
            }
            validateColor(value.fillColor);
            validateColor(value.fontColor);
            if (value.bold !== true) {
                throw new Error('Native cellIs bold style is invalid.');
            }
            return;

        case 'containsText':
            assertAllowedKeys(
                value,
                [
                    'type',
                    'text',
                    'fillColor',
                    'fontColor',
                    'bold',
                ],
                'Native containsText rule'
            );
            validateShortText(value.text, 'Native containsText value');
            validateColor(value.fillColor);
            validateColor(value.fontColor);
            if (value.bold !== true) {
                throw new Error('Native containsText bold style is invalid.');
            }
            return;

        case 'colorScale':
            assertAllowedKeys(
                value,
                ['type', 'colors'],
                'Native colorScale rule'
            );
            if (!Array.isArray(value.colors) || value.colors.length !== 3) {
                throw new Error(
                    'Native colorScale rule requires exactly three colors.'
                );
            }
            value.colors.forEach(validateColor);
            return;

        case 'dataBar':
            assertAllowedKeys(
                value,
                ['type', 'color'],
                'Native dataBar rule'
            );
            validateColor(value.color);
            return;

        case 'iconSet':
            assertAllowedKeys(
                value,
                ['type', 'iconSet', 'thresholds'],
                'Native iconSet rule'
            );
            if (
                value.iconSet !== '3TrafficLights1' ||
                !Array.isArray(value.thresholds) ||
                value.thresholds.length !== 2 ||
                value.thresholds[0] !== 33 ||
                value.thresholds[1] !== 67
            ) {
                throw new Error('Native iconSet rule is invalid.');
            }
            return;

        default:
            throw new Error(
                `Unsupported native conditional-formatting rule: ${value.type}`
            );
    }
}

function excelColumnNumber(letters: string): number {
    let number = 0;
    for (const letter of letters) {
        number = number * 26 + letter.charCodeAt(0) - 64;
    }
    return number;
}

function normalizeRangeRef(value: unknown): string {
    if (typeof value !== 'string') {
        throw new Error('Native workbook range is invalid.');
    }
    const normalized = value.trim().replace(/\$/g, '').toUpperCase();
    const match = normalized.match(
        /^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$/
    );
    if (!match) {
        throw new Error('Native workbook range is invalid.');
    }
    const columns = [match[1], match[3]].filter(Boolean);
    const rows = [match[2], match[4]].filter(Boolean).map(Number);
    if (
        columns.some(column => excelColumnNumber(column) > 16_384) ||
        rows.some(row => row > 1_048_576)
    ) {
        throw new Error(
            'Native workbook range exceeds Excel limits.'
        );
    }
    const startColumn = excelColumnNumber(match[1]);
    const endColumn = excelColumnNumber(match[3] || match[1]);
    const startRow = Number(match[2]);
    const endRow = Number(match[4] || match[2]);
    if (startColumn > endColumn || startRow > endRow) {
        throw new Error('Native workbook range must run from top-left to bottom-right.');
    }
    return normalized;
}

function boundedWorkbookObjectRangeCellCount(
    rangeRef: string,
    label: string
): number {
    const normalized = normalizeRangeRef(rangeRef);
    const match = normalized.match(
        /^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$/
    );
    if (!match) throw new Error(`${label} is not a valid local A1 range.`);
    const startColumn = excelColumnNumber(match[1]);
    const endColumn = excelColumnNumber(match[3] || match[1]);
    const startRow = Number(match[2]);
    const endRow = Number(match[4] || match[2]);
    const cellCount = (endColumn - startColumn + 1) * (endRow - startRow + 1);
    if (cellCount > MAX_WORKBOOK_OBJECT_RANGE_CELLS) {
        throw new Error(
            `${label} exceeds the ${MAX_WORKBOOK_OBJECT_RANGE_CELLS}-cell native Excel safety limit.`
        );
    }
    return cellCount;
}

function workbookObjectRangeCellCost(
    tableOrChart: SheetTableData | SheetChartData,
    label: string
): number {
    if ('rangeRef' in tableOrChart) {
        return boundedWorkbookObjectRangeCellCount(
            tableOrChart.rangeRef,
            `${label}.rangeRef`
        );
    }
    let total = 0;
    if (tableOrChart.sourceRangeRef) {
        total += boundedWorkbookObjectRangeCellCount(
            tableOrChart.sourceRangeRef,
            `${label}.sourceRangeRef`
        );
    }
    for (const [seriesIndex, series] of (tableOrChart.series ?? []).entries()) {
        for (const property of [
            'nameRange',
            'categoryRange',
            'valuesRange',
            'xValuesRange',
            'bubbleSizesRange',
        ] as const) {
            const rangeRef = series[property];
            if (rangeRef) {
                total += boundedWorkbookObjectRangeCellCount(
                    rangeRef,
                    `${label}.series[${seriesIndex}].${property}`
                );
            }
        }
    }
    return total;
}

function validateObjectId(value: unknown, label: string): asserts value is string {
    if (
        typeof value !== 'string' ||
        value.length < 1 ||
        value.length > 512 ||
        /[\0-\x1f\x7f]/.test(value)
    ) {
        throw new Error(`${label} must contain 1-512 characters without controls.`);
    }
}

function validateObjectName(
    value: unknown,
    label: string,
    identifierOnly = false
): asserts value is string {
    if (
        typeof value !== 'string' ||
        value.length < 1 ||
        value.length > 255 ||
        /[\0-\x1f\x7f]/.test(value) ||
        (identifierOnly && !/^[A-Za-z_][A-Za-z0-9_]{0,254}$/.test(value))
    ) {
        throw new Error(`${label} is invalid.`);
    }
}

function normalizeTableName(value: unknown, label: string): string {
    if (typeof value !== 'string' || /[\0-\x1f\x7f]/.test(value)) {
        throw new Error(`${label} is invalid.`);
    }
    const name = normalizeExcelTableName(value);
    if (!isValidExcelTableName(name)) throw new Error(`${label} is invalid.`);
    return name;
}

function validateBoolean(value: unknown, label: string): asserts value is boolean {
    if (typeof value !== 'boolean') {
        throw new Error(`${label} must be a boolean.`);
    }
}

function validateFiniteNumber(
    value: unknown,
    label: string,
    minimum: number,
    maximum: number,
    integer = false
): asserts value is number {
    if (
        typeof value !== 'number' ||
        !Number.isFinite(value) ||
        value < minimum ||
        value > maximum ||
        (integer && !Number.isInteger(value))
    ) {
        throw new Error(`${label} is outside the supported range.`);
    }
}

function normalizeTableData(value: unknown): SheetTableData {
    if (!isRecord(value)) {
        throw new Error('Native table definition must be an object.');
    }
    assertAllowedKeys(
        value,
        ['id', 'name', 'displayName', 'rangeRef', 'headerRow', 'totalsRow', 'style'],
        'Native table definition'
    );
    validateObjectId(value.id, 'Native table id');
    const name = normalizeTableName(value.name, 'Native table name');
    const displayName = normalizeTableName(value.displayName, 'Native table display name');
    validateBoolean(value.headerRow, 'Native table headerRow');
    validateBoolean(value.totalsRow, 'Native table totalsRow');
    if (!isRecord(value.style)) {
        throw new Error('Native table style must be an object.');
    }
    assertAllowedKeys(
        value.style,
        ['name', 'showFirstColumn', 'showLastColumn', 'showRowStripes', 'showColumnStripes'],
        'Native table style'
    );
    if (
        typeof value.style.name !== 'string' ||
        !/^TableStyle(?:Light(?:[1-9]|1\d|2[01])|Medium(?:[1-9]|1\d|2[0-8])|Dark(?:[1-9]|1[01]))$/.test(value.style.name)
    ) {
        throw new Error('Native table style name is invalid.');
    }
    for (const key of [
        'showFirstColumn',
        'showLastColumn',
        'showRowStripes',
        'showColumnStripes',
    ] as const) {
        validateBoolean(value.style[key], `Native table style ${key}`);
    }
    const rangeRef = normalizeRangeRef(value.rangeRef);
    const rangeMatch = /^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$/.exec(rangeRef);
    const startRow = Number(rangeMatch?.[2]);
    const endRow = Number(rangeMatch?.[4] ?? rangeMatch?.[2]);
    if (endRow - startRow + 1 < minimumExcelTableRangeRows(value.totalsRow)) {
        throw new Error('Native table range does not contain enough rows for its header, data and optional totals row.');
    }
    return {
        ...(value as unknown as SheetTableData),
        name,
        displayName,
        rangeRef,
    };
}

function normalizeChartAxis(value: unknown, label: string): SheetChartAxisData {
    if (!isRecord(value)) {
        throw new Error(`${label} must be an object.`);
    }
    assertAllowedKeys(
        value,
        [
            'visible', 'title', 'minimumScale', 'maximumScale', 'majorUnit',
            'minorUnit', 'logarithmic', 'reverseOrder', 'numberFormat',
            'majorGridlines', 'minorGridlines',
        ],
        label
    );
    for (const key of [
        'visible', 'logarithmic', 'reverseOrder', 'majorGridlines', 'minorGridlines',
    ] as const) {
        if (value[key] !== undefined) validateBoolean(value[key], `${label}.${key}`);
    }
    if (value.title !== undefined) {
        if (typeof value.title !== 'string' || value.title.length > 1000 || value.title.includes('\0')) {
            throw new Error(`${label}.title is invalid.`);
        }
    }
    if (value.numberFormat !== undefined) {
        if (typeof value.numberFormat !== 'string' || value.numberFormat.length > 255 || value.numberFormat.includes('\0')) {
            throw new Error(`${label}.numberFormat is invalid.`);
        }
    }
    for (const key of ['minimumScale', 'maximumScale'] as const) {
        if (value[key] !== undefined && value[key] !== null) {
            validateFiniteNumber(value[key], `${label}.${key}`, -1e307, 1e307);
        }
    }
    for (const key of ['majorUnit', 'minorUnit'] as const) {
        if (value[key] !== undefined && value[key] !== null) {
            validateFiniteNumber(value[key], `${label}.${key}`, Number.MIN_VALUE, 1e307);
        }
    }
    if (
        typeof value.minimumScale === 'number' &&
        typeof value.maximumScale === 'number' &&
        value.minimumScale >= value.maximumScale
    ) {
        throw new Error(`${label} minimumScale must be below maximumScale.`);
    }
    if (value.logarithmic === true && (
        (typeof value.minimumScale === 'number' && value.minimumScale <= 0)
        || (typeof value.maximumScale === 'number' && value.maximumScale <= 0)
    )) {
        throw new Error(`${label} logarithmic scale cannot use a non-positive bound.`);
    }
    return value as SheetChartAxisData;
}

function normalizeChartSeries(value: unknown, index: number): SheetChartSeriesData {
    const label = `Native chart series ${index + 1}`;
    if (!isRecord(value)) throw new Error(`${label} must be an object.`);
    assertAllowedKeys(
        value,
        [
            'id', 'name', 'nameRange', 'categoryRange', 'valuesRange',
            'xValuesRange', 'bubbleSizesRange', 'chartType', 'axisGroup',
            'color', 'lineColor', 'lineWidth', 'dashStyle', 'markerStyle',
            'markerSize', 'smooth', 'visible', 'dataLabels',
        ],
        label
    );
    validateObjectId(value.id, `${label}.id`);
    if (value.name !== undefined && value.nameRange !== undefined) {
        throw new Error(`${label} cannot define both name and nameRange.`);
    }
    if (value.categoryRange !== undefined && value.xValuesRange !== undefined) {
        throw new Error(`${label} cannot define both categoryRange and xValuesRange.`);
    }
    if (value.name !== undefined) {
        validateObjectName(value.name, `${label}.name`);
        if (/^[=+\-@]/.test(value.name.trim())) {
            throw new Error(`${label}.name cannot be an Excel formula.`);
        }
    }
    const normalized = { ...value } as Record<string, unknown>;
    for (const key of [
        'nameRange', 'categoryRange', 'valuesRange', 'xValuesRange', 'bubbleSizesRange',
    ] as const) {
        if (key === 'valuesRange' || value[key] !== undefined) {
            normalized[key] = normalizeRangeRef(value[key]);
        }
    }
    if (
        typeof normalized.nameRange === 'string' &&
        boundedWorkbookObjectRangeCellCount(normalized.nameRange, `${label}.nameRange`) !== 1
    ) {
        throw new Error(`${label}.nameRange must identify exactly one cell.`);
    }
    if (value.chartType !== undefined) {
        validateFiniteNumber(value.chartType, `${label}.chartType`, -10_000, 10_000, true);
        if (!VALID_CHART_TYPES.has(value.chartType)) throw new Error(`${label}.chartType is unknown.`);
        if (!isChartTypeCreatable(value.chartType)) {
            throw new Error(`${label}.chartType is not permitted for offline native creation.`);
        }
    }
    if (value.axisGroup !== undefined && !['primary', 'secondary'].includes(String(value.axisGroup))) {
        throw new Error(`${label}.axisGroup is invalid.`);
    }
    for (const key of ['color', 'lineColor'] as const) {
        if (value[key] !== undefined) validateColor(value[key]);
    }
    if (value.lineWidth !== undefined) validateFiniteNumber(value.lineWidth, `${label}.lineWidth`, 0.1, 20);
    if (value.markerSize !== undefined) validateFiniteNumber(value.markerSize, `${label}.markerSize`, 2, 72, true);
    if (value.dashStyle !== undefined && !['solid', 'dash', 'dot', 'dashDot'].includes(String(value.dashStyle))) {
        throw new Error(`${label}.dashStyle is invalid.`);
    }
    if (value.markerStyle !== undefined && ![
        'automatic', 'circle', 'dash', 'diamond', 'dot', 'none', 'picture',
        'plus', 'square', 'star', 'triangle', 'x',
    ].includes(String(value.markerStyle))) throw new Error(`${label}.markerStyle is invalid.`);
    for (const key of ['smooth', 'visible'] as const) {
        if (value[key] !== undefined) validateBoolean(value[key], `${label}.${key}`);
    }
    if (value.dataLabels !== undefined) {
        if (!isRecord(value.dataLabels)) throw new Error(`${label}.dataLabels must be an object.`);
        assertAllowedKeys(
            value.dataLabels,
            ['showValue', 'showCategoryName', 'showSeriesName', 'showPercentage', 'showBubbleSize', 'position'],
            `${label}.dataLabels`
        );
		if (!chartDataLabelsHaveExplicitShowOption(value.dataLabels as SheetChartSeriesData['dataLabels'])) {
			throw new Error(`${label}.dataLabels must explicitly define at least one show option.`);
		}
        for (const key of ['showValue', 'showCategoryName', 'showSeriesName', 'showPercentage', 'showBubbleSize'] as const) {
            if (value.dataLabels[key] !== undefined) validateBoolean(value.dataLabels[key], `${label}.dataLabels.${key}`);
        }
		if (
			value.dataLabels.position !== undefined
			&& !chartDataLabelsHaveEnabledShowOption(value.dataLabels as SheetChartSeriesData['dataLabels'])
		) {
			throw new Error(`${label}.dataLabels.position requires at least one enabled show option.`);
		}
        if (value.dataLabels.position !== undefined && ![
            'above', 'below', 'bestFit', 'center', 'insideBase', 'insideEnd',
            'left', 'outsideEnd', 'right',
        ].includes(String(value.dataLabels.position))) {
            throw new Error(`${label}.dataLabels.position is invalid.`);
        }
    }
    return normalized as unknown as SheetChartSeriesData;
}

function normalizeChartData(value: unknown): SheetChartData {
    if (!isRecord(value)) throw new Error('Native chart definition must be an object.');
    assertAllowedKeys(
        value,
        [
            'id', 'name', 'chartType', 'sourceRangeRef', 'plotBy', 'anchor',
            'title', 'legend', 'categoryAxis', 'valueAxis',
            'secondaryCategoryAxis', 'secondaryValueAxis', 'series', 'style',
            'roundedCorners', 'gapWidth', 'overlap', 'alternativeText',
        ],
        'Native chart definition'
    );
    validateObjectId(value.id, 'Native chart id');
    validateObjectName(value.name, 'Native chart name');
    validateFiniteNumber(value.chartType, 'Native chart type', -10_000, 10_000, true);
    if (!VALID_CHART_TYPES.has(value.chartType)) throw new Error('Native chart type is unknown.');
    if (!isChartTypeCreatable(value.chartType)) {
        throw new Error('Native chart type is not permitted for offline native creation.');
    }
    if (!['columns', 'rows'].includes(String(value.plotBy))) throw new Error('Native chart plotBy is invalid.');
    if (!isRecord(value.anchor)) throw new Error('Native chart anchor must be an object.');
    assertAllowedKeys(value.anchor, ['left', 'top', 'width', 'height'], 'Native chart anchor');
    validateFiniteNumber(value.anchor.left, 'Native chart left', 0, 10_000_000);
    validateFiniteNumber(value.anchor.top, 'Native chart top', 0, 10_000_000);
    validateFiniteNumber(value.anchor.width, 'Native chart width', 20, 1_000_000);
    validateFiniteNumber(value.anchor.height, 'Native chart height', 20, 1_000_000);

    const normalized = { ...value } as Record<string, unknown>;
    if (value.sourceRangeRef !== undefined && value.series !== undefined) {
        throw new Error('Native chart cannot define both sourceRangeRef and series.');
    }
    if (value.sourceRangeRef !== undefined) normalized.sourceRangeRef = normalizeRangeRef(value.sourceRangeRef);
    if (value.title !== undefined) {
        if (!isRecord(value.title)) throw new Error('Native chart title must be an object.');
        assertAllowedKeys(value.title, ['visible', 'text'], 'Native chart title');
        validateBoolean(value.title.visible, 'Native chart title.visible');
        if (typeof value.title.text !== 'string' || value.title.text.length > 1000 || value.title.text.includes('\0')) {
            throw new Error('Native chart title.text is invalid.');
        }
    }
    if (value.legend !== undefined) {
        if (!isRecord(value.legend)) throw new Error('Native chart legend must be an object.');
        assertAllowedKeys(value.legend, ['visible', 'position'], 'Native chart legend');
        validateBoolean(value.legend.visible, 'Native chart legend.visible');
        if (!['bottom', 'corner', 'custom', 'left', 'right', 'top'].includes(String(value.legend.position))) {
            throw new Error('Native chart legend.position is invalid.');
        }
    }
    for (const key of ['categoryAxis', 'valueAxis', 'secondaryCategoryAxis', 'secondaryValueAxis'] as const) {
        if (value[key] !== undefined) normalized[key] = normalizeChartAxis(value[key], `Native chart ${key}`);
    }
    if (value.series !== undefined) {
        if (!Array.isArray(value.series) || value.series.length > MAX_CHART_SERIES) {
            throw new Error(`Native chart series must contain at most ${MAX_CHART_SERIES} entries.`);
        }
        const ids = new Set<string>();
        const normalizedSeries = value.series.map((series, index) => {
            const result = normalizeChartSeries(series, index);
            if (ids.has(result.id)) throw new Error('Native chart series ids must be unique.');
            ids.add(result.id);
            return result;
        });
        if (normalizedSeries.length > 0) normalized.series = normalizedSeries;
        else delete normalized.series;
    }
    if (!normalized.sourceRangeRef && (!Array.isArray(normalized.series) || normalized.series.length === 0)) {
        throw new Error('Native chart requires a source range or at least one series.');
    }
    const requestedChartType = value.chartType as number;
    let series = (normalized.series as SheetChartSeriesData[] | undefined) ?? [];
    const requestedSeriesChartTypes = series.map(item => item.chartType ?? requestedChartType);
	if (!chartSeriesTypesCanCoexist(requestedSeriesChartTypes)) {
		throw new Error('Native charts cannot mix bubble and non-bubble series because Excel silently promotes every series to bubble.');
	}
    const canonicalChartType = canonicalChartTypeForSeries(
		requestedChartType,
		requestedSeriesChartTypes,
	);
    if (canonicalChartType === -4152 && requestedChartType !== -4152) {
        series = series.map(item => ({
            ...item,
            chartType: item.chartType ?? requestedChartType,
        }));
        normalized.series = series;
    }
	normalized.chartType = canonicalChartType;
    const chartType = normalized.chartType as number;
    const seriesChartTypes = series.map(item => item.chartType ?? chartType);
    const axisDefinitions = [
        normalized.categoryAxis,
        normalized.valueAxis,
        normalized.secondaryCategoryAxis,
        normalized.secondaryValueAxis,
    ];
    if (!chartTypeSupportsAxes(chartType) && axisDefinitions.some(axis => axis !== undefined)) {
        throw new Error('Native chart axes are not supported by this chart type.');
    }
    if (chartType === -4152) {
        if (
            series.length === 0 ||
            series.some(item =>
                typeof item.chartType !== 'number' ||
                item.chartType === -4152 ||
                !isChartTypeCreatable(item.chartType)
            ) ||
            new Set(series.map(item => item.chartType)).size < 2
        ) {
            throw new Error('Native custom combo charts require at least two explicit series with distinct concrete chartTypes.');
        }
    }
    const categoryScaleKeys = [
        'minimumScale', 'maximumScale', 'majorUnit', 'minorUnit', 'logarithmic',
    ] as const;
	for (const key of ['categoryAxis', 'secondaryCategoryAxis'] as const) {
            const axis = normalized[key] as SheetChartAxisData | undefined;
		const axisGroup = key === 'categoryAxis' ? 'primary' : 'secondary';
		if (
			!chartAxisGroupSupportsCategoryScale(chartType, series, axisGroup)
			&& axis
			&& categoryScaleKeys.some(property => axis[property] !== undefined)
		) {
                throw new Error(`Native chart ${key} numeric scale settings require a scatter or bubble chart.`);
            }
    }
	for (const [index, item] of series.entries()) {
		const effectiveSeriesType = item.chartType ?? chartType;
		if (item.bubbleSizesRange && !chartSeriesSupportsBubbleSizes(effectiveSeriesType)) {
			throw new Error(`Native chart series ${index + 1} bubbleSizesRange requires a bubble chart type.`);
		}
		if (item.smooth !== undefined && !chartSeriesSupportsSmooth(effectiveSeriesType)) {
			throw new Error(`Native chart series ${index + 1} smooth requires a line or scatter chart type.`);
		}
		if (item.dataLabels?.showBubbleSize === true && !chartSeriesSupportsBubbleSizes(effectiveSeriesType)) {
			throw new Error(`Native chart series ${index + 1} showBubbleSize requires a bubble chart type.`);
		}
		if (
			item.dataLabels?.showPercentage === true
			&& !chartSeriesSupportsPercentageDataLabels(effectiveSeriesType)
		) {
			throw new Error(`Native chart series ${index + 1} showPercentage requires a pie or doughnut chart type.`);
		}
		if (
			item.dataLabels?.position !== undefined
			&& !chartSeriesSupportsDataLabelPosition(
				effectiveSeriesType,
				item.dataLabels.position,
			)
		) {
			throw new Error(`Native chart series ${index + 1} data label position is not supported by this chart type.`);
		}
	}
    for (const key of ['secondaryCategoryAxis', 'secondaryValueAxis'] as const) {
        const axis = normalized[key] as SheetChartAxisData | undefined;
        if (axis && (axis.majorGridlines !== undefined || axis.minorGridlines !== undefined)) {
            throw new Error(`Native chart ${key} cannot define gridlines.`);
        }
    }
    if (
        !chartTypeSupportsSecondaryAxes(chartType, seriesChartTypes) &&
        (normalized.secondaryCategoryAxis !== undefined || normalized.secondaryValueAxis !== undefined)
    ) {
        throw new Error('Native chart secondary axes are not supported by this chart type.');
    }
    if (series.some(item => (
        item.axisGroup === 'secondary'
        && !chartTypeSupportsSecondaryAxes(item.chartType ?? chartType)
    ))) {
        throw new Error('Native chart series cannot use a secondary axis with this chart type.');
    }
    if (value.style !== undefined) validateFiniteNumber(value.style, 'Native chart style', 1, 48, true);
    if (value.roundedCorners !== undefined) validateBoolean(value.roundedCorners, 'Native chart roundedCorners');
    if (value.gapWidth !== undefined) {
        validateFiniteNumber(value.gapWidth, 'Native chart gapWidth', 0, 500, true);
        if (!chartTypeSupportsGapWidth(chartType, seriesChartTypes)) {
            throw new Error('Native chart gapWidth is not supported by this chart type.');
        }
    }
    if (value.overlap !== undefined) {
        validateFiniteNumber(value.overlap, 'Native chart overlap', -100, 100, true);
        if (!chartTypeSupportsOverlap(chartType, seriesChartTypes)) {
            throw new Error('Native chart overlap is not supported by this chart type.');
        }
    }
    if (value.alternativeText !== undefined) {
        if (typeof value.alternativeText !== 'string' || value.alternativeText.length > 1000 || value.alternativeText.includes('\0')) {
            throw new Error('Native chart alternativeText is invalid.');
        }
    }
    return normalized as unknown as SheetChartData;
}

export function validateNativeEdits(
    workbookPath: string,
    operations: NativeExcelEditOperation[]
): NativeExcelEditOperation[] {
    if (process.platform !== 'win32') {
        throw new Error('Native Excel editing requires Windows.');
    }
    const workbookExtension = extname(workbookPath).toLowerCase();
    if (!['.xlsx', '.xlsm'].includes(workbookExtension)) {
        throw new Error('Native editing is limited to .xlsx and .xlsm files.');
    }
    if (
        !Array.isArray(operations) ||
        operations.length === 0 ||
        operations.length > MAX_NATIVE_EDIT_OPERATIONS
    ) {
        throw new Error(
            `Native edit batch must contain 1-${MAX_NATIVE_EDIT_OPERATIONS} operations.`
        );
    }
    const conditionalAddsBySheet = new Map<string, number>();
    const workbookObjectOperationsBySheet = new Map<string, number>();
    let workbookObjectRangeCells = 0;
    const chargeWorkbookObjectRanges = (
        value: SheetTableData | SheetChartData,
        label: string
    ): void => {
        workbookObjectRangeCells += workbookObjectRangeCellCost(value, label);
        if (
            workbookObjectRangeCells
            > MAX_WORKBOOK_OBJECT_TRANSACTION_RANGE_CELLS
        ) {
            throw new Error(
                `Native workbook-object ranges exceed the ${MAX_WORKBOOK_OBJECT_TRANSACTION_RANGE_CELLS}-cell transaction safety budget.`
            );
        }
    };
    const normalizedOperations = operations.map(operation => {
        if (!isRecord(operation)) {
            throw new Error('Native edit operation must be an object.');
        }
        validateSheetName(operation.sheetName);
        const kind = operation.kind ?? 'cell';

        if (kind === 'cell') {
            validateExcelRow(operation.row);
            validateExcelColumn(operation.column);
            if (!operation.value && !operation.style) {
                throw new Error(
                    'Native edit operation must change a value or style.'
                );
            }
            if (operation.value) {
                if (!isRecord(operation.value)) {
                    throw new Error(
                        'Native edit operation contains an invalid value.'
                    );
                }
                const { kind: valueKind, value } = operation.value;
                if (
                    !['blank', 'formula', 'number', 'text'].includes(
                        String(valueKind)
                    ) ||
                    (valueKind === 'blank' && value !== undefined) ||
                    (valueKind === 'number' &&
                        (typeof value !== 'number' ||
                            !Number.isFinite(value))) ||
                    ((valueKind === 'formula' || valueKind === 'text') &&
                        typeof value !== 'string') ||
                    (valueKind === 'formula' &&
                        !(value as string).startsWith('='))
                ) {
                    throw new Error(
                        'Native edit operation contains an invalid value.'
                    );
                }
            }
            return operation;
        }

        if (kind === 'columnWidth') {
            assertAllowedKeys(
                operation,
                ['kind', 'sheetName', 'column', 'widthPx'],
                'Native column-width operation'
            );
            validateExcelColumn(operation.column);
            if (
                typeof operation.widthPx !== 'number' ||
                !Number.isFinite(operation.widthPx) ||
                operation.widthPx <= 5 ||
                operation.widthPx > 1_790
            ) {
                throw new Error('Native column width is invalid.');
            }
            return operation;
        }

        if (kind === 'rowHeight') {
            assertAllowedKeys(
                operation,
                ['kind', 'sheetName', 'row', 'heightPx'],
                'Native row-height operation'
            );
            validateExcelRow(operation.row);
            if (
                typeof operation.heightPx !== 'number' ||
                !Number.isFinite(operation.heightPx) ||
                operation.heightPx <= 0 ||
                operation.heightPx > 546
            ) {
                throw new Error('Native row height is invalid.');
            }
            return operation;
        }

        if (kind === 'addConditionalFormatting') {
            assertAllowedKeys(
                operation,
                ['kind', 'sheetName', 'rangeRef', 'rule'],
                'Native conditional-formatting operation'
            );
            const rangeRef = normalizeRangeRef(operation.rangeRef);
            validateConditionalFormattingRule(operation.rule);
            const addCount =
                (conditionalAddsBySheet.get(operation.sheetName) ?? 0) + 1;
            if (
                addCount >
                MAX_CONDITIONAL_FORMATTING_ADDS_PER_SHEET
            ) {
                throw new Error(
                    `Native edit cannot append more than ${MAX_CONDITIONAL_FORMATTING_ADDS_PER_SHEET} conditional-formatting rules per worksheet.`
                );
            }
            conditionalAddsBySheet.set(operation.sheetName, addCount);
            return { ...operation, rangeRef };
        }

        if (kind === 'clearConditionalFormatting') {
            assertAllowedKeys(
                operation,
                ['kind', 'sheetName'],
                'Native clear-conditional-formatting operation'
            );
            return operation;
        }

        if ([
            'createTable', 'updateTable', 'deleteTable',
            'createChart', 'updateChart', 'deleteChart',
        ].includes(kind)) {
            const count =
                (workbookObjectOperationsBySheet.get(operation.sheetName) ?? 0) + 1;
            if (count > MAX_WORKBOOK_OBJECTS_PER_SHEET) {
                throw new Error(
                    `Native edit cannot apply more than ${MAX_WORKBOOK_OBJECTS_PER_SHEET} workbook-object operations per worksheet.`
                );
            }
            workbookObjectOperationsBySheet.set(operation.sheetName, count);
        }

        if (kind === 'createTable') {
            assertAllowedKeys(operation, ['kind', 'sheetName', 'table'], 'Native createTable operation');
            const table = normalizeTableData(operation.table);
			if (!table.headerRow) {
				throw new Error('Native table creation with headerRow=false is disabled because Excel can move worksheet cells.');
			}
			if (table.totalsRow) {
				throw new Error('Native table creation with totalsRow=true is disabled because Excel moves worksheet cells and rewrites formula references.');
			}
            chargeWorkbookObjectRanges(table, 'Native createTable operation.table');
            return { ...operation, table };
        }
        if (kind === 'updateTable') {
            assertAllowedKeys(operation, ['kind', 'sheetName', 'name', 'table'], 'Native updateTable operation');
            const name = normalizeTableName(operation.name, 'Native current table name');
            const table = normalizeTableData(operation.table);
            chargeWorkbookObjectRanges(table, 'Native updateTable operation.table');
            return { ...operation, name, table };
        }
        if (kind === 'deleteTable') {
            assertAllowedKeys(operation, ['kind', 'sheetName', 'name'], 'Native deleteTable operation');
            const name = normalizeTableName(operation.name, 'Native current table name');
            return { ...operation, name };
        }
        if (kind === 'createChart') {
            assertAllowedKeys(operation, ['kind', 'sheetName', 'chart'], 'Native createChart operation');
            const chart = normalizeChartData(operation.chart);
			if (chart.legend?.position === 'custom') {
				throw new Error('Native createChart cannot create a custom legend layout; custom only preserves an existing manual Excel layout.');
			}
            chargeWorkbookObjectRanges(chart, 'Native createChart operation.chart');
            return { ...operation, chart };
        }
        if (kind === 'updateChart') {
            assertAllowedKeys(
                operation,
				['kind', 'sheetName', 'name', 'chart', 'preserveAnchor', 'preserveSeries', 'allowSeriesFormattingChange'],
                'Native updateChart operation'
            );
            validateObjectName(operation.name, 'Native current chart name');
            if (operation.preserveAnchor !== undefined) {
                validateBoolean(operation.preserveAnchor, 'Native updateChart preserveAnchor');
            }
            if (operation.preserveSeries !== undefined) {
                validateBoolean(operation.preserveSeries, 'Native updateChart preserveSeries');
            }
			if (operation.allowSeriesFormattingChange !== undefined) {
				validateBoolean(operation.allowSeriesFormattingChange, 'Native updateChart allowSeriesFormattingChange');
				if (operation.allowSeriesFormattingChange !== true || operation.preserveSeries !== true) {
					throw new Error('Native updateChart allowSeriesFormattingChange requires preserveSeries=true.');
				}
			}
            const chart = normalizeChartData(operation.chart);
            chargeWorkbookObjectRanges(chart, 'Native updateChart operation.chart');
            return { ...operation, chart };
        }
        if (kind === 'deleteChart') {
            assertAllowedKeys(operation, ['kind', 'sheetName', 'name'], 'Native deleteChart operation');
            validateObjectName(operation.name, 'Native current chart name');
            return operation;
        }

        throw new Error(`Unsupported native edit operation: ${String(kind)}`);
    });
    const tableRangesBySheet = new Map<string, Array<{ name: string; rangeRef: string }>>();
    const tableNames = new Set<string>();
    for (const operation of normalizedOperations) {
        if (operation.kind !== 'createTable' && operation.kind !== 'updateTable') continue;
        const table = operation.table as SheetTableData;
        for (const nameKey of new Set([
            excelTableNameComparisonKey(table.name),
            excelTableNameComparisonKey(table.displayName),
        ])) {
            if (tableNames.has(nameKey)) {
                throw new Error(`Native table name ${table.name} is requested more than once in the workbook.`);
            }
            tableNames.add(nameKey);
        }
        const sheetKey = operation.sheetName.toLocaleLowerCase('en-US');
        const ranges = tableRangesBySheet.get(sheetKey) ?? [];
        const overlap = ranges.find(candidate => simpleA1RangesOverlap(candidate.rangeRef, table.rangeRef));
        if (overlap) {
            throw new Error(`Native tables ${overlap.name} and ${table.name} overlap on worksheet ${operation.sheetName}.`);
        }
        ranges.push({ name: table.name, rangeRef: table.rangeRef });
        tableRangesBySheet.set(sheetKey, ranges);
    }
    return normalizedOperations;
}

export async function applyNativeExcelEdits(
    workbookPath: string,
    operations: NativeExcelEditOperation[],
    expectedWorkbookSha256: string
): Promise<NativeExcelEditResult> {
    const canonicalWorkbookPath = await fs.realpath(workbookPath);
    await assertNoReparsePointChain(canonicalWorkbookPath);
    await assertSafeZoneIdentifierForNativeAutomation(canonicalWorkbookPath);
    await assertOoxmlPackageUnsignedForMutation(canonicalWorkbookPath);
    await assertOoxmlHasNoXlmMacroSheetsForAutomation(canonicalWorkbookPath);
    const validatedOperations = validateNativeEdits(
        canonicalWorkbookPath,
        operations
    );
    if (!/^[0-9a-f]{64}$/.test(expectedWorkbookSha256)) {
        throw new Error('Native edit requires the SHA-256 baseline loaded by the editor.');
    }
    const payload: NativeExcelEditPayload = {
        version: 2,
        transactionId: randomUUID().toLowerCase(),
        expectedWorkbookSha256,
        operations: validatedOperations,
    };
    const serialized = JSON.stringify(payload);
    if (Buffer.byteLength(serialized, 'utf8') > MAX_NATIVE_EDIT_PAYLOAD_BYTES) {
        throw new Error('Native edit payload exceeds the 4 MiB safety limit.');
    }

    const payloadPath = join(
        tmpdir(),
        `excel-ai-vba-native-edits-${randomUUID()}.json`
    );
    const scriptPath = join(
        getExtensionRoot(),
        'scripts',
        'office-ai-apply-edits.ps1'
    );
    await fs.writeFile(payloadPath, serialized, {
        encoding: 'utf8',
        flag: 'wx',
    });

    try {
        const output = await new Promise<string>((resolvePromise, reject) => {
            let timedOut = false;
            let timeout: NodeJS.Timeout | undefined;
            const child = execFile(
                getPowerShellPath(),
                [
                    '-NoLogo',
                    '-NoProfile',
                    '-NonInteractive',
                    '-ExecutionPolicy',
                    'Bypass',
                    '-File',
                    scriptPath,
                    '-WorkbookPath',
                    canonicalWorkbookPath,
                    '-OperationsPath',
                    payloadPath,
                ],
                {
                    windowsHide: true,
                    maxBuffer: 1024 * 1024,
                },
                async (error, stdout, stderr) => {
                    if (timeout) {
                        clearTimeout(timeout);
                        timeout = undefined;
                    }
                    if (error) {
                        const stdoutStr = String(stdout);
                        if (timedOut) {
                            const ownedExcelProcess = parseOwnedExcelProcessId(stdoutStr);
                            if (ownedExcelProcess) {
                                await terminateExactProcess(ownedExcelProcess);
                            }
                        }
                        reject(
                            new Error(
                                String(stderr || stdout || error.message).trim()
                            )
                        );
                        return;
                    }
                    resolvePromise(String(stdout));
                }
            );
            timeout = setTimeout(() => {
                if (child.exitCode !== null || child.signalCode !== null) return;
                timedOut = child.kill();
            }, NATIVE_EDIT_TIMEOUT_MS);
        });

        const workbookExtension = extname(canonicalWorkbookPath).toLowerCase() as
            | '.xlsx'
            | '.xlsm';
        const result = parseNativeEditResult(output, workbookExtension);
        const canonicalBackupPath = await fs.realpath(result.backupPath);
        await assertNoReparsePointChain(canonicalBackupPath);
        const expectedBackupDirectory = resolve(
            dirname(canonicalWorkbookPath),
            '.excel-ai-vba-backups'
        );
        if (
            resolve(dirname(canonicalBackupPath)).toLowerCase() !==
            expectedBackupDirectory.toLowerCase()
        ) {
            throw new Error(
                'Native edit script returned a backup outside the managed directory.'
            );
        }
        if ((await getFileSha256(canonicalBackupPath)) !== expectedWorkbookSha256) {
            throw new Error('Native edit backup does not match the original workbook.');
        }
        if (
            (await getFileSha256(canonicalWorkbookPath)) !==
            result.workbookSha256
        ) {
            throw new Error(
                'Native edit result hash does not match the saved workbook.'
            );
        }
        return {
            backupPath: canonicalBackupPath,
            workbookSha256: result.workbookSha256,
        };
    } finally {
        await fs.rm(payloadPath, { force: true });
    }
}
