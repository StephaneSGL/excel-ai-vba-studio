import { execFile } from 'child_process';
import { createHash, randomUUID } from 'crypto';
import { createReadStream, promises as fs } from 'fs';
import { tmpdir } from 'os';
import { dirname, extname, isAbsolute, join, resolve } from 'path';
import * as vscode from 'vscode';
import {
    type NativeExcelConditionalFormattingRule,
    type NativeExcelEditOperation,
    type NativeExcelEditPayload,
    type NativeExcelEditResult,
} from '@/common/nativeExcelEdits';
import { assertNoReparsePointChain } from '@/excelAiVbaStudio/security';

const MAX_NATIVE_EDIT_OPERATIONS = 10_000;
const MAX_CONDITIONAL_FORMATTING_ADDS_PER_SHEET = 64;
const MAX_NATIVE_EDIT_PAYLOAD_BYTES = 4 * 1024 * 1024;
const NATIVE_EDIT_TIMEOUT_MS = 120_000;

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

async function terminateExactProcess(processId: number): Promise<void> {
    if (
        process.platform !== 'win32' ||
        !Number.isSafeInteger(processId) ||
        processId <= 0
    ) {
        return;
    }
    const taskkillPath = process.env.SystemRoot
        ? join(process.env.SystemRoot, 'System32', 'taskkill.exe')
        : 'taskkill.exe';
    await new Promise<void>(resolvePromise => {
        execFile(
            taskkillPath,
            ['/PID', String(processId), '/F'],
            { windowsHide: true },
            () => resolvePromise()
        );
    });
}

function parseOwnedExcelProcessId(output: string): number | undefined {
    const match = output.match(/(?:^|\r?\n)OWNED_EXCEL_PID\|(\d+)(?:\r?\n|$)/);
    if (!match) {
        return undefined;
    }
    const processId = Number.parseInt(match[1], 10);
    return Number.isSafeInteger(processId) && processId > 0
        ? processId
        : undefined;
}

export async function getFileSha256(filePath: string): Promise<string> {
    const hash = createHash('sha256');
    for await (const chunk of createReadStream(filePath)) {
        hash.update(chunk);
    }
    return hash.digest('hex');
}

function parseNativeEditResult(output: string): NativeExcelEditResult {
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
        extname(backupPath).toLowerCase() !== '.xlsm'
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
        throw new Error('Native conditional-formatting range is invalid.');
    }
    const normalized = value.toUpperCase();
    const match = normalized.match(
        /^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$/
    );
    if (!match) {
        throw new Error('Native conditional-formatting range is invalid.');
    }
    const columns = [match[1], match[3]].filter(Boolean);
    const rows = [match[2], match[4]].filter(Boolean).map(Number);
    if (
        columns.some(column => excelColumnNumber(column) > 16_384) ||
        rows.some(row => row > 1_048_576)
    ) {
        throw new Error(
            'Native conditional-formatting range exceeds Excel limits.'
        );
    }
    return normalized;
}

function validateNativeEdits(
    workbookPath: string,
    operations: NativeExcelEditOperation[]
): NativeExcelEditOperation[] {
    if (process.platform !== 'win32') {
        throw new Error('Native XLSM editing requires Windows.');
    }
    if (extname(workbookPath).toLowerCase() !== '.xlsm') {
        throw new Error('Native editing is currently limited to .xlsm files.');
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
    return operations.map(operation => {
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

        throw new Error(`Unsupported native edit operation: ${String(kind)}`);
    });
}

export async function applyNativeExcelEdits(
    workbookPath: string,
    operations: NativeExcelEditOperation[],
    expectedWorkbookSha256: string
): Promise<NativeExcelEditResult> {
    const canonicalWorkbookPath = await fs.realpath(workbookPath);
    await assertNoReparsePointChain(canonicalWorkbookPath);
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
            execFile(
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
                    timeout: NATIVE_EDIT_TIMEOUT_MS,
                    maxBuffer: 1024 * 1024,
                },
                async (error, stdout, stderr) => {
                    if (error) {
                        const stdoutStr = String(stdout);
                        const ownedExcelProcessId =
                            parseOwnedExcelProcessId(stdoutStr);
                        if (ownedExcelProcessId) {
                            await terminateExactProcess(ownedExcelProcessId);
                        } else if (
                            typeof error === 'object' &&
                            error !== null &&
                            'killed' in error &&
                            (error as unknown as Record<string, unknown>).killed
                        ) {
                            // Process was killed (timeout). The PID may
                            // still be extractable from partial stdout.
                            const pidFromStdout =
                                parseOwnedExcelProcessId(stdoutStr);
                            if (pidFromStdout) {
                                await terminateExactProcess(pidFromStdout);
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
        });

        const result = parseNativeEditResult(output);
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
