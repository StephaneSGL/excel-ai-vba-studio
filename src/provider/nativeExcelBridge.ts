import { execFile } from 'child_process';
import { randomUUID } from 'crypto';
import { promises as fs } from 'fs';
import { tmpdir } from 'os';
import { extname, join, resolve } from 'path';
import * as vscode from 'vscode';
import {
    type NativeExcelCellEdit,
    type NativeExcelEditPayload,
} from '@/common/nativeExcelEdits';

const MAX_NATIVE_EDIT_OPERATIONS = 10_000;
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

function validateNativeEdits(
    workbookPath: string,
    operations: NativeExcelCellEdit[]
): NativeExcelEditPayload {
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
    for (const operation of operations) {
        if (
            !operation ||
            typeof operation.sheetName !== 'string' ||
            !operation.sheetName.trim() ||
            !Number.isInteger(operation.row) ||
            operation.row < 1 ||
            operation.row > 1_048_576 ||
            !Number.isInteger(operation.column) ||
            operation.column < 1 ||
            operation.column > 16_384
        ) {
            throw new Error('Native edit operation contains an invalid cell target.');
        }
        if (!operation.value && !operation.style) {
            throw new Error('Native edit operation must change a value or style.');
        }
        if (operation.value) {
            const { kind, value } = operation.value;
            if (
                !['blank', 'formula', 'number', 'text'].includes(kind) ||
                (kind === 'blank' && value !== undefined) ||
                (kind === 'number' &&
                    (typeof value !== 'number' || !Number.isFinite(value))) ||
                ((kind === 'formula' || kind === 'text') &&
                    typeof value !== 'string') ||
                (kind === 'formula' && !(value as string).startsWith('='))
            ) {
                throw new Error('Native edit operation contains an invalid value.');
            }
        }
    }
    return { version: 1, operations };
}

export async function applyNativeExcelEdits(
    workbookPath: string,
    operations: NativeExcelCellEdit[]
): Promise<void> {
    const canonicalWorkbookPath = await fs.realpath(workbookPath);
    const payload = validateNativeEdits(canonicalWorkbookPath, operations);
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
        await new Promise<void>((resolvePromise, reject) => {
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
                    // Verify the script reported success
                    const stdoutStr = String(stdout);
                    if (!stdoutStr.includes('"ok":true')) {
                        const ownedExcelProcessId =
                            parseOwnedExcelProcessId(stdoutStr);
                        if (ownedExcelProcessId) {
                            await terminateExactProcess(ownedExcelProcessId);
                        }
                        reject(
                            new Error(
                                `Native edit script did not report success. Output: ${stdoutStr.slice(0, 500)}`
                            )
                        );
                        return;
                    }
                    resolvePromise();
                }
            );
        });
    } finally {
        await fs.rm(payloadPath, { force: true });
    }
}
