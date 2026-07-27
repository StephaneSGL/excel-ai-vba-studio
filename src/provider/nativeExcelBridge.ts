import { execFile } from 'child_process';
import { createHash, randomUUID } from 'crypto';
import { createReadStream, promises as fs } from 'fs';
import { tmpdir } from 'os';
import { dirname, extname, isAbsolute, join, resolve } from 'path';
import * as vscode from 'vscode';
import {
    type NativeExcelCellEdit,
    type NativeExcelEditPayload,
    type NativeExcelEditResult,
} from '@/common/nativeExcelEdits';
import { assertNoReparsePointChain } from '@/excelAiVbaStudio/security';

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

function validateNativeEdits(
    workbookPath: string,
    operations: NativeExcelCellEdit[]
): NativeExcelCellEdit[] {
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
    return operations;
}

export async function applyNativeExcelEdits(
    workbookPath: string,
    operations: NativeExcelCellEdit[],
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
