import { Handler } from '@/common/handler';
import type { NativeExcelCellEdit } from '@/common/nativeExcelEdits';
import { parseSafeExternalUri } from '@/common/webviewUri';
import { applyNativeExcelEdits } from '@/provider/nativeExcelBridge';
import {
    emitFileOfficeOpen,
    emitVirtualOfficeOpen,
    getEmbeddedSpreadsheetReadOnlyState,
    isVirtualUri,
    requiresNativeExcelForEditing,
    supportsNativeMacroEditing,
    type EmbeddedSpreadsheetReadOnlyReason,
} from '@/provider/handlers/officeContent';
import { createHash, randomUUID } from 'crypto';
import { basename, extname, join, parse } from 'path';
import * as vscode from 'vscode';
import { Uri, workspace } from 'vscode';

const fileSaveTimes: Record<string, number> = {};
const INTERNAL_SAVE_CHANGE_WINDOW_MS = 1500;
const MACRO_WRITE_BLOCKED_MESSAGE =
    'Protection VBA : les fichiers .xlsm et .xls sont en lecture seule dans l’éditeur intégré afin de préserver leurs macros et données héritées. Ouvrez le fichier dans Microsoft Excel pour le modifier.';
const NATIVE_BINARY_WRITE_BLOCKED_MESSAGE =
    'La sauvegarde XLSM complète est bloquée pour préserver le projet VBA. Utilisez la sauvegarde native ciblée de l’extension.';
const SAVE_FORMATS: Record<string, { label: string; exts: string[] }> = {
    xlsx: { label: 'Excel Workbook', exts: ['xlsx'] },
    xlsm: { label: 'Excel Macro-Enabled Workbook', exts: ['xlsm'] },
    xls: { label: 'Excel 97-2003 Workbook', exts: ['xls'] },
    ods: { label: 'OpenDocument Spreadsheet', exts: ['ods'] },
    csv: { label: 'CSV (Comma delimited)', exts: ['csv'] },
    tsv: { label: 'TSV (Tab delimited)', exts: ['tsv'] },
};

export interface OfficeEditorSession {
    save(): Promise<void>;
    saveAs(destination: Uri): Promise<void>;
    revert(): Promise<void>;
    backup(): Promise<OfficeEditorBackup>;
}

export interface OfficeEditorBackup {
    sheets: unknown;
    sourceSha256: string;
}

export interface OfficeDocumentLifecycle {
    readonly restoredSheets?: unknown;
    readonly restoredSourceSha256?: string;
    markChanged(): void;
    registerSession(session: OfficeEditorSession): vscode.Disposable;
}

export function shouldSkipFileChange(uri: Uri): boolean {
    const lastSaveTime = fileSaveTimes[uri.toString()];
    return Boolean(
        lastSaveTime &&
        Date.now() - lastSaveTime < INTERNAL_SAVE_CHANGE_WINDOW_MS
    );
}

function setDirty(handler: Handler, uri: Uri, dirty: boolean): void {
    const fileName = basename(uri.fsPath);
    handler.panel.title = fileName;
    if (dirty) {
        void vscode.commands.executeCommand('workbench.action.keepEditor');
    }
}

function toBytes(content: unknown): Uint8Array {
    if (Array.isArray(content)) {
        return new Uint8Array(content);
    }
    if (typeof content === 'string') {
        return new TextEncoder().encode(content);
    }
    throw new Error('Invalid spreadsheet save payload.');
}

function notifyMacroWriteBlocked(handler: Handler): void {
    handler.emit('writeBlocked', {
        reason: 'macro-preservation',
        message: MACRO_WRITE_BLOCKED_MESSAGE,
    });
    void vscode.window.showWarningMessage(MACRO_WRITE_BLOCKED_MESSAGE);
}

function notifyNativeBinaryWriteBlocked(handler: Handler): void {
    handler.emit('writeBlocked', {
        reason: 'native-excel-editing',
        message: NATIVE_BINARY_WRITE_BLOCKED_MESSAGE,
    });
    void vscode.window.showWarningMessage(NATIVE_BINARY_WRITE_BLOCKED_MESSAGE);
}

export function handleCommonEvent(
    uri: Uri,
    handler: Handler,
    lifecycle?: OfficeDocumentLifecycle
): void {
    let nativeLoadGeneration: string | undefined;
    let sendQueue: Promise<void> = Promise.resolve();
    let restoredSheets = lifecycle?.restoredSheets;
    let restoredSourceSha256 = lifecycle?.restoredSourceSha256;
    let hasUnsavedChanges = restoredSheets !== undefined;
    let externalChangePrompt: Promise<void> | undefined;
    let pendingSave:
        | {
              promise: Promise<void>;
              resolve: () => void;
              reject: (error: Error) => void;
              timer: ReturnType<typeof setTimeout>;
          }
        | undefined;
    let pendingSaveAs:
        | {
              destination: Uri;
              promise: Promise<void>;
              resolve: () => void;
              reject: (error: Error) => void;
              timer: ReturnType<typeof setTimeout>;
          }
        | undefined;
    let pendingBackup:
        | {
              requestId: string;
              promise: Promise<OfficeEditorBackup>;
              resolve: (state: OfficeEditorBackup) => void;
              reject: (error: Error) => void;
              timer: ReturnType<typeof setTimeout>;
          }
        | undefined;
    let readOnlyReason: EmbeddedSpreadsheetReadOnlyReason | undefined =
        supportsNativeMacroEditing(uri)
            ? 'native-excel-editing'
            : requiresNativeExcelForEditing(uri)
              ? 'macro-preservation'
              : undefined;

    const refreshReadOnlyState = async () => {
        const state = await getEmbeddedSpreadsheetReadOnlyState(uri);
        readOnlyReason = state.readOnlyReason;
        return state;
    };

    const updateDirty = (dirty: boolean): void => {
        hasUnsavedChanges = dirty;
        setDirty(handler, uri, dirty);
        if (dirty) {
            lifecycle?.markChanged();
        }
    };

    const send = (force = false): Promise<void> => {
        const queuedSend = sendQueue.then(async () => {
            if (!force && shouldSkipFileChange(uri)) {
                return;
            }
            await refreshReadOnlyState();
            if (isVirtualUri(uri)) {
                nativeLoadGeneration = undefined;
                await emitVirtualOfficeOpen(handler, uri, {
                    backupSheets: restoredSheets,
                    backupSourceSha256: restoredSourceSha256,
                });
                restoredSheets = undefined;
                restoredSourceSha256 = undefined;
                return;
            }
            nativeLoadGeneration = supportsNativeMacroEditing(uri)
                ? randomUUID()
                : undefined;
            await emitFileOfficeOpen(handler, uri, handler.panel.webview, {
                nativeLoadGeneration,
                backupSheets: restoredSheets,
                backupSourceSha256: restoredSourceSha256,
            });
            restoredSheets = undefined;
            restoredSourceSha256 = undefined;
        });
        sendQueue = queuedSend.catch(() => undefined);
        return queuedSend;
    };

    const failPendingSave = (error: Error): void => {
        for (const request of [pendingSave, pendingSaveAs]) {
            if (request) {
                clearTimeout(request.timer);
                request.reject(error);
            }
        }
        pendingSave = undefined;
        pendingSaveAs = undefined;
    };

    const completeSave = (): void => {
        const requests = [pendingSave, pendingSaveAs];
        pendingSave = undefined;
        pendingSaveAs = undefined;
        updateDirty(false);
        for (const request of requests) {
            if (request) {
                clearTimeout(request.timer);
                request.resolve();
            }
        }
    };

    const requestSave = (): Promise<void> => {
        if (pendingSave) {
            return pendingSave.promise;
        }
        if (pendingSaveAs) {
            return Promise.reject(new Error('A Save As operation is already running.'));
        }
        let resolveRequest!: () => void;
        let rejectRequest!: (error: Error) => void;
        const promise = new Promise<void>((resolvePromise, rejectPromise) => {
            resolveRequest = resolvePromise;
            rejectRequest = rejectPromise;
        });
        const request = {
            promise,
            resolve: resolveRequest,
            reject: rejectRequest,
            timer: setTimeout(() => {
                if (pendingSave === request) {
                    pendingSave = undefined;
                    rejectRequest(new Error('Timed out while waiting for the spreadsheet save.'));
                }
            }, 120_000),
        };
        pendingSave = request;
        handler.emit('requestSave');
        return promise;
    };

    const requestSaveAs = (destination: Uri): Promise<void> => {
        if (pendingSaveAs) {
            return pendingSaveAs.promise;
        }
        if (pendingSave) {
            return Promise.reject(new Error('A save operation is already running.'));
        }
        let resolveRequest!: () => void;
        let rejectRequest!: (error: Error) => void;
        const promise = new Promise<void>((resolvePromise, rejectPromise) => {
            resolveRequest = resolvePromise;
            rejectRequest = rejectPromise;
        });
        const request = {
            destination,
            promise,
            resolve: resolveRequest,
            reject: rejectRequest,
            timer: setTimeout(() => {
                if (pendingSaveAs === request) {
                    pendingSaveAs = undefined;
                    rejectRequest(new Error('Timed out while waiting for Save As.'));
                }
            }, 120_000),
        };
        pendingSaveAs = request;
        handler.emit('requestSaveAs', {
            ext: extname(destination.fsPath).replace(/^\./, '').toLowerCase(),
        });
        return promise;
    };

    const requestBackup = (): Promise<OfficeEditorBackup> => {
        if (pendingBackup) {
            return pendingBackup.promise;
        }
        const requestId = randomUUID();
        let resolveRequest!: (state: OfficeEditorBackup) => void;
        let rejectRequest!: (error: Error) => void;
        const promise = new Promise<OfficeEditorBackup>((resolvePromise, rejectPromise) => {
            resolveRequest = resolvePromise;
            rejectRequest = rejectPromise;
        });
        const request = {
            requestId,
            promise,
            resolve: resolveRequest,
            reject: rejectRequest,
            timer: setTimeout(() => {
                if (pendingBackup === request) {
                    pendingBackup = undefined;
                    rejectRequest(new Error('Timed out while backing up the spreadsheet.'));
                }
            }, 30_000),
        };
        pendingBackup = request;
        handler.emit('requestBackup', { requestId });
        return promise;
    };

    const executeForActiveDocument = async (command: string): Promise<void> => {
        const activeTab = vscode.window.tabGroups.activeTabGroup.activeTab;
        const input = activeTab?.input;
        if (
            !handler.panel.active ||
            !(input instanceof vscode.TabInputCustom) ||
            input.viewType !== 'excelAiVbaStudio.officeViewer' ||
            input.uri.toString() !== uri.toString()
        ) {
            throw new Error(
                'Save refused because this spreadsheet is no longer the active editor.'
            );
        }
        await vscode.commands.executeCommand(command);
    };

    const sessionRegistration = lifecycle?.registerSession({
        save: requestSave,
        saveAs: requestSaveAs,
        revert: async () => {
            failPendingSave(new Error('Save cancelled by revert.'));
            restoredSheets = undefined;
            restoredSourceSha256 = undefined;
            updateDirty(false);
            await send(true);
        },
        backup: requestBackup,
    });

    handler
        .on('editInVSCode', (full: boolean) => {
            const side = full ? vscode.ViewColumn.Active : vscode.ViewColumn.Beside;
            return vscode.commands.executeCommand('vscode.openWith', uri, 'default', side);
        })
        .on('init', () => send())
        .on('fileChange', () => {
            if (shouldSkipFileChange(uri)) {
                return;
            }
            if (!hasUnsavedChanges) {
                return send();
            }
            if (externalChangePrompt) {
                return externalChangePrompt;
            }
            const pendingPrompt = (async () => {
                const reload = 'Recharger le fichier';
                const keep = 'Conserver mes modifications';
                const choice = await vscode.window.showWarningMessage(
                    'Le fichier a changé sur le disque pendant que vous avez des modifications non sauvegardées.',
                    {
                        modal: true,
                        detail:
                            'Recharger abandonne les modifications de la grille. ' +
                            'Les conserver empêche tout rechargement automatique.',
                    },
                    reload,
                    keep
                );
                if (choice === reload) {
                    try {
                        await executeForActiveDocument(
                            'workbench.action.files.revert'
                        );
                    } catch (error) {
                        void vscode.window.showWarningMessage(
                            error instanceof Error
                                ? error.message
                                : String(error)
                        );
                    }
                }
            })();
            externalChangePrompt = pendingPrompt.finally(() => {
                externalChangePrompt = undefined;
            });
            return externalChangePrompt;
        })
        .on('change', () => {
            if (readOnlyReason === 'macro-preservation') {
                notifyMacroWriteBlocked(handler);
                return;
            }
            updateDirty(true);
        })
        .on('save', async (content: unknown) => {
            try {
                const state = await refreshReadOnlyState();
                if (state.readOnlyReason === 'macro-preservation') {
                    notifyMacroWriteBlocked(handler);
                    throw new Error(MACRO_WRITE_BLOCKED_MESSAGE);
                }
                if (state.readOnlyReason === 'native-excel-editing') {
                    notifyNativeBinaryWriteBlocked(handler);
                    throw new Error(NATIVE_BINARY_WRITE_BLOCKED_MESSAGE);
                }
                const bytes = toBytes(content);
                if (state.readOnly) {
                    throw new Error('Read-only spreadsheets must be saved to a new file.');
                }
                fileSaveTimes[uri.toString()] = Date.now();
                await workspace.fs.writeFile(uri, bytes);
                fileSaveTimes[uri.toString()] = Date.now();
                completeSave();
                handler.emit('saveDone', {
                    sourceSha256: createHash('sha256')
                        .update(bytes)
                        .digest('hex'),
                });
            } catch (error) {
                const saveError =
                    error instanceof Error ? error : new Error(String(error));
                failPendingSave(saveError);
                handler.emit('writeBlocked', {
                    message: saveError.message,
                });
                throw saveError;
            }
        })
        .on('saveAs', async (payload: { content: number[]; ext?: string }) => {
            try {
                const state = await refreshReadOnlyState();
                if (state.readOnlyReason === 'macro-preservation') {
                    notifyMacroWriteBlocked(handler);
                    throw new Error(MACRO_WRITE_BLOCKED_MESSAGE);
                }
                if (state.readOnlyReason === 'native-excel-editing') {
                    notifyNativeBinaryWriteBlocked(handler);
                    throw new Error(NATIVE_BINARY_WRITE_BLOCKED_MESSAGE);
                }
                const ext = (payload?.ext ?? 'xlsx').toLowerCase();
                const format = SAVE_FORMATS[ext];
                if (!format || !Array.isArray(payload?.content)) {
                    throw new Error(`Unsupported spreadsheet save format: ${ext}`);
                }

                const bytes = new Uint8Array(payload.content);
                const providerSaveAs = Boolean(pendingSaveAs);
                let target = pendingSaveAs?.destination;
                if (!target) {
                    const { dir, name } = parse(uri.fsPath);
                    const defaultFileName = `${name}.${ext}`;
                    const defaultUri = uri.scheme === 'file'
                        ? Uri.file(join(dir, defaultFileName))
                        : Uri.joinPath(uri, '..', defaultFileName);
                    target = await vscode.window.showSaveDialog({
                        defaultUri,
                        filters: { [format.label]: format.exts },
                    });
                }
                if (!target) {
                    const cancelled = new Error('Save As was cancelled.');
                    failPendingSave(cancelled);
                    handler.emit('writeBlocked', {
                        message: cancelled.message,
                    });
                    return;
                }

                fileSaveTimes[target.toString()] = Date.now();
                await workspace.fs.writeFile(target, bytes);
                fileSaveTimes[target.toString()] = Date.now();
                completeSave();
                handler.emit('saveDone', {
                    sourceSha256: createHash('sha256')
                        .update(bytes)
                        .digest('hex'),
                });
                if (!providerSaveAs) {
                    await vscode.commands.executeCommand(
                        'vscode.openWith',
                        target,
                        'excelAiVbaStudio.officeViewer'
                    );
                }
            } catch (error) {
                const saveError =
                    error instanceof Error ? error : new Error(String(error));
                failPendingSave(saveError);
                handler.emit('writeBlocked', {
                    message: saveError.message,
                });
                throw saveError;
            }
        })
        .on('saveNative', async (payload: {
            operations: NativeExcelCellEdit[];
            expectedWorkbookSha256: string;
            nativeLoadGeneration: string;
        }) => {
            try {
                const state = await refreshReadOnlyState();
                if (
                    state.readOnly ||
                    state.readOnlyReason !== 'native-excel-editing' ||
                    uri.scheme !== 'file'
                ) {
                    notifyNativeBinaryWriteBlocked(handler);
                    throw new Error(NATIVE_BINARY_WRITE_BLOCKED_MESSAGE);
                }

                if (
                    !nativeLoadGeneration ||
                    payload?.nativeLoadGeneration !== nativeLoadGeneration
                ) {
                    throw new Error(
                        'La sauvegarde XLSM est refusée : une version plus récente est en cours de chargement.'
                    );
                }

                fileSaveTimes[uri.toString()] = Date.now();
                await applyNativeExcelEdits(
                    uri.fsPath,
                    payload.operations,
                    payload.expectedWorkbookSha256
                );
                nativeLoadGeneration = undefined;
                fileSaveTimes[uri.toString()] = Date.now();
                completeSave();
                handler.emit('saveDone');
                await send(true);
            } catch (error) {
                const message =
                    error instanceof Error
                        ? error.message
                        : 'La sauvegarde native XLSM a échoué.';
                handler.emit('writeBlocked', {
                    reason: 'native-excel-editing',
                    message,
                });
                failPendingSave(
                    error instanceof Error ? error : new Error(message)
                );
                void vscode.window.showErrorMessage(message);
            }
        })
        .on('clean', () => {
            completeSave();
            handler.emit('saveDone');
        })
        .on('saveRejected', (payload?: { message?: string }) => {
            failPendingSave(
                new Error(payload?.message || 'Spreadsheet save was refused.')
            );
        })
        .on('requestHostSave', async () => {
            try {
                await executeForActiveDocument('workbench.action.files.save');
            } catch (error) {
                handler.emit('writeBlocked', {
                    message:
                        error instanceof Error
                            ? error.message
                            : String(error),
                });
            }
        })
        .on('requestHostSaveAs', async () => {
            try {
                await executeForActiveDocument('workbench.action.files.saveAs');
            } catch (error) {
                handler.emit('writeBlocked', {
                    message:
                        error instanceof Error
                            ? error.message
                            : String(error),
                });
            }
        })
        .on('backupState', (payload?: {
            requestId?: string;
            sheets?: unknown;
            sourceSha256?: string;
            error?: string;
        }) => {
            if (
                !pendingBackup ||
                payload?.requestId !== pendingBackup.requestId
            ) {
                return;
            }
            const request = pendingBackup;
            pendingBackup = undefined;
            clearTimeout(request.timer);
            if (payload?.error) {
                request.reject(new Error(payload.error));
            } else if (
                !/^[0-9a-f]{64}$/.test(payload?.sourceSha256 ?? '')
            ) {
                request.reject(
                    new Error('Spreadsheet backup is missing its source hash.')
                );
            } else {
                request.resolve({
                    sheets: payload?.sheets,
                    sourceSha256: payload.sourceSha256!,
                });
            }
        })
        .on('openVbaDeveloper', () =>
            vscode.commands.executeCommand(
                'excelAiVbaStudio.openVbaDeveloper',
                uri
            )
        )
        .on('openExcel', () =>
            vscode.commands.executeCommand('excelAiVbaStudio.openExcel', uri)
        )
        .on('openVbe', () =>
            vscode.commands.executeCommand('excelAiVbaStudio.openVbe', uri)
        )
        .on('exportWorkbookContext', () =>
            vscode.commands.executeCommand(
                'excelAiVbaStudio.exportWorkbook',
                uri
            )
        )
        .on('openVbaExplorer', () =>
            vscode.commands.executeCommand(
                'excelAiVbaStudio.openVbaExplorer',
                uri
            )
        )
        .on('askCopilotAboutWorkbook', (request?: string) =>
            vscode.commands.executeCommand(
                'excelAiVbaStudio.askCopilotAboutWorkbook',
                {
                    resourceUri: uri,
                    request: typeof request === 'string' ? request.slice(0, 4000) : undefined,
                }
            )
        )
        .on('openExternal', (url: string) => {
            const externalUri = parseSafeExternalUri(url);
            if (externalUri) {
                return vscode.env.openExternal(externalUri);
            }
        })
        .on('dispose', () => {
            sessionRegistration?.dispose();
            failPendingSave(new Error('Spreadsheet editor was closed.'));
            if (pendingBackup) {
                clearTimeout(pendingBackup.timer);
                pendingBackup.reject(
                    new Error('Spreadsheet editor was closed before backup.')
                );
                pendingBackup = undefined;
            }
            delete fileSaveTimes[uri.toString()];
        });
}
