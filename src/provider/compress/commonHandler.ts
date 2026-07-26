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
import { basename, join, parse } from 'path';
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

export function shouldSkipFileChange(uri: Uri): boolean {
    const lastSaveTime = fileSaveTimes[uri.toString()];
    return Boolean(
        lastSaveTime &&
        Date.now() - lastSaveTime < INTERNAL_SAVE_CHANGE_WINDOW_MS
    );
}

function setDirty(handler: Handler, uri: Uri, dirty: boolean): void {
    const fileName = basename(uri.fsPath);
    handler.panel.title = dirty ? `● ${fileName}` : fileName;
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

export function handleCommonEvent(uri: Uri, handler: Handler): void {
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

    const send = async (force = false): Promise<void> => {
        if (!force && shouldSkipFileChange(uri)) {
            return;
        }
        await refreshReadOnlyState();
        if (isVirtualUri(uri)) {
            await emitVirtualOfficeOpen(handler, uri);
            return;
        }
        await emitFileOfficeOpen(handler, uri, handler.panel.webview);
    };

    handler
        .on('editInVSCode', (full: boolean) => {
            const side = full ? vscode.ViewColumn.Active : vscode.ViewColumn.Beside;
            return vscode.commands.executeCommand('vscode.openWith', uri, 'default', side);
        })
        .on('init', () => send())
        .on('fileChange', () => send())
        .on('change', () => {
            if (readOnlyReason === 'macro-preservation') {
                notifyMacroWriteBlocked(handler);
                return;
            }
            setDirty(handler, uri, true);
        })
        .on('save', async (content: unknown) => {
            const state = await refreshReadOnlyState();
            if (state.readOnlyReason === 'macro-preservation') {
                notifyMacroWriteBlocked(handler);
                return;
            }
            if (state.readOnlyReason === 'native-excel-editing') {
                notifyNativeBinaryWriteBlocked(handler);
                return;
            }
            const bytes = toBytes(content);
            if (state.readOnly) {
                handler.emit('saveAs', { content: [...bytes] });
                return;
            }
            fileSaveTimes[uri.toString()] = Date.now();
            await workspace.fs.writeFile(uri, bytes);
            fileSaveTimes[uri.toString()] = Date.now();
            setDirty(handler, uri, false);
            handler.emit('saveDone');
        })
        .on('saveAs', async (payload: { content: number[]; ext?: string }) => {
            const state = await refreshReadOnlyState();
            if (state.readOnlyReason === 'macro-preservation') {
                notifyMacroWriteBlocked(handler);
                return;
            }
            if (state.readOnlyReason === 'native-excel-editing') {
                notifyNativeBinaryWriteBlocked(handler);
                return;
            }
            const ext = (payload?.ext ?? 'xlsx').toLowerCase();
            const format = SAVE_FORMATS[ext];
            if (!format || !Array.isArray(payload?.content)) {
                throw new Error(`Unsupported spreadsheet save format: ${ext}`);
            }

            const bytes = new Uint8Array(payload.content);
            const { dir, name } = parse(uri.fsPath);
            const defaultFileName = `${name}.${ext}`;
            const defaultUri = uri.scheme === 'file'
                ? Uri.file(join(dir, defaultFileName))
                : Uri.joinPath(uri, '..', defaultFileName);
            const target = await vscode.window.showSaveDialog({
                defaultUri,
                filters: { [format.label]: format.exts },
            });
            if (!target) {
                return;
            }

            fileSaveTimes[target.toString()] = Date.now();
            await workspace.fs.writeFile(target, bytes);
            fileSaveTimes[target.toString()] = Date.now();
            setDirty(handler, uri, false);
            handler.emit('saveDone');
            await vscode.commands.executeCommand(
                'vscode.openWith',
                target,
                'excelAiVbaStudio.officeViewer'
            );
        })
        .on('saveNative', async (operations: NativeExcelCellEdit[]) => {
            try {
                const state = await refreshReadOnlyState();
                if (
                    state.readOnly ||
                    state.readOnlyReason !== 'native-excel-editing' ||
                    uri.scheme !== 'file'
                ) {
                    notifyNativeBinaryWriteBlocked(handler);
                    return;
                }

                fileSaveTimes[uri.toString()] = Date.now();
                await applyNativeExcelEdits(uri.fsPath, operations);
                fileSaveTimes[uri.toString()] = Date.now();
                setDirty(handler, uri, false);
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
                void vscode.window.showErrorMessage(message);
            }
        })
        .on('openVbaDeveloper', () =>
            vscode.commands.executeCommand(
                'excelAiVbaStudio.openVbaDeveloper',
                uri
            )
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
            delete fileSaveTimes[uri.toString()];
        });
}
