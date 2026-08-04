import { App, Button, ConfigProvider, Spin } from "antd";
import { useCallback, useEffect, useRef, useState } from "react";
import { handler, vscodeApi } from "../../util/vscode.ts";
import { isVscodeEditorDark, observeVscodeThemeChange } from "../../util/vscodeTheme.ts";
import { loadOfficeBuffer } from "../../util/loadOfficeContent.ts";
import { antThemeConfig } from '../../antThemeConfig.ts';
import './Excel.less';
import {
    buildFormattingSnapshot,
    hasFormattingChanged,
    MIN_VIEW_COLS,
    MIN_VIEW_ROWS,
} from "./excel_meta.ts";
import { detectCsvEncoding } from "./csvEncoding.ts";
import Spreadsheet, { type SheetData } from './x-spreadsheet/index';
import FindReplacePanel, { type FindReplacePanelHandle } from './FindReplacePanel';
import { parseSpreadsheetLink } from './excel_hyperlink';
import { initExcelLocale, t } from './excel_i18n';
import ExcelRibbon from './ExcelRibbon';
import {
    buildNativeExcelEditPlan,
    initializeNativeEditSheets,
} from './native_edit_diff';

initExcelLocale();

type EmbeddedReadOnlyReason =
    | 'macro-preservation'
    | 'native-excel-editing'
    | 'file-permissions'
    | 'package-signature'
    | 'package-signature-verification';

function isPackageSignatureReason(
    reason: EmbeddedReadOnlyReason | null
): boolean {
    return reason === 'package-signature'
        || reason === 'package-signature-verification';
}

function blocksSaveAs(reason: EmbeddedReadOnlyReason | null): boolean {
    return reason === 'macro-preservation'
        || reason === 'native-excel-editing'
        || isPackageSignatureReason(reason);
}

type ExcelViewState = { ri: number; ci: number; sheetIndex: number };

const EXCEL_VIEW_STATE_SUFFIX = '-excel-view';

let excelWriterPromise: Promise<typeof import('./excel_writer.ts')> | undefined;

function loadExcelWriter() {
    excelWriterPromise ??= import('./excel_writer.ts');
    return excelWriterPromise;
}

function getViewStateKey(documentCacheId: string): string {
    return `${documentCacheId}${EXCEL_VIEW_STATE_SUFFIX}`;
}

function loadViewState(documentCacheId: string): ExcelViewState | null {
    if (!documentCacheId) return null;
    const key = getViewStateKey(documentCacheId);
    try {
        const raw = localStorage.getItem(key);
        if (!raw) return null;
        return JSON.parse(raw) as ExcelViewState;
    } catch {
        return null;
    }
}

function saveViewState(documentCacheId: string, view: ExcelViewState) {
    if (!documentCacheId) return;
    const key = getViewStateKey(documentCacheId);
    try {
        localStorage.setItem(key, JSON.stringify(view));
    } catch {
        // ignore quota / private mode errors
    }
}

function restoreViewState(spreadSheet: Spreadsheet, saved: ExcelViewState) {
    const sheets = spreadSheet.getData();
    if (!sheets.length) return;
    const sheetIndex = Math.min(Math.max(0, saved.sheetIndex), sheets.length - 1);
    const sheet = sheets[sheetIndex];
    const maxRi = Math.max(0, (sheet.rows?.len ?? 1) - 1);
    const maxCi = Math.max(0, (sheet.cols?.len ?? 1) - 1);
    const ri = Math.min(Math.max(0, saved.ri), maxRi);
    const ci = Math.min(Math.max(0, saved.ci), maxCi);
    spreadSheet.scrollToCell(ri, ci, sheetIndex);
}

function isCsvLikeExt(ext: string): boolean {
    return /^(csv|tsv)$/i.test(ext.replace(/^\./, ''));
}

const WORKBOOK_OBJECT_OPERATION_KINDS = new Set([
    'createTable', 'updateTable', 'deleteTable',
    'createChart', 'updateChart', 'deleteChart',
]);

function isFindPanelTarget(target: EventTarget | null): boolean {
    return target instanceof Element && Boolean(target.closest('.frp-panel'));
}

function cloneSheets(sheets: SheetData[]): SheetData[] {
    return JSON.parse(JSON.stringify(sheets)) as SheetData[];
}

function isOoxmlSaveAsExt(ext: string): boolean {
    return /^(xlsx|xlsm)$/i.test(ext.replace(/^\./, ''));
}

function sheetsContainUnsafeNativeOoxmlObjects(sheets: readonly SheetData[]): boolean {
    return sheets.some(sheet => (sheet.tables?.length ?? 0) > 0
        || (sheet.charts?.length ?? 0) > 0
        || sheet.hasNativeChartParts === true);
}

async function sha256Buffer(buffer: ArrayBuffer): Promise<string> {
    const digest = await crypto.subtle.digest('SHA-256', buffer);
    return Array.from(new Uint8Array(digest), byte =>
        byte.toString(16).padStart(2, '0')
    ).join('');
}

function ExcelViewer() {
    const { message } = App.useApp();
    const [loading, setLoading] = useState(true)
    const [vscodeDark, setVscodeDark] = useState(isVscodeEditorDark)
    const [readOnly, setReadOnly] = useState(false)
    const [readOnlyReason, setReadOnlyReason] = useState<EmbeddedReadOnlyReason | null>(null)
    const [findPanel, setFindPanel] = useState<'find' | 'replace' | null>(null)
    const findPanelRef = useRef<'find' | 'replace' | null>(null)
    const findReplacePanelRef = useRef<FindReplacePanelHandle>(null)
    const [loadError, setLoadError] = useState<string | null>(null)
    const [activeSpreadsheet, setActiveSpreadsheet] = useState<Spreadsheet | null>(null)
    const [editorBusy, setEditorBusy] = useState(false)
    const extRef = useRef('')
    const documentCacheIdRef = useRef('')
    const readOnlyRef = useRef(false)
    const readOnlyReasonRef = useRef<EmbeddedReadOnlyReason | null>(null)
    const spreadSheetRef = useRef<Spreadsheet | null>(null)
    const csvEncodingRef = useRef<'utf8' | 'gbk'>('utf8')
    const csvDelimiterRef = useRef(',')
    const initialFormattingRef = useRef('')
    const initialSheetsRef = useRef<SheetData[]>([])
    const editorBusyRef = useRef(false)
    const nativeSnapshotRef = useRef<{
        expectedWorkbookSha256: string;
        nativeLoadGeneration: string;
    } | null>(null)
    const loadedWorkbookSha256Ref = useRef<string | null>(null)
    const openGenerationRef = useRef(0)
    const openQueueRef = useRef<Promise<void>>(Promise.resolve())
    const lastMacroBlockedNoticeRef = useRef(0)

    const notifyMacroWriteBlocked = useCallback(() => {
        const now = Date.now();
        if (now - lastMacroBlockedNoticeRef.current < 1000) {
            return;
        }
        lastMacroBlockedNoticeRef.current = now;
        message.warning({
            duration: 4,
            content: t('viewer.macroWriteBlocked'),
            className: 'excel-validation-error-message',
        });
    }, [message]);

    const notifyPackageSignatureWriteBlocked = useCallback(() => {
        message.warning({
            duration: 5,
            content: t('viewer.packageSignatureWriteBlocked'),
            className: 'excel-validation-error-message',
        });
    }, [message]);

    const setEditorBusyState = useCallback((busy: boolean) => {
        editorBusyRef.current = busy;
        setEditorBusy(busy);
        spreadSheetRef.current?.setMode(
            busy || readOnlyRef.current ? 'read' : 'edit'
        );
    }, []);

    useEffect(() => {
        findPanelRef.current = findPanel;
    }, [findPanel]);

    const themedDark = vscodeDark;

    useEffect(() => {
        return observeVscodeThemeChange(() => {
            setVscodeDark(isVscodeEditorDark());
            requestAnimationFrame(() => spreadSheetRef.current?.reRender());
        });
    }, []);

    useEffect(() => {
        document.body.classList.add('office-adaptive')
        document.body.classList.toggle('office-dark', themedDark)
        document.documentElement.style.colorScheme = themedDark ? 'dark' : 'light'
        return () => {
            document.body.classList.remove('office-adaptive')
            document.body.classList.remove('office-dark')
            document.documentElement.style.removeProperty('color-scheme')
        }
    }, [themedDark])

    useEffect(() => {
        spreadSheetRef.current?.reRender()
    }, [themedDark])

    const handleAutoFitColumns = useCallback(() => {
        const spreadSheet = spreadSheetRef.current;
        if (!spreadSheet || editorBusyRef.current) return;
        spreadSheet.autoFitColumns();
        if (!readOnlyRef.current) {
            spreadSheet.setSaveEnabled(true);
            handler.emit('change');
        }
        message.success({ duration: 1.5, content: t('viewer.autoFitDone'), className: 'excel-save-success-message' });
    }, [message]);

    const handleSaveAs = useCallback(() => {
        if (isPackageSignatureReason(readOnlyReasonRef.current)) {
            notifyPackageSignatureWriteBlocked();
            return;
        }
        if (blocksSaveAs(readOnlyReasonRef.current)) {
            notifyMacroWriteBlocked();
            return;
        }
        handler.emit('requestHostSaveAs');
    }, [notifyMacroWriteBlocked, notifyPackageSignatureWriteBlocked]);

    const handleSave = useCallback(async () => {
        const spreadSheet = spreadSheetRef.current;
        if (!spreadSheet) {
            handler.emit('saveRejected', {
                message: 'Spreadsheet is not ready to save.',
            });
            return;
        }
        if (editorBusyRef.current) {
            handler.emit('saveRejected', {
                message: 'Spreadsheet is busy. Wait for the current operation to finish.',
            });
            return;
        }
        if (readOnlyReasonRef.current === 'macro-preservation') {
            notifyMacroWriteBlocked();
            handler.emit('saveRejected', {
                message: t('viewer.macroWriteBlocked'),
            });
            return;
        }
        if (isPackageSignatureReason(readOnlyReasonRef.current)) {
            handler.emit('saveRejected', {
                message: t('viewer.packageSignatureWriteBlocked'),
            });
            return;
        }
        if (readOnlyRef.current) {
            handler.emit('saveRejected', {
                message: 'Read-only spreadsheets must be saved to a new file.',
            });
            return;
        }

        if (document.activeElement instanceof HTMLElement) {
            document.activeElement.blur();
        }
        const ext = extRef.current.replace(/^\./, '').toLowerCase();
        const sheets = spreadSheet.getData();
        const candidateNativePlan =
            readOnlyReasonRef.current === 'native-excel-editing' || ext === 'xlsx'
                ? buildNativeExcelEditPlan(initialSheetsRef.current, sheets)
                : null;
        const hasWorkbookObjectChanges = Boolean(
            candidateNativePlan?.operations.some(operation =>
                WORKBOOK_OBJECT_OPERATION_KINDS.has(operation.kind ?? 'cell')
            )
        );
        const hasExistingNativeObjects = initialSheetsRef.current.some(
            sheet => (sheet.tables?.length ?? 0) > 0
                || (sheet.charts?.length ?? 0) > 0
                || sheet.hasNativeChartParts === true
        );
        if (
            readOnlyReasonRef.current === 'native-excel-editing' ||
            (ext === 'xlsx' && (hasWorkbookObjectChanges || hasExistingNativeObjects))
        ) {
            const snapshot = nativeSnapshotRef.current;
            if (!snapshot) {
                const reloadMessage = t('viewer.nativeReloadRequired');
                message.warning({
                    duration: 4,
                    content: reloadMessage,
                    className: 'excel-validation-error-message',
                });
                spreadSheet.setSaveEnabled(true);
                handler.emit('saveRejected', { message: reloadMessage });
                return;
            }
            const plan = candidateNativePlan!;
            if (plan.unsupportedChanges.length > 0) {
                const unsupportedMessage = t(
                    'viewer.nativeUnsupportedChange',
                    plan.unsupportedChanges.slice(0, 3).join(', ')
                );
                message.warning({
                    duration: 6,
                    content: unsupportedMessage,
                    className: 'excel-validation-error-message',
                });
                spreadSheet.setSaveEnabled(true);
                handler.emit('saveRejected', {
                    message: unsupportedMessage,
                });
                return;
            }
            if (plan.operations.length === 0) {
                spreadSheet.setSaveEnabled(false);
                handler.emit('clean');
                return;
            }
            setEditorBusyState(true);
            handler.emit('saveNative', {
                operations: plan.operations,
                ...snapshot,
            });
            return;
        }

        setEditorBusyState(true);
        let export_xlsx: Awaited<ReturnType<typeof loadExcelWriter>>['export_xlsx'];
        try {
            ({ export_xlsx } = await loadExcelWriter());
        } catch (error) {
            setEditorBusyState(false);
            handler.emit('saveRejected', {
                message:
                    error instanceof Error
                        ? error.message
                        : String(error),
            });
            return;
        }
        const csvEncoding = csvEncodingRef.current;
        const csvDelimiter = csvDelimiterRef.current;

        if (ext !== 'xlsx' && ext !== 'xlsm' && hasFormattingChanged(initialFormattingRef.current, sheets)) {
            const formatMessage = t(
                'viewer.formatCannotPreserveContent',
                ext.toUpperCase()
            );
            message.warning({
                duration: 6,
                content: formatMessage,
                className: 'excel-validation-error-message',
            });
            setEditorBusyState(false);
            handler.emit('saveRejected', { message: formatMessage });
            return;
        }

        try {
            await export_xlsx(spreadSheet, extRef.current, csvEncoding, undefined, csvDelimiter);
            spreadSheet.setSaveEnabled(false);
        } catch (error) {
            console.error(`Failed to save Excel file: ${(error as Error).message}`);
            setEditorBusyState(false);
            handler.emit('saveRejected', {
                message:
                    error instanceof Error
                        ? error.message
                        : String(error),
            });
        }
    }, [
        message,
        notifyMacroWriteBlocked,
        setEditorBusyState,
    ]);

    useEffect(() => {
        const onKeyDown = (e: KeyboardEvent) => {
            if ((e.ctrlKey || e.metaKey) && e.code === 'KeyS') {
                e.preventDefault();
                if (readOnlyRef.current) {
                    void handleSaveAs();
                } else {
                    handler.emit('requestHostSave');
                }
                return;
            }
            if ((e.ctrlKey || e.metaKey) && e.code === 'KeyF') {
                e.preventDefault();
                if (findPanelRef.current) {
                    if (findPanelRef.current !== 'find') {
                        setFindPanel('find');
                    }
                    findReplacePanelRef.current?.focusFindInput(true);
                } else {
                    setFindPanel('find');
                }
                return;
            }
            if ((e.ctrlKey || e.metaKey) && e.code === 'KeyH') {
                e.preventDefault();
                setFindPanel(readOnlyRef.current ? 'find' : 'replace');
                return;
            }
            if ((e.ctrlKey || e.metaKey) && e.altKey && e.code === 'Digit0') {
                e.preventDefault();
                handleAutoFitColumns();
                return;
            }

            if (
                (
                    readOnlyReasonRef.current === 'macro-preservation' ||
                    isPackageSignatureReason(readOnlyReasonRef.current)
                )
                && !isFindPanelTarget(e.target)
            ) {
                const modifierEdit = (e.ctrlKey || e.metaKey)
                    && ['KeyV', 'KeyX', 'KeyY', 'KeyZ'].includes(e.code);
                const directEdit = !e.ctrlKey
                    && !e.metaKey
                    && !e.altKey
                    && (e.key.length === 1
                        || e.key === 'Backspace'
                        || e.key === 'Delete'
                        || e.key === 'F2');
                if (modifierEdit || directEdit) {
                    e.preventDefault();
                    e.stopPropagation();
                    if (isPackageSignatureReason(readOnlyReasonRef.current)) {
                        notifyPackageSignatureWriteBlocked();
                    } else {
                        notifyMacroWriteBlocked();
                    }
                }
            }
        };
        const blockInputEvent = (e: Event) => {
            if (
                (
                    readOnlyReasonRef.current !== 'macro-preservation' &&
                    !isPackageSignatureReason(readOnlyReasonRef.current)
                )
                || isFindPanelTarget(e.target)
            ) {
                return;
            }
            e.preventDefault();
            e.stopPropagation();
            if (isPackageSignatureReason(readOnlyReasonRef.current)) {
                notifyPackageSignatureWriteBlocked();
            } else {
                notifyMacroWriteBlocked();
            }
        };
        window.addEventListener('keydown', onKeyDown, true);
        document.addEventListener('beforeinput', blockInputEvent, true);
        document.addEventListener('paste', blockInputEvent, true);
        document.addEventListener('cut', blockInputEvent, true);
        document.addEventListener('drop', blockInputEvent, true);
        return () => {
            window.removeEventListener('keydown', onKeyDown, true);
            document.removeEventListener('beforeinput', blockInputEvent, true);
            document.removeEventListener('paste', blockInputEvent, true);
            document.removeEventListener('cut', blockInputEvent, true);
            document.removeEventListener('drop', blockInputEvent, true);
        };
    }, [
        handleAutoFitColumns,
        handleSave,
        handleSaveAs,
        notifyMacroWriteBlocked,
        notifyPackageSignatureWriteBlocked,
    ]);

    useEffect(() => {
        const container = document.getElementById('container');

        const initSpreadsheet = async (buffer: ArrayBuffer, payload: any) => {
            const fileReadOnly = payload.readOnly === true;
            const preserveSourceIntegrity =
                payload.readOnlyReason === 'macro-preservation'
                || payload.readOnlyReason === 'native-excel-editing'
                || payload.readOnlyReason === 'package-signature'
                || payload.readOnlyReason === 'package-signature-verification';
            if (payload.ext?.match(/csv/i)) {
                csvEncodingRef.current = detectCsvEncoding(buffer);
            }
            const { loadSheets } = await import('./excel_reader.ts');
            const { sheets, maxLength, maxCols, csvDelimiter } = await loadSheets(buffer, payload.ext);
            const restoredSheets = Array.isArray(payload.backupSheets)
                ? payload.backupSheets as SheetData[]
                : null;
            if (csvDelimiter) {
                csvDelimiterRef.current = csvDelimiter;
            }
            const viewRowLen = Math.max(maxLength ?? 0, MIN_VIEW_ROWS);
            const viewColLen = Math.max(maxCols ?? 0, MIN_VIEW_COLS);
            container.innerHTML = '';
            const spreadSheet = new Spreadsheet(container, {
                mode: fileReadOnly ? 'read' : 'edit',
                allowSaveAs: !preserveSourceIntegrity,
                showToolbar: false,
                showEditInVSCode: isCsvLikeExt(payload.ext ?? ''),
                row: { len: viewRowLen, height: 30 },
                col: { len: viewColLen },
                view: { height: () => Math.max(180, container.clientHeight || window.innerHeight - 130) },
            });
            spreadSheetRef.current = spreadSheet;
            setActiveSpreadsheet(spreadSheet);
            initialSheetsRef.current = initializeNativeEditSheets(
                spreadSheet,
                sheets,
                restoredSheets
            );
            if (restoredSheets) {
                spreadSheet.setSaveEnabled(true);
            }
            setEditorBusyState(false);
            setLoading(false);
            requestAnimationFrame(() => spreadSheet.resize());
            if (!fileReadOnly) {
                spreadSheet.on('save', () => handler.emit('requestHostSave'));
            }
            spreadSheet.on('save-as', () => { void handleSaveAs(); });
            spreadSheet.on('edit-in-vscode', () => { handler.emit('editInVSCode', true); });
            spreadSheet.on('find', () => {
                if (findPanelRef.current) {
                    if (findPanelRef.current !== 'find') {
                        setFindPanel('find');
                    }
                    findReplacePanelRef.current?.focusFindInput(true);
                } else {
                    setFindPanel('find');
                }
            });
            const persistView = () => {
                saveViewState(documentCacheIdRef.current, spreadSheet.getSelection());
            };
            spreadSheet.on('cell-selected', () => { persistView(); });
            spreadSheet.onSheetChange(() => { persistView(); });
            spreadSheet.onOpenLink((linkPayload) => {
                const parsed = parseSpreadsheetLink(linkPayload.link);
                if (parsed.type === 'internal') {
                    spreadSheet.followHyperlink(linkPayload);
                } else {
                    handler.emit('openExternal', parsed.url);
                }
            });
            spreadSheet.onProtectedCellDblClick(() => {
                message.info({ duration: 2, content: t('viewer.protectedCell'), className: 'excel-protected-cell-message' });
            });
            spreadSheet.onValidationError((errMessage) => {
                message.warning({ duration: 2, content: errMessage, className: 'excel-validation-error-message' });
            });
            spreadSheet.on('change', () => {
                if (!fileReadOnly) {
                    spreadSheet.setSaveEnabled(true);
                    handler.emit('change');
                }
            });
            const savedView = loadViewState(documentCacheIdRef.current);
            if (savedView) {
                requestAnimationFrame(() => {
                    requestAnimationFrame(() => { restoreViewState(spreadSheet, savedView); });
                });
            }
            const normalizedExt = (payload.ext ?? '').replace(/^\./, '').toLowerCase();
            initialFormattingRef.current = normalizedExt !== 'xlsx' && normalizedExt !== 'xlsm'
                ? buildFormattingSnapshot(sheets)
                : '';
        };

        handler.on("open", (payload) => {
            const openGeneration = ++openGenerationRef.current;
            const previousOpen = openQueueRef.current;
            setEditorBusyState(true);
            setLoading(true);
            setLoadError(null);
            nativeSnapshotRef.current = null;
            loadedWorkbookSha256Ref.current = null;
            extRef.current = payload.ext ?? '';
            documentCacheIdRef.current = payload.documentCacheId ?? '';
            const fileReadOnly = payload.readOnly === true;
            const reason = payload.readOnlyReason === 'macro-preservation'
                || payload.readOnlyReason === 'native-excel-editing'
                || payload.readOnlyReason === 'file-permissions'
                || payload.readOnlyReason === 'package-signature'
                || payload.readOnlyReason === 'package-signature-verification'
                ? payload.readOnlyReason as EmbeddedReadOnlyReason
                : null;
            readOnlyRef.current = fileReadOnly;
            readOnlyReasonRef.current = reason;
            setReadOnly(fileReadOnly);
            setReadOnlyReason(reason);
            const openTask = (async () => {
                try {
                    const buffer = await loadOfficeBuffer(payload);
                    const workbookSha256 = await sha256Buffer(buffer);
                    if (
                        Array.isArray(payload.backupSheets) &&
                        (
                            !/^[0-9a-f]{64}$/.test(
                                payload.backupSourceSha256 ?? ''
                            ) ||
                            payload.backupSourceSha256 !== workbookSha256
                        )
                    ) {
                        throw new Error(
                            'Spreadsheet recovery was stopped because the ' +
                            'source changed while the backup was opening.'
                        );
                    }
                    await previousOpen.catch(() => undefined);
                    if (openGeneration !== openGenerationRef.current) {
                        return;
                    }
                    await initSpreadsheet(buffer, payload);
                    loadedWorkbookSha256Ref.current = workbookSha256;
                    if (
                        openGeneration === openGenerationRef.current &&
                        ['xlsx', 'xlsm'].includes(
                            String(payload.ext ?? '').replace(/^\./, '').toLowerCase()
                        ) &&
                        typeof payload.nativeLoadGeneration === 'string'
                    ) {
                        nativeSnapshotRef.current = {
                            expectedWorkbookSha256: workbookSha256,
                            nativeLoadGeneration: payload.nativeLoadGeneration,
                        };
                    }
                } catch (e) {
                    await previousOpen.catch(() => undefined);
                    if (openGeneration !== openGenerationRef.current) {
                        return;
                    }
                    setEditorBusyState(false);
                    const msg = (e as Error).message || String(e);
                    console.error(`Failed to load Excel file: ${msg}`, e);
                    setLoadError(msg);
                    setLoading(false);
                }
            })();
            openQueueRef.current = openTask;
            return openTask;
        }).on("saveDone", (payload) => {
            const spreadSheet = spreadSheetRef.current;
            if (
                spreadSheet &&
                readOnlyReasonRef.current !== 'native-excel-editing'
            ) {
                initialSheetsRef.current = cloneSheets(spreadSheet.getData());
                spreadSheet.setSaveEnabled(false);
                setEditorBusyState(false);
            }
            if (/^[0-9a-f]{64}$/.test(payload?.sourceSha256 ?? '')) {
                loadedWorkbookSha256Ref.current = payload.sourceSha256;
            }
            message.success({
                duration: 2,
                content: t('viewer.saveSuccess'),
                className: 'excel-save-success-message',
            });
        }).on("writeBlocked", (payload) => {
            setEditorBusyState(false);
            spreadSheetRef.current?.setSaveEnabled(true);
            message.warning({
                duration: 4,
                content: payload?.message || t('viewer.macroWriteBlocked'),
                className: 'excel-validation-error-message',
            });
        }).on("requestSave", async () => {
            try {
                await handleSave();
            } catch (error) {
                setEditorBusyState(false);
                handler.emit('saveRejected', {
                    message:
                        error instanceof Error
                            ? error.message
                            : String(error),
                });
            }
        }).on("requestSaveAs", async (payload) => {
            let lockedForSaveAs = false;
            try {
                const spreadSheet = spreadSheetRef.current;
                const ext =
                    typeof payload?.ext === 'string'
                        ? payload.ext.replace(/^\./, '').toLowerCase()
                        : '';
                if (!spreadSheet || !ext) {
                    throw new Error('Spreadsheet is not ready for Save As.');
                }
                if (blocksSaveAs(readOnlyReasonRef.current)) {
                    throw new Error(
                        isPackageSignatureReason(readOnlyReasonRef.current)
                            ? t('viewer.packageSignatureWriteBlocked')
                            : t('viewer.macroWriteBlocked')
                    );
                }
                if (
                    isOoxmlSaveAsExt(ext)
                    && sheetsContainUnsafeNativeOoxmlObjects(spreadSheet.getData())
                ) {
                    throw new Error(
                        'Enregistrer sous au format XLSX/XLSM est désactivé pour préserver les tableaux et graphiques Excel natifs. '
                        + 'Choisissez CSV, TSV, ODS ou XLS pour un export aplati, ou créez la copie depuis Microsoft Excel ou l’Explorateur de fichiers.'
                    );
                }
                if (editorBusyRef.current) {
                    throw new Error(
                        'Spreadsheet is busy. Wait for the current operation to finish.'
                    );
                }
                setEditorBusyState(true);
                lockedForSaveAs = true;
                const { exportSaveAs } = await loadExcelWriter();
                await exportSaveAs(
                    spreadSheet,
                    ext,
                    csvEncodingRef.current,
                    csvDelimiterRef.current
                );
            } catch (error) {
                if (lockedForSaveAs) {
                    setEditorBusyState(false);
                }
                handler.emit('saveRejected', {
                    message:
                        error instanceof Error
                            ? error.message
                            : String(error),
                });
            }
        }).on("requestBackup", (payload) => {
            try {
                const spreadSheet = spreadSheetRef.current;
                if (!spreadSheet) {
                    throw new Error('Spreadsheet is not ready for backup.');
                }
                const sourceSha256 = loadedWorkbookSha256Ref.current;
                if (!sourceSha256) {
                    throw new Error(
                        'Spreadsheet source hash is not ready for backup.'
                    );
                }
                handler.emit('backupState', {
                    requestId: payload?.requestId,
                    sheets: cloneSheets(spreadSheet.getData()),
                    sourceSha256,
                });
            } catch (error) {
                handler.emit('backupState', {
                    requestId: payload?.requestId,
                    error:
                        error instanceof Error
                            ? error.message
                            : String(error),
                });
            }
        }).emit("init")

        let themeTimer: ReturnType<typeof setTimeout>;
        const themeObserver = new MutationObserver(() => {
            clearTimeout(themeTimer);
            themeTimer = setTimeout(() => spreadSheetRef.current?.reRender(), 120);
        });
        themeObserver.observe(document.head, { childList: true, subtree: true });

        return () => {
            spreadSheetRef.current = null;
            setActiveSpreadsheet(null);
            themeObserver.disconnect();
            clearTimeout(themeTimer);
        };
    }, [message, handleSave, handleSaveAs, setEditorBusyState])

    return (
        <div className='excel-viewer'>
            <Spin spinning={loading} fullscreen={true} />
            <ExcelRibbon
                spreadsheet={activeSpreadsheet}
                readOnly={readOnly || editorBusy}
                allowSaveAs={!blocksSaveAs(readOnlyReason)}
                showEditInVscode={isCsvLikeExt(extRef.current)}
                onAutoFitColumns={handleAutoFitColumns}
                onOpenExcel={() => handler.emit('openExcel')}
                onOpenVbe={() => handler.emit('openVbe')}
                onOpenVbaDeveloper={() => handler.emit('openVbaDeveloper')}
                onExportWorkbookContext={() => handler.emit('exportWorkbookContext')}
                onOpenVbaExplorer={() => handler.emit('openVbaExplorer')}
                onAskCopilotAboutWorkbook={(request) => handler.emit('askCopilotAboutWorkbook', request)}
            />
            {loadError && !loading && (
                <div className="excel-load-error">
                    <div className="excel-load-error-panel">
                        <svg className="excel-load-error-icon" width="44" height="44" viewBox="0 0 44 44" fill="none" aria-hidden>
                            <circle cx="22" cy="22" r="20" stroke="currentColor" strokeWidth="1.8" />
                            <path d="M22 13v12" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" />
                            <circle cx="22" cy="31" r="1.8" fill="currentColor" />
                        </svg>
                        <h2 className="excel-load-error-title">Failed to open file</h2>
                        <span className="excel-load-error-message">{loadError}</span>
                    </div>
                </div>
            )}
            {readOnly && !loading && !loadError && (
                <div
                    className={`excel-readonly-banner${
                        readOnlyReason === 'macro-preservation' ||
                        isPackageSignatureReason(readOnlyReason)
                            ? ' excel-macro-preservation-banner'
                            : ''
                    }`}
                    role="status"
                    aria-live="polite"
                >
                    <span>
                        {readOnlyReason === 'macro-preservation'
                            ? t('viewer.macroReadonlyBanner')
                            : readOnlyReason === 'package-signature'
                              ? t('viewer.packageSignatureReadonlyBanner')
                              : readOnlyReason === 'package-signature-verification'
                                ? t('viewer.packageSignatureVerificationReadonlyBanner')
                            : t('viewer.readonlyBanner')}
                    </span>
                    {readOnlyReason === 'macro-preservation' && (
                        <span className="excel-readonly-actions">
                            <Button
                                size="small"
                                onClick={() => handler.emit('openVbaDeveloper')}
                            >
                                {t('viewer.openVbaDeveloper')}
                            </Button>
                        </span>
                    )}
                </div>
            )}
            {findPanel && !loading && !loadError && (
                <FindReplacePanel
                    ref={findReplacePanelRef}
                    spreadSheet={activeSpreadsheet}
                    mode={findPanel}
                    onClose={() => setFindPanel(null)}
                    readOnly={readOnly || editorBusy}
                    onChanged={() => {
                        if (
                            !readOnlyRef.current &&
                            !editorBusyRef.current
                        ) {
                            spreadSheetRef.current?.setSaveEnabled(true);
                        }
                    }}
                />
            )}
            <div id='container'></div>
        </div>
    )
}

export default function Excel() {
    return (
        <ConfigProvider componentSize='small' theme={antThemeConfig}>
            <App className="excel-app" message={{ top: 16 }}>
                <ExcelViewer />
            </App>
        </ConfigProvider>
    );
}
