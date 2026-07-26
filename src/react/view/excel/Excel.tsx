import { App, Button, ConfigProvider, Modal, Radio, Spin } from "antd";
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
import { loadSheets } from "./excel_reader.ts";
import Spreadsheet, { type SheetData } from './x-spreadsheet/index';
import FindReplacePanel, { type FindReplacePanelHandle } from './FindReplacePanel';
import { parseSpreadsheetLink } from './excel_hyperlink';
import { initExcelLocale, t } from './excel_i18n';
import ExcelRibbon from './ExcelRibbon';
import { buildNativeExcelEditPlan } from './native_edit_diff';

initExcelLocale();

type EmbeddedReadOnlyReason =
    | 'macro-preservation'
    | 'native-excel-editing'
    | 'file-permissions';

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

function isFindPanelTarget(target: EventTarget | null): boolean {
    return target instanceof Element && Boolean(target.closest('.frp-panel'));
}

function cloneSheets(sheets: SheetData[]): SheetData[] {
    return JSON.parse(JSON.stringify(sheets)) as SheetData[];
}

function ExcelViewer() {
    const { message, modal } = App.useApp();
    const [loading, setLoading] = useState(true)
    const [vscodeDark, setVscodeDark] = useState(isVscodeEditorDark)
    const [readOnly, setReadOnly] = useState(false)
    const [readOnlyReason, setReadOnlyReason] = useState<EmbeddedReadOnlyReason | null>(null)
    const [findPanel, setFindPanel] = useState<'find' | 'replace' | null>(null)
    const findPanelRef = useRef<'find' | 'replace' | null>(null)
    const findReplacePanelRef = useRef<FindReplacePanelHandle>(null)
    const [loadError, setLoadError] = useState<string | null>(null)
    const [saveAsVisible, setSaveAsVisible] = useState(false)
    const [saveAsFormat, setSaveAsFormat] = useState('xlsx')
    const [activeSpreadsheet, setActiveSpreadsheet] = useState<Spreadsheet | null>(null)
    const extRef = useRef('')
    const documentCacheIdRef = useRef('')
    const readOnlyRef = useRef(false)
    const readOnlyReasonRef = useRef<EmbeddedReadOnlyReason | null>(null)
    const spreadSheetRef = useRef<Spreadsheet | null>(null)
    const csvEncodingRef = useRef<'utf8' | 'gbk'>('utf8')
    const csvDelimiterRef = useRef(',')
    const initialFormattingRef = useRef('')
    const initialSheetsRef = useRef<SheetData[]>([])
    const nativeSavePendingRef = useRef(false)
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
        if (!spreadSheet) return;
        spreadSheet.autoFitColumns();
        if (!readOnlyRef.current) {
            spreadSheet.setSaveEnabled(true);
            handler.emit('change');
        }
        message.success({ duration: 1.5, content: t('viewer.autoFitDone'), className: 'excel-save-success-message' });
    }, [message]);

    const handleSaveAs = useCallback(() => {
        if (
            readOnlyReasonRef.current === 'macro-preservation'
            || readOnlyReasonRef.current === 'native-excel-editing'
        ) {
            notifyMacroWriteBlocked();
            return;
        }
        setSaveAsVisible(true);
    }, [notifyMacroWriteBlocked]);

    const handleSave = useCallback(async () => {
        const spreadSheet = spreadSheetRef.current;
        if (!spreadSheet) return;
        if (readOnlyReasonRef.current === 'macro-preservation') {
            notifyMacroWriteBlocked();
            return;
        }
        if (readOnlyRef.current) {
            await handleSaveAs();
            return;
        }

        const ext = extRef.current.replace(/^\./, '').toLowerCase();
        const sheets = spreadSheet.getData();
        if (readOnlyReasonRef.current === 'native-excel-editing') {
            if (nativeSavePendingRef.current) {
                return;
            }
            const plan = buildNativeExcelEditPlan(
                initialSheetsRef.current,
                sheets
            );
            if (plan.unsupportedChanges.length > 0) {
                message.warning({
                    duration: 6,
                    content: t(
                        'viewer.nativeUnsupportedChange',
                        plan.unsupportedChanges.slice(0, 3).join(', ')
                    ),
                    className: 'excel-validation-error-message',
                });
                spreadSheet.setSaveEnabled(true);
                return;
            }
            if (plan.operations.length === 0) {
                spreadSheet.setSaveEnabled(false);
                return;
            }
            nativeSavePendingRef.current = true;
            handler.emit('saveNative', plan.operations);
            return;
        }

        const { export_xlsx } = await loadExcelWriter();
        const csvEncoding = csvEncodingRef.current;
        const csvDelimiter = csvDelimiterRef.current;

        if (ext !== 'xlsx' && ext !== 'xlsm' && hasFormattingChanged(initialFormattingRef.current, sheets)) {
            await new Promise<void>((resolve) => {
                const dialog = modal.confirm({
                    title: t('viewer.formatCannotPreserveTitle'),
                    content: t('viewer.formatCannotPreserveContent', ext.toUpperCase()),
                    okText: t('viewer.saveAsXlsx'),
                    cancelText: t('button.cancel'),
                    centered: true,
                    getContainer: () => document.body,
                    onOk: async () => {
                        try {
                            await export_xlsx(spreadSheet, 'xlsx', csvEncoding, { saveAs: true }, csvDelimiter);
                        } catch (error) {
                            console.error(`Failed to save Excel file: ${(error as Error).message}`);
                            throw error;
                        }
                    },
                    onCancel: () => { },
                    footer: () => (
                        <>
                            <Button
                                style={{ padding: '3px 12px', height: 'auto' }}
                                onClick={() => dialog.destroy()}
                            >
                                {t('button.cancel')}
                            </Button>
                            <Button
                                style={{ padding: '3px 12px', height: 'auto' }}
                                onClick={() => {
                                    void (async () => {
                                        dialog.destroy();
                                        try {
                                            await export_xlsx(spreadSheet, extRef.current, csvEncoding, undefined, csvDelimiter);
                                        } catch (error) {
                                            console.error(`Failed to save Excel file: ${(error as Error).message}`);
                                        }
                                    })();
                                }}
                            >
                                {t('viewer.saveAsOriginal')}
                            </Button>
                            <Button
                                type="primary"
                                style={{ padding: '3px 12px', height: 'auto' }}
                                onClick={() => {
                                    void (async () => {
                                        try {
                                            dialog.destroy();
                                            await export_xlsx(spreadSheet, 'xlsx', csvEncoding, { saveAs: true }, csvDelimiter);
                                        } catch (error) {
                                            console.error(`Failed to save Excel file: ${(error as Error).message}`);
                                        }
                                    })();
                                }}
                            >
                                {t('viewer.saveAsXlsx')}
                            </Button>
                        </>
                    ),
                    afterClose: () => resolve(),
                });
            });
            return;
        }

        try {
            await export_xlsx(spreadSheet, extRef.current, csvEncoding, undefined, csvDelimiter);
            spreadSheet.setSaveEnabled(false);
        } catch (error) {
            console.error(`Failed to save Excel file: ${(error as Error).message}`);
        }
    }, [message, modal, handleSaveAs, notifyMacroWriteBlocked]);

    const confirmSaveAs = useCallback(async (fmt: string) => {
        const spreadSheet = spreadSheetRef.current;
        if (!spreadSheet) return;
        if (
            readOnlyReasonRef.current === 'macro-preservation'
            || readOnlyReasonRef.current === 'native-excel-editing'
        ) {
            setSaveAsVisible(false);
            notifyMacroWriteBlocked();
            return;
        }
        setSaveAsVisible(false);
        try {
            const { exportSaveAs } = await loadExcelWriter();
            await exportSaveAs(spreadSheet, fmt, csvEncodingRef.current, csvDelimiterRef.current);
            if (!readOnlyRef.current) {
                spreadSheet.setSaveEnabled(false);
            }
        } catch (error) {
            console.error(`Failed to save Excel file: ${(error as Error).message}`);
        }
    }, [notifyMacroWriteBlocked]);

    useEffect(() => {
        const onKeyDown = (e: KeyboardEvent) => {
            if ((e.ctrlKey || e.metaKey) && e.code === 'KeyS') {
                e.preventDefault();
                if (readOnlyRef.current) {
                    void handleSaveAs();
                } else {
                    void handleSave();
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
                readOnlyReasonRef.current === 'macro-preservation'
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
                    notifyMacroWriteBlocked();
                }
            }
        };
        const blockInputEvent = (e: Event) => {
            if (
                readOnlyReasonRef.current !== 'macro-preservation'
                || isFindPanelTarget(e.target)
            ) {
                return;
            }
            e.preventDefault();
            e.stopPropagation();
            notifyMacroWriteBlocked();
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
    }, [handleAutoFitColumns, handleSave, handleSaveAs, notifyMacroWriteBlocked]);

    useEffect(() => {
        const container = document.getElementById('container');

        const initSpreadsheet = async (buffer: ArrayBuffer, payload: any) => {
            const fileReadOnly = payload.readOnly === true;
            const preserveMacros =
                payload.readOnlyReason === 'macro-preservation'
                || payload.readOnlyReason === 'native-excel-editing';
            if (payload.ext?.match(/csv/i)) {
                csvEncodingRef.current = detectCsvEncoding(buffer);
            }
            const { sheets, maxLength, maxCols, csvDelimiter } = await loadSheets(buffer, payload.ext);
            if (csvDelimiter) {
                csvDelimiterRef.current = csvDelimiter;
            }
            const viewRowLen = Math.max(maxLength ?? 0, MIN_VIEW_ROWS);
            const viewColLen = Math.max(maxCols ?? 0, MIN_VIEW_COLS);
            container.innerHTML = '';
            const spreadSheet = new Spreadsheet(container, {
                mode: fileReadOnly ? 'read' : 'edit',
                allowSaveAs: !preserveMacros,
                showToolbar: false,
                showEditInVSCode: isCsvLikeExt(payload.ext ?? ''),
                row: { len: viewRowLen, height: 30 },
                col: { len: viewColLen },
                view: { height: () => Math.max(180, container.clientHeight || window.innerHeight - 130) },
            });
            spreadSheetRef.current = spreadSheet;
            setActiveSpreadsheet(spreadSheet);
            spreadSheet.loadData(sheets);
            initialSheetsRef.current = cloneSheets(sheets);
            nativeSavePendingRef.current = false;
            setLoading(false);
            requestAnimationFrame(() => spreadSheet.resize());
            if (!fileReadOnly) {
                spreadSheet.on('save', () => void handleSave());
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
                ? buildFormattingSnapshot(spreadSheet.getData())
                : '';
        };

        handler.on("open", (payload) => {
            extRef.current = payload.ext ?? '';
            documentCacheIdRef.current = payload.documentCacheId ?? '';
            const fileReadOnly = payload.readOnly === true;
            const reason = payload.readOnlyReason === 'macro-preservation'
                || payload.readOnlyReason === 'native-excel-editing'
                || payload.readOnlyReason === 'file-permissions'
                ? payload.readOnlyReason as EmbeddedReadOnlyReason
                : null;
            readOnlyRef.current = fileReadOnly;
            readOnlyReasonRef.current = reason;
            setReadOnly(fileReadOnly);
            setReadOnlyReason(reason);
            loadOfficeBuffer(payload).then(async (buffer) => {
                try {
                    await initSpreadsheet(buffer, payload);
                } catch (e) {
                    const msg = (e as Error).message || String(e);
                    console.error(`Failed to load Excel file: ${msg}`, e);
                    setLoadError(msg);
                    setLoading(false);
                }
            }).catch(error => {
                const msg = (error as Error).message || String(error);
                console.error(`Failed to load Excel file: ${msg}`, error);
                setLoadError(msg);
                setLoading(false);
            });
        }).on("saveDone", () => {
            nativeSavePendingRef.current = false;
            const spreadSheet = spreadSheetRef.current;
            if (spreadSheet) {
                initialSheetsRef.current = cloneSheets(spreadSheet.getData());
                spreadSheet.setSaveEnabled(false);
            }
            message.success({
                duration: 2,
                content: t('viewer.saveSuccess'),
                className: 'excel-save-success-message',
            });
        }).on("writeBlocked", (payload) => {
            nativeSavePendingRef.current = false;
            spreadSheetRef.current?.setSaveEnabled(true);
            message.warning({
                duration: 4,
                content: payload?.message || t('viewer.macroWriteBlocked'),
                className: 'excel-validation-error-message',
            });
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
    }, [message, handleSave, handleSaveAs])

    return (
        <div className='excel-viewer'>
            <Spin spinning={loading} fullscreen={true} />
            <ExcelRibbon
                spreadsheet={activeSpreadsheet}
                readOnly={readOnly}
                allowSaveAs={
                    readOnlyReason !== 'macro-preservation'
                    && readOnlyReason !== 'native-excel-editing'
                }
                showEditInVscode={isCsvLikeExt(extRef.current)}
                onAutoFitColumns={handleAutoFitColumns}
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
                    className={`excel-readonly-banner${readOnlyReason === 'macro-preservation' ? ' excel-macro-preservation-banner' : ''}`}
                    role="status"
                    aria-live="polite"
                >
                    <span>
                        {readOnlyReason === 'macro-preservation'
                            ? t('viewer.macroReadonlyBanner')
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
                    readOnly={readOnly}
                    onChanged={() => {
                        if (!readOnlyRef.current) {
                            spreadSheetRef.current?.setSaveEnabled(true);
                        }
                    }}
                />
            )}
            <Modal
                open={saveAsVisible}
                title={t('button.saveAs')}
                onCancel={() => setSaveAsVisible(false)}
                footer={[
                    <Button key="cancel" onClick={() => setSaveAsVisible(false)} style={{ padding: '3px 12px', height: 'auto' }}>
                        {t('button.cancel')}
                    </Button>,
                    <Button key="ok" type="primary" onClick={() => void confirmSaveAs(saveAsFormat)} style={{ padding: '3px 12px', height: 'auto' }}>
                        {t('button.save')}
                    </Button>,
                ]}
                getContainer={() => document.body}
                centered
                width={360}
            >
                <div style={{ padding: '8px 0 16px' }}>
                    <div style={{ marginBottom: 12, opacity: 0.65, fontSize: 12 }}>
                        {t('viewer.chooseExportFormat')}
                    </div>
                    <Radio.Group
                        value={saveAsFormat}
                        onChange={e => setSaveAsFormat(e.target.value as string)}
                        style={{ display: 'flex', flexDirection: 'column', gap: 10 }}
                    >
                        {[
                            { value: 'xlsx', label: t('viewer.exportXlsxLabel'), desc: t('viewer.exportXlsxDesc') },
                            { value: 'csv', label: t('viewer.exportCsvLabel'), desc: t('viewer.exportCsvDesc') },
                            { value: 'xls', label: t('viewer.exportXlsLabel'), desc: t('viewer.exportXlsDesc') },
                            { value: 'ods', label: t('viewer.exportOdsLabel'), desc: t('viewer.exportOdsDesc') },
                        ].map(f => (
                            <Radio key={f.value} value={f.value} style={{ alignItems: 'flex-start' }}>
                                <span style={{ fontWeight: 500 }}>{f.label}</span>
                                <span style={{ display: 'block', fontSize: 11, opacity: 0.55, marginTop: 1 }}>{f.desc}</span>
                            </Radio>
                        ))}
                    </Radio.Group>
                </div>
            </Modal>
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
