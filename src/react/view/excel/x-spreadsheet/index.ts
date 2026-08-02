/* global window, document */
import { h } from './component/element';
import DataProxy from './core/data_proxy';
import Sheet from './component/sheet';
import Bottombar from './component/bottombar';
import { cssPrefix } from './config';
import { locale } from './locale/locale';
import './index.less';
import '@vscode/codicons/dist/codicon.css';
import {
    findAllInSheets,
    findFirstMatch,
    findNextMatch,
    replaceAllInSheets,
    replaceCellText,
    type FindMatch,
    type FindOptions,
} from '../excel_find';
import { parseSpreadsheetLink } from '../excel_hyperlink';
import {
    buildExcelTableStyleCatalog,
    excelTableNameComparisonKey,
    isValidExcelTableName,
	minimumExcelTableRangeRows,
    normalizeA1Range,
    normalizeExcelTableName,
    SIMPLE_A1_RANGE,
    type SheetChartData,
    type SheetTableData,
    type SheetTableStyle,
} from '../../../../common/excelWorkbookObjects';
import { expr2xy, xy2expr } from './core/alphabet';
import { getFontSizePxByPt } from './core/font';
import CellRange from './core/cell_range';

export interface ExtendToolbarOption {
    tip?: string;
    el?: HTMLElement;
    icon?: string;
    onClick?: (data: object, sheet: object) => void;
}

export interface Options {
    mode?: 'edit' | 'read';
    showToolbar?: boolean;
    showGrid?: boolean;
    showContextmenu?: boolean;
    showBottomBar?: boolean;
    showEditInVSCode?: boolean;
    /** Hide and suppress Save As for source formats that must stay untouched. */
    allowSaveAs?: boolean;
    extendToolbar?: {
        left?: ExtendToolbarOption[];
        right?: ExtendToolbarOption[];
    };
    autoFocus?: boolean;
    view?: {
        height: () => number;
        width?: () => number;
    };
    row?: {
        len: number;
        height: number;
    };
    col?: {
        len: number;
        width?: number;
        indexWidth?: number;
        minWidth?: number;
    };
    style?: {
        bgcolor: string;
        align: 'left' | 'center' | 'right';
        valign: 'top' | 'middle' | 'bottom';
        textwrap: boolean;
        strike: boolean;
        underline: boolean;
        color: string;
        font: {
            name: 'Helvetica';
            size: number;
            bold: boolean;
            italic: false;
        };
    };
}

export type CELL_SELECTED = 'cell-selected';
export type CELLS_SELECTED = 'cells-selected';
export type CELL_EDITED = 'cell-edited';

export type CellMerge = [number, number];

export interface SpreadsheetEventHandler {
    (envt: CELL_SELECTED, callback: (cell: Cell, rowIndex: number, colIndex: number) => void): void;
    (envt: CELLS_SELECTED, callback: (cell: Cell, parameters: { sri: number; sci: number; eri: number; eci: number }) => void): void;
    (evnt: CELL_EDITED, callback: (text: string, rowIndex: number, colIndex: number) => void): void;
}

export interface ColProperties {
    width?: number;
}

export interface CellData {
    text: string;
    style?: number;
    merge?: CellMerge;
    /** Cached Excel result used when the embedded formula engine cannot evaluate a formula. */
    formulaResult?: string | number | boolean;
    /** false 表示不可编辑（对应 Excel 锁定单元格） */
    editable?: boolean;
}

export interface RowData {
    cells: {
        [key: number]: CellData;
    };
    height?: number;
}

export interface RowsData {
    len?: number;
    [key: number]: RowData | number | undefined;
}

export interface SheetAutofilterData {
    ref: string;
    filters?: { ci: number; operator: string; value: unknown }[];
    sort?: { ci: number; order: string };
}

export interface SheetHyperlinkData {
    link: string;
    tooltip?: string;
}

export interface SheetValidationData {
    refs: string[];
    mode: string;
    type: string;
    required?: boolean;
    operator?: string;
    value?: string | string[] | number;
}

export interface SheetImageAnchor {
    col: number;
    row: number;
    width?: number;
    height?: number;
    brCol?: number;
    brRow?: number;
    editAs?: string;
}

export interface SheetImage {
    id: string;
    imageId: number;
    extension: 'jpeg' | 'png' | 'gif';
    base64: string;
    anchor: SheetImageAnchor;
}

export interface SheetBackgroundImage {
    imageId: number;
    extension: 'jpeg' | 'png' | 'gif';
    base64: string;
}

export interface SheetCommentData {
    text: string;
    author?: string;
}

export interface SheetConditionalFormattingRule extends Record<string, unknown> {
    type?: string;
    priority?: number;
    style?: Record<string, unknown>;
    /** Spreadsheet-native style used only by the canvas renderer. */
    displayStyle?: CellStyle;
}

export interface SheetConditionalFormatting {
    ref: string;
    rules: SheetConditionalFormattingRule[];
}

export interface FormatTableOptions {
    /** Excel table name. Invalid names are rejected and duplicates receive a stable suffix. */
    name?: string;
    /** Alias for name; Excel keeps Name and DisplayName aligned for newly created tables. */
    displayName?: string;
    /** Explicit A1 range; defaults to the current selection. */
    rangeRef?: string;
    headerRow?: boolean;
    totalsRow?: boolean;
    style?: Partial<SheetTableStyle>;
}

export interface SheetTableCreationContext {
    /** Stable within the current workbook session, for example the worksheet index. */
    idScope: string;
    sheetName: string;
    existingTables?: readonly SheetTableData[];
    usedNames?: readonly string[];
}

const DEFAULT_TABLE_STYLE: SheetTableStyle = {
    name: 'TableStyleMedium2',
    showFirstColumn: false,
    showLastColumn: false,
    showRowStripes: true,
    showColumnStripes: false,
};
const EXCEL_TABLE_STYLES = new Set(buildExcelTableStyleCatalog());

type WorkbookObjectAwareDataProxy = DataProxy & {
    tables?: SheetTableData[];
    charts?: SheetChartData[];
    hasNativeChartParts?: boolean;
    unsupportedNativeChartCount?: number;
};

function cloneTables(tables: readonly SheetTableData[] | undefined): SheetTableData[] {
    return (tables ?? []).map(table => ({
        ...table,
        style: { ...table.style },
    }));
}

function cloneCharts(charts: readonly SheetChartData[] | undefined): SheetChartData[] {
    return charts ? JSON.parse(JSON.stringify(charts)) as SheetChartData[] : [];
}

/**
 * DataProxy deliberately serializes a fixed core schema. Enriching each instance here keeps
 * workbook objects in normal change snapshots, so native undo/redo also restores tables/charts.
 */
function enableWorkbookObjectSnapshots(data: DataProxy): WorkbookObjectAwareDataProxy {
    const target = data as WorkbookObjectAwareDataProxy;
    const baseGetData = data.getData.bind(data);
    data.getData = () => {
        const result = baseGetData();
        const tables = cloneTables(target.tables);
        const charts = cloneCharts(target.charts);
        return {
            ...result,
            // Keep empty arrays in history snapshots so undo can clear a newly created object.
            tables,
            charts,
            ...(target.hasNativeChartParts ? { hasNativeChartParts: true } : {}),
            ...(target.unsupportedNativeChartCount
                ? { unsupportedNativeChartCount: target.unsupportedNativeChartCount }
                : {}),
        };
    };
    return target;
}

function tableNamesFromDataProxies(datas: readonly DataProxy[]): string[] {
    return datas.flatMap(data => cloneTables((data as WorkbookObjectAwareDataProxy).tables)
        .flatMap(table => [table.name, table.displayName]));
}

function parseTableRange(rangeRef: string): CellRange {
    const normalized = normalizeA1Range(rangeRef);
    if (!SIMPLE_A1_RANGE.test(normalized)) {
        throw new Error(`Plage de tableau Excel invalide : ${rangeRef}`);
    }
    const range = CellRange.valueOf(normalized);
    if (
        range.sri < 0 || range.sci < 0
        || range.eri < range.sri || range.eci < range.sci
        || range.eri >= 1_048_576 || range.eci >= 16_384
    ) {
        throw new Error(`Plage de tableau Excel hors limites : ${rangeRef}`);
    }
    return range;
}

function assertValidTableName(value: string): string {
    const normalized = normalizeExcelTableName(value);
    if (!isValidExcelTableName(normalized)) {
        throw new Error(
            'Le nom du tableau doit commencer par une lettre ou un soulignement, '
            + 'puis contenir uniquement des lettres, chiffres, points ou soulignements.'
        );
    }
    return normalized;
}

function uniqueTableName(base: string, usedNames: Iterable<string>): string {
    const used = new Set(Array.from(usedNames, excelTableNameComparisonKey));
    if (!used.has(excelTableNameComparisonKey(base))) return base;
    let suffix = 2;
    while (true) {
        const suffixText = `_${suffix}`;
        const candidate = `${base.slice(0, 255 - suffixText.length)}${suffixText}`;
        if (!used.has(excelTableNameComparisonKey(candidate))) return candidate;
        suffix += 1;
    }
}

function defaultTableName(sheetName: string, range: CellRange): string {
    const safeSheetName = sheetName
        .replace(/[^A-Za-z0-9_]/g, '_')
        .replace(/^[^A-Za-z_]+/, '')
        .slice(0, 120) || 'Sheet';
    return `Table_${safeSheetName}_${xy2expr(range.sci, range.sri)}`;
}

/** Build one native Excel table definition without changing workbook data. */
export function createSheetTableDefinition(
    rangeRef: string,
    context: SheetTableCreationContext,
    options: FormatTableOptions = {},
): SheetTableData {
    const normalizedRange = normalizeA1Range(rangeRef);
    const range = parseTableRange(normalizedRange);
    const headerRow = options.headerRow !== false;
    const totalsRow = options.totalsRow === true;
    const rowCount = range.eri - range.sri + 1;
    const minimumRows = minimumExcelTableRangeRows(totalsRow);
    if (rowCount < minimumRows) {
        throw new Error(
            `La plage ${normalizedRange} ne contient pas assez de lignes pour les options du tableau.`
        );
    }

    const existingTables = context.existingTables ?? [];
    for (const existing of existingTables) {
        const existingRange = parseTableRange(existing.rangeRef);
        if (range.intersects(existingRange)) {
            throw new Error(
                `La plage ${normalizedRange} chevauche le tableau ${existing.name} (${existing.rangeRef}).`
            );
        }
    }

    const style: SheetTableStyle = {
        ...DEFAULT_TABLE_STYLE,
        ...(options.style ?? {}),
    };
    if (!EXCEL_TABLE_STYLES.has(style.name)) {
        throw new Error(`Style de tableau Excel non pris en charge : ${style.name}`);
    }

    const requestedName = assertValidTableName(
        options.name ?? options.displayName ?? defaultTableName(context.sheetName, range)
    );
    const usedNames = [
        ...(context.usedNames ?? []),
        ...existingTables.flatMap(table => [table.name, table.displayName]),
    ];
    const name = uniqueTableName(requestedName, usedNames);
    return {
        id: `table:${context.idScope}:${normalizedRange}`,
        name,
        displayName: name,
        rangeRef: normalizedRange,
        headerRow,
        totalsRow,
        style,
    };
}

export interface SheetData {
    name?: string;
    freeze?: string;
    autofilter?: SheetAutofilterData;
    tables?: SheetTableData[];
    charts?: SheetChartData[];
    /** Native chart parts exist and must be preserved through Excel itself. */
    hasNativeChartParts?: boolean;
    /** Native chart/chartEx objects detected but not safely editable here. */
    unsupportedNativeChartCount?: number;
    hyperlinks?: Record<string, SheetHyperlinkData>;
    validations?: SheetValidationData[];
    /** Excel 工作表保护配置（不含密码） */
    sheetProtection?: Record<string, unknown>;
    images?: SheetImage[];
    backgroundImage?: SheetBackgroundImage;
    comments?: Record<string, SheetCommentData>;
    conditionalFormattings?: SheetConditionalFormatting[];
    pageSetup?: Record<string, unknown>;
    styles?: CellStyle[];
    merges?: string[];
    cols?: {
        len?: number;
        [key: number]: ColProperties | number | undefined;
    };
    rows?: RowsData;
}

export function isWorkbookTableNameAvailableInSheets(
    sheets: readonly Pick<SheetData, 'tables'>[],
    activeSheetIndex: number,
    requestedName: string,
    currentTableId?: string,
): boolean {
    const normalizedName = excelTableNameComparisonKey(requestedName);
    if (!normalizedName) return true;
    return !sheets.some((sheet, sheetIndex) => (
        (sheet.tables ?? []).some(table => {
            if (sheetIndex === activeSheetIndex && currentTableId && table.id === currentTableId) {
                return false;
            }
            return [table.name, table.displayName].some(
                name => excelTableNameComparisonKey(name) === normalizedName,
            );
        })
    ));
}

export interface SpreadsheetData {
    name?: string;
    [index: number]: SheetData;
}

export interface WorkbookStatistics {
    sheets: number;
    populatedCells: number;
    formulas: number;
    comments: number;
    conditionalFormattingRules: number;
}

export interface CellStyle {
    align?: 'left' | 'center' | 'right';
    valign?: 'top' | 'middle' | 'bottom';
    font?: {
        name?: string;
        size?: number;
        bold?: boolean;
        italic?: boolean;
    };
    bgcolor?: string;
    textwrap?: boolean;
    strike?: boolean;
    underline?: boolean;
    color?: string;
    border?: {
        top?: string[];
        right?: string[];
        bottom?: string[];
        left?: string[];
    };
    format?: string;
}

export interface Editor { }
export interface Element { }
export interface Row { }
export interface Table { }
export interface Cell { }
export interface Sheet { }

export class Spreadsheet {
    private options: Options;
    private sheetIndex: number;
    private datas: DataProxy[];
    private bottombar: Bottombar | null;
    private data: DataProxy;
    private sheet: Sheet;
    private sheetChangeListeners: ((index: number) => void)[] = [];

    constructor(selectors: string | HTMLElement, options: Options = {}) {
        let targetEl = selectors;
        this.options = { showBottomBar: true, ...options };
        this.sheetIndex = 1;
        this.datas = [];
        this.sheetChangeListeners = [];

        if (typeof selectors === 'string') {
            targetEl = document.querySelector(selectors) as HTMLElement;
        }

        this.bottombar = this.options.showBottomBar ? new Bottombar(
            () => {
                if (this.options.mode === 'read') return;
                const d = this.addSheet();
                this.sheet.resetData(d);
            },
            (index: number) => {
                const d = this.datas[index];
                this.sheet.resetData(d);
                this.data = d;
                for (const listener of this.sheetChangeListeners) {
                    listener(index);
                }
            },
            (key: string) => {
                this.handleSheetMenu(key);
            },
            (index: number, value: string) => {
                this.datas[index].name = value;
                this.sheet.trigger('change');
            },
            (from: number, to: number) => {
                this.moveSheetTo(from, to);
            }
        ) : null;

        this.data = this.addSheet();
        const rootEl = h('div', `${cssPrefix}`)
            .on('contextmenu', (evt: Event) => evt.preventDefault());

        (targetEl as HTMLElement).appendChild(rootEl.el);
        this.sheet = new Sheet(rootEl, this.data);

        if (this.bottombar !== null) {
            rootEl.child(this.bottombar.el);
        }
    }

    addSheet(name?: string, active = true): DataProxy {
        const n = name || `sheet${this.sheetIndex}`;
        const d = enableWorkbookObjectSnapshots(new DataProxy(n, this.options));
        d.change = (...args: any[]) => {
            this.sheet.trigger('change', ...args);
        };
        this.datas.push(d);

        if (this.bottombar !== null) {
            this.bottombar.addItem(n, active, this.options);
        }
        this.sheetIndex += 1;
        return d;
    }

    deleteSheet(): void {
        if (this.bottombar === null) return;

        const [oldIndex, nindex] = this.bottombar.deleteItem();
        if (oldIndex >= 0) {
            this.datas.splice(oldIndex, 1);
            if (nindex >= 0) {
                this.data = this.datas[nindex];
                this.sheet.resetData(this.datas[nindex]);
                for (const listener of this.sheetChangeListeners) {
                    listener(nindex);
                }
            }
            this.sheet.trigger('change');
        }
    }

    private uniqueSheetName(base: string): string {
        const names = this.datas.map(d => d.name);
        const trimmed = base.trim() || 'Sheet';
        if (!names.includes(trimmed)) return trimmed;
        let i = 2;
        while (names.includes(`${trimmed} (${i})`)) {
            i += 1;
        }
        return `${trimmed} (${i})`;
    }

    handleSheetMenu(key: string): void {
        if (this.options.mode === 'read' || this.bottombar === null) return;
        const index = this.bottombar.getContextSheetIndex();
        if (index < 0) return;
        if (key === 'delete') {
            this.deleteSheet();
        } else if (key === 'rename') {
            this.bottombar.startRename(index);
        } else if (key === 'duplicate') {
            this.duplicateSheet(index);
        }
    }

    moveSheetTo(from: number, to: number): void {
        if (this.bottombar === null) return;
        if (from === to || from < 0 || to < 0 || from >= this.datas.length || to >= this.datas.length) return;
        const [sheet] = this.datas.splice(from, 1);
        this.datas.splice(to, 0, sheet);
        this.bottombar.moveItem(from, to);
        const active = this.getActiveSheetIndex();
        if (active === from || active === to) {
            for (const listener of this.sheetChangeListeners) {
                listener(this.getActiveSheetIndex());
            }
        }
        this.sheet.trigger('change');
    }

    duplicateSheet(sourceIndex: number): void {
        if (this.bottombar === null) return;
        const source = this.datas[sourceIndex];
        if (!source) return;

        const copyName = this.uniqueSheetName(source.name);
        const sheetData = JSON.parse(JSON.stringify(source.getData())) as SheetData;
        sheetData.name = copyName;
        const usedNames = tableNamesFromDataProxies(this.datas);
        sheetData.tables = cloneTables((source as WorkbookObjectAwareDataProxy).tables).map((table, tableIndex) => {
            const name = uniqueTableName(assertValidTableName(table.name), usedNames);
            usedNames.push(name);
            return {
                ...table,
                id: `table:copy-${sourceIndex + 1}-${tableIndex + 1}:${normalizeA1Range(table.rangeRef)}`,
                name,
                displayName: name,
                style: { ...table.style },
            };
        });
        sheetData.charts = cloneCharts((source as WorkbookObjectAwareDataProxy).charts).map((chart, chartIndex) => ({
            ...chart,
            id: `chart:copy-${sourceIndex + 1}-${chartIndex + 1}`,
        }));

        const nd = enableWorkbookObjectSnapshots(new DataProxy(copyName, this.options));
        nd.change = (...args: any[]) => {
            this.sheet.trigger('change', ...args);
        };
        nd.setData(sheetData);

        const insertAt = sourceIndex + 1;
        this.datas.splice(insertAt, 0, nd);
        this.bottombar.insertItem(insertAt, copyName, true, this.options);

        this.data = nd;
        this.sheet.resetData(nd);
        for (const listener of this.sheetChangeListeners) {
            listener(insertAt);
        }
        this.sheet.trigger('change');
    }

    loadData(data: SpreadsheetData | SpreadsheetData[]): this {
        const ds = Array.isArray(data) ? data : [data];
        if (this.bottombar !== null) {
            this.bottombar.clear();
        }
        this.datas = [];
        if (ds.length > 0) {
            for (let i = 0; i < ds.length; i += 1) {
                const it = ds[i];
                const nd = this.addSheet(it.name, i === 0);
                nd.setData({
                    ...it,
                    ...(it.tables ? { tables: cloneTables(it.tables) } : {}),
                    ...(it.charts ? { charts: cloneCharts(it.charts) } : {}),
                });
                if (i === 0) {
                    this.data = nd;
                    this.sheet.resetData(nd);
                }
            }
        }
        return this;
    }

    getData(): SpreadsheetData[] {
        return this.datas.map(it => {
            const data = it.getData() as SheetData;
            const tables = cloneTables((it as WorkbookObjectAwareDataProxy).tables);
            const charts = cloneCharts((it as WorkbookObjectAwareDataProxy).charts);
            const { tables: _tables, charts: _charts, ...coreData } = data;
            return {
                ...coreData,
                ...(tables.length ? { tables } : {}),
                ...(charts.length ? { charts } : {}),
            };
        });
    }

    cellText(ri: number, ci: number, text: string, sheetIndex = 0): this {
        this.datas[sheetIndex].setCellText(ri, ci, text, 'finished');
        return this;
    }

    cell(ri: number, ci: number, sheetIndex = 0): Cell {
        return this.datas[sheetIndex].getCell(ri, ci);
    }

    cellStyle(ri: number, ci: number, sheetIndex = 0): CellStyle {
        return this.datas[sheetIndex].getCellStyle(ri, ci);
    }

    reRender(): this {
        this.sheet.table.render();
        const sheet = this.sheet as any;
        if (sheet.editor?.cell) {
            sheet.editor.applyCellStyle(sheet.data.getSelectedCellStyle());
        }
        return this;
    }

    resize(): this {
        this.sheet.reload();
        return this;
    }

    setZoom(scale: number): this {
        if (this.data.setZoomScale(scale)) {
            this.sheet.reload();
        }
        return this;
    }

    setSaveEnabled(enabled: boolean): this {
        (this.sheet as any).toolbar.setSaveEnabled(enabled);
        return this;
    }

    setMode(mode: 'edit' | 'read'): this {
        this.options.mode = mode;
        for (const data of this.datas) {
            data.settings.mode = mode;
        }
        this.sheet.resetData(this.data);
        return this;
    }

    executeCommand(type: string, value?: unknown): this {
        (this.sheet as any).toolbar.change(type, value);
        return this;
    }

    toggleCommand(type: string): this {
        const toolbar = (this.sheet as any).toolbar;
        const style = this.data.getSelectedCellStyle();
        if (type === 'font-bold') {
            toolbar.change(type, !style.font.bold);
        } else if (type === 'font-italic') {
            toolbar.change(type, !style.font.italic);
        } else if (type === 'underline' || type === 'strike' || type === 'textwrap') {
            toolbar.change(type, !style[type]);
        } else if (type === 'merge') {
            toolbar.change(type, !this.data.canUnmerge());
        } else if (type === 'freeze') {
            toolbar.change(type, !this.data.freezeIsActive());
        } else if (type === 'autofilter') {
            toolbar.change(type);
        } else if (type === 'paintformat') {
            toolbar.trigger(type);
        } else {
            toolbar.change(type);
        }
        return this;
    }

    executeContextCommand(type: string): this {
        (this.sheet as any).contextMenu.itemClick(type);
        return this;
    }

    setGridVisible(visible: boolean): this {
        for (const data of this.datas) {
            data.settings.showGrid = visible;
        }
        this.sheet.table.render();
        return this;
    }

    getActiveSheetIndex(): number {
        if (this.bottombar === null) return 0;
        const bar = this.bottombar as any;
        const idx = bar.items.findIndex((it: any) => it === bar.activeEl);
        return idx >= 0 ? idx : 0;
    }

    activateSheet(index: number): this {
        if (index < 0 || index >= this.datas.length) return this;
        if (this.getActiveSheetIndex() === index) return this;
        if (this.bottombar !== null) {
            const bar = this.bottombar as any;
            const item = bar.items[index];
            if (item) bar.clickSwap2(item);
        } else {
            this.data = this.datas[index];
            this.sheet.resetData(this.data);
        }
        return this;
    }

    scrollToCell(ri: number, ci: number, sheetIndex = 0): this {
        this.activateSheet(sheetIndex);
        const data = this.datas[sheetIndex];
        if (!data) {
            return this;
        }
        const maxRi = Math.max(0, (data.rows.len ?? 1) - 1);
        const maxCi = Math.max(0, (data.cols.len ?? 1) - 1);
        const row = Math.min(Math.max(0, ri), maxRi);
        const col = Math.min(Math.max(0, ci), maxCi);
        (this.sheet as any).scrollToCell(row, col);
        return this;
    }

    getSelection(): { ri: number; ci: number; sheetIndex: number } {
        const { ri = 0, ci = 0 } = this.data.selector ?? {};
        return { ri, ci, sheetIndex: this.getActiveSheetIndex() };
    }

    getSelectionRangeRef(): string {
        return this.data.selector.range.toString();
    }

    getActiveSheetCharts(): SheetChartData[] {
        return cloneCharts((this.data as WorkbookObjectAwareDataProxy).charts);
    }

    hasNativeWorkbookObjects(): boolean {
        return this.datas.some(data => {
            const target = data as WorkbookObjectAwareDataProxy;
            return (target.tables?.length ?? 0) > 0
                || (target.charts?.length ?? 0) > 0
                || target.hasNativeChartParts === true;
        });
    }

    getActiveSheetUnsupportedNativeChartCount(): number {
        const count = (this.data as WorkbookObjectAwareDataProxy).unsupportedNativeChartCount;
        return typeof count === 'number' && Number.isInteger(count) && count > 0 ? count : 0;
    }

    addOrUpdateChart(chart: SheetChartData): this {
        if (!chart.id?.trim()) throw new Error('Un graphique doit avoir un identifiant stable.');
        const nextChart = cloneCharts([chart])[0];
        this.data.changeData(() => {
            const target = this.data as WorkbookObjectAwareDataProxy;
            const charts = cloneCharts(target.charts);
            const index = charts.findIndex(existing => existing.id === nextChart.id);
            if (index >= 0) charts[index] = nextChart;
            else charts.push(nextChart);
            target.charts = charts;
        });
        return this;
    }

    removeChart(chartId: string): this {
        this.data.changeData(() => {
            const target = this.data as WorkbookObjectAwareDataProxy;
            target.charts = cloneCharts(target.charts).filter(chart => chart.id !== chartId);
        });
        return this;
    }

    getSelectedComment(): SheetCommentData | undefined {
        const { ri, ci } = this.getSelection();
        return this.data.getComment(ri, ci) ?? undefined;
    }

    setSelectedComment(text: string, author?: string): this {
        const { ri, ci } = this.getSelection();
        this.data.setComment(ri, ci, {
            text: text.trim(),
            ...(author ? { author } : {}),
        });
        this.reRender();
        return this;
    }

    listComments(): Array<SheetCommentData & { sheet: string; address: string }> {
        return this.datas.flatMap((data: any) => Object.entries(data.comments ?? {}).map(
            ([address, comment]) => ({
                ...(comment as SheetCommentData),
                address,
                sheet: data.name,
            }),
        ));
    }

    getWorkbookStatistics(): WorkbookStatistics {
        let populatedCells = 0;
        let formulas = 0;
        let comments = 0;
        let conditionalFormattingRules = 0;
        this.datas.forEach((data: any) => {
            data.rows.each((_ri: number, row: RowData) => {
                Object.values(row?.cells ?? {}).forEach(cell => {
                    if (cell?.text != null && `${cell.text}` !== '') populatedCells += 1;
                    if (`${cell?.text ?? ''}`.startsWith('=')) formulas += 1;
                });
            });
            comments += Object.keys(data.comments ?? {}).length;
            conditionalFormattingRules += (data.conditionalFormattings ?? [])
                .reduce((total: number, item: SheetConditionalFormatting) => (
                    total + (item.rules?.length ?? 0)
                ), 0);
        });
        return {
            sheets: this.datas.length,
            populatedCells,
            formulas,
            comments,
            conditionalFormattingRules,
        };
    }

    isSheetProtected(): boolean {
        return !!this.data.sheetProtection;
    }

    toggleSheetProtection(): boolean {
        this.data.changeData(() => {
            this.data.sheetProtection = this.data.sheetProtection
                ? null
                : { sheet: true };
        });
        this.reRender();
        return this.isSheetProtected();
    }

    addConditionalFormatting(rule: SheetConditionalFormattingRule): this {
        const ref = this.data.selector.range.toString();
        this.data.changeData(() => {
            this.data.conditionalFormattings = this.data.conditionalFormattings ?? [];
            const nextPriority = this.data.conditionalFormattings
                .flatMap((item: SheetConditionalFormatting) => item.rules ?? [])
                .reduce((highest: number, item: SheetConditionalFormattingRule) => (
                    Math.max(highest, Number(item.priority) || 0)
                ), 0) + 1;
            this.data.conditionalFormattings.push({
                ref,
                rules: [{
                    ...rule,
                    priority: rule.priority ?? nextPriority,
                }],
            });
        });
        this.reRender();
        return this;
    }

    clearConditionalFormatting(): this {
        this.data.changeData(() => {
            this.data.conditionalFormattings = [];
        });
        this.reRender();
        return this;
    }

    addBlankSheet(name = 'Nouveau classeur'): this {
        this.addSheet(this.uniqueSheetName(name), false);
        this.activateSheet(this.datas.length - 1);
        this.sheet.trigger('change');
        return this;
    }

    appendSheets(sheets: SheetData[]): this {
        for (const sheetData of sheets) {
            const name = this.uniqueSheetName(sheetData.name || 'Données importées');
            const target = this.addSheet(name, false);
            const usedNames = tableNamesFromDataProxies(this.datas);
            const tables = cloneTables(sheetData.tables).map((table, tableIndex) => {
                const tableName = uniqueTableName(assertValidTableName(table.name), usedNames);
                usedNames.push(tableName);
                return {
                    ...table,
                    id: `table:append-${this.datas.length}-${tableIndex + 1}:${normalizeA1Range(table.rangeRef)}`,
                    name: tableName,
                    displayName: tableName,
                    style: { ...table.style },
                };
            });
            const charts = cloneCharts(sheetData.charts);
            target.setData({
                ...sheetData,
                name,
                ...(tables.length ? { tables } : {}),
                ...(charts.length ? { charts } : {}),
            });
        }
        if (sheets.length > 0) {
            this.activateSheet(this.datas.length - 1);
            this.sheet.trigger('change');
        }
        return this;
    }

    getActiveSheetTables(): SheetTableData[] {
        return cloneTables((this.data as WorkbookObjectAwareDataProxy).tables);
    }

    isWorkbookTableNameAvailable(name: string, currentTableId?: string): boolean {
        return isWorkbookTableNameAvailableInSheets(
            this.datas.map(data => ({
                tables: cloneTables((data as WorkbookObjectAwareDataProxy).tables),
            })),
            this.getActiveSheetIndex(),
            name,
            currentTableId,
        );
    }

    updateTable(table: SheetTableData): this {
        const currentTables = this.getActiveSheetTables();
        const tableIndex = currentTables.findIndex(existing => existing.id === table.id);
        if (tableIndex < 0) throw new Error(`Tableau introuvable : ${table.id}`);
        const existingTables = currentTables.filter((_, index) => index !== tableIndex);
        const usedNames = this.datas.flatMap(data => cloneTables(
            (data as WorkbookObjectAwareDataProxy).tables,
        ).filter(existing => data !== this.data || existing.id !== table.id)
            .flatMap(existing => [existing.name, existing.displayName]));
        const normalized = createSheetTableDefinition(table.rangeRef, {
            idScope: `sheet-${this.getActiveSheetIndex() + 1}`,
            sheetName: this.data.name,
            existingTables,
            usedNames,
        }, {
            name: table.name,
            displayName: table.displayName,
            headerRow: table.headerRow,
            totalsRow: table.totalsRow,
            style: table.style,
        });
        normalized.id = currentTables[tableIndex].id;
        this.data.changeData(() => {
            currentTables[tableIndex] = normalized;
            (this.data as WorkbookObjectAwareDataProxy).tables = currentTables;
        });
        this.reRender();
        return this;
    }

    removeTable(tableId: string): this {
        this.data.changeData(() => {
            const target = this.data as WorkbookObjectAwareDataProxy;
            // Match Excel's Unlist behavior: remove the native table but preserve cell formatting.
            target.tables = cloneTables(target.tables).filter(table => table.id !== tableId);
        });
        this.reRender();
        return this;
    }

    formatSelectionAsTable(options: FormatTableOptions = {}): this {
		if (options.headerRow === false) {
			throw new Error(
				'La création d’un tableau sans en-tête est désactivée pour éviter que Microsoft Excel déplace les cellules.'
			);
		}
        const range = options.rangeRef
            ? parseTableRange(options.rangeRef)
            : this.data.selector.range;
        const existingTables = cloneTables((this.data as WorkbookObjectAwareDataProxy).tables);
        const table = createSheetTableDefinition(range.toString(), {
            idScope: `sheet-${this.getActiveSheetIndex() + 1}`,
            sheetName: this.data.name,
            existingTables,
            usedNames: tableNamesFromDataProxies(this.datas),
        }, options);
        this.data.changeData(() => {
            (this.data as WorkbookObjectAwareDataProxy).tables = [...existingTables, table];
        });
        this.reRender();
        return this;
    }

    setPageSetup(patch: Record<string, unknown>): this {
        this.data.changeData(() => {
            this.data.pageSetup = {
                ...(this.data.pageSetup ?? {}),
                ...patch,
            };
        });
        return this;
    }

    autoFitRows(sheetIndex = this.getActiveSheetIndex()): this {
        const data = this.datas[sheetIndex];
        if (!data) return this;
        const rowLen = data.rows?.len ?? 0;
        const colLen = data.cols?.len ?? 0;
        data.changeData(() => {
            for (let ri = 0; ri < rowLen; ri += 1) {
                let lineCount = 1;
                for (let ci = 0; ci < colLen; ci += 1) {
                    const text = `${data.getCell(ri, ci)?.text ?? ''}`;
                    lineCount = Math.max(lineCount, text.split(/\r\n|\r|\n/).length);
                }
                data.rows.setHeight(ri, Math.min(240, Math.max(24, lineCount * 22)));
            }
        });
        if (sheetIndex === this.getActiveSheetIndex()) this.sheet.resetData(data);
        return this;
    }

    toggleFormulaDisplay(): boolean {
        this.data.settings.evalPaused = !this.data.settings.evalPaused;
        this.reRender();
        return this.data.settings.evalPaused;
    }

    setFormulaBarVisible(visible: boolean): this {
        const formulaBar = (this.sheet as any).formulaBar?.el;
        if (visible) formulaBar?.show();
        else formulaBar?.hide();
        this.sheet.reload();
        return this;
    }

    setHeadingsVisible(visible: boolean): this {
        this.datas.forEach(data => {
            data.settings.showHeaders = visible;
        });
        this.reRender();
        return this;
    }

    copySelectionToNewSheet(name = 'Table extraite'): this {
        const source = this.data;
        const range = source.selector.range;
        const target = this.addSheet(this.uniqueSheetName(name), false);
        target.changeData(() => {
            for (let ri = range.sri; ri <= range.eri; ri += 1) {
                for (let ci = range.sci; ci <= range.eci; ci += 1) {
                    const sourceCell = source.getCell(ri, ci);
                    if (!sourceCell) continue;
                    target.rows.setCell(ri - range.sri, ci - range.sci, structuredClone(sourceCell));
                }
            }
            target.rows.len = Math.max(1, range.eri - range.sri + 1);
            target.cols.len = Math.max(1, range.eci - range.sci + 1);
        });
        this.activateSheet(this.datas.length - 1);
        this.sheet.trigger('change');
        return this;
    }

    getSelectionMatrix(): string[][] {
        const range = this.data.selector.range;
        const result: string[][] = [];
        for (let ri = range.sri; ri <= range.eri; ri += 1) {
            const row: string[] = [];
            for (let ci = range.sci; ci <= range.eci; ci += 1) {
                const cell = this.data.getCell(ri, ci);
                row.push(`${cell?.formulaResult ?? cell?.text ?? ''}`);
            }
            result.push(row);
        }
        return result;
    }

    textToColumns(delimiter: string): this {
        const range = this.data.selector.range;
        const separator = delimiter || ',';
        this.data.changeData(() => {
            for (let ri = range.sri; ri <= range.eri; ri += 1) {
                const source = `${this.data.getCell(ri, range.sci)?.text ?? ''}`;
                source.split(separator).forEach((value, offset) => {
                    const cell = this.data.rows.getCellOrNew(ri, range.sci + offset);
                    cell.text = value.trim();
                    delete cell.formulaResult;
                });
            }
        });
        this.reRender();
        return this;
    }

    removeDuplicateRows(): number {
        const range = this.data.selector.range;
        const seen = new Set<string>();
        const unique: CellData[][] = [];
        for (let ri = range.sri; ri <= range.eri; ri += 1) {
            const row: CellData[] = [];
            for (let ci = range.sci; ci <= range.eci; ci += 1) {
                row.push(structuredClone(this.data.getCell(ri, ci) ?? { text: '' }));
            }
            const key = JSON.stringify(row.map(cell => cell.text));
            if (!seen.has(key)) {
                seen.add(key);
                unique.push(row);
            }
        }
        const removed = range.eri - range.sri + 1 - unique.length;
        this.data.changeData(() => {
            for (let offset = 0; offset <= range.eri - range.sri; offset += 1) {
                for (let ci = range.sci; ci <= range.eci; ci += 1) {
                    const target = this.data.rows.getCellOrNew(range.sri + offset, ci);
                    const source = unique[offset]?.[ci - range.sci];
                    if (source) Object.assign(target, source);
                    else {
                        target.text = '';
                        delete target.formulaResult;
                    }
                }
            }
        });
        this.reRender();
        return removed;
    }

    addSubtotal(): this {
        const range = this.data.selector.range;
        const targetRow = range.eri + 1;
        this.data.changeData(() => {
            const label = this.data.rows.getCellOrNew(targetRow, range.sci);
            label.text = 'Total';
            for (let ci = range.sci + 1; ci <= range.eci; ci += 1) {
                const cell = this.data.rows.getCellOrNew(targetRow, ci);
                cell.text = `=SUM(${xy2expr(ci, range.sri)}:${xy2expr(ci, range.eri)})`;
            }
            this.data.rows.len = Math.max(this.data.rows.len, targetRow + 1);
        });
        this.reRender();
        return this;
    }

    addForecastRow(): this {
        const range = this.data.selector.range;
        const targetRow = range.eri + 1;
        this.data.changeData(() => {
            for (let ci = range.sci; ci <= range.eci; ci += 1) {
                const last = Number(this.data.getCell(range.eri, ci)?.text);
                const previous = Number(this.data.getCell(Math.max(range.sri, range.eri - 1), ci)?.text);
                const cell = this.data.rows.getCellOrNew(targetRow, ci);
                cell.text = Number.isFinite(last) && Number.isFinite(previous)
                    ? String(last + (last - previous))
                    : '';
            }
            this.data.rows.len = Math.max(this.data.rows.len, targetRow + 1);
        });
        this.reRender();
        return this;
    }

    toggleWorkbookProtection(): boolean {
        const protectedWorkbook = this.datas.every(data => !!data.sheetProtection);
        this.datas.forEach(data => {
            data.changeData(() => {
                data.sheetProtection = protectedWorkbook ? null : { sheet: true };
            });
        });
        this.reRender();
        return !protectedWorkbook;
    }

    formulaAudit(kind: 'precedents' | 'dependents' | 'errors'): string[] {
        const { ri, ci } = this.getSelection();
        const selectedAddress = xy2expr(ci, ri);
        if (kind === 'precedents') {
            const formula = `${this.data.getCell(ri, ci)?.text ?? ''}`;
            return Array.from(new Set(
                formula.match(/\$?[A-Z]{1,3}\$?\d+/gi)?.map(value => value.replace(/\$/g, '').toUpperCase()) ?? [],
            ));
        }
        const matches: string[] = [];
        this.data.rows.each((rowIndex: number, row: RowData) => {
            Object.entries(row?.cells ?? {}).forEach(([columnIndex, cell]) => {
                const text = `${cell?.text ?? ''}`;
                if (kind === 'errors' && /^#(?:REF|VALUE|NAME|DIV\/0|N\/A|NUM|NULL)!?/i.test(text)) {
                    matches.push(xy2expr(Number(columnIndex), rowIndex));
                } else if (
                    kind === 'dependents'
                    && text.startsWith('=')
                    && new RegExp(`(?:^|[^A-Z0-9_])\\$?${selectedAddress.replace(/\d+$/, '')}\\$?${selectedAddress.match(/\d+$/)?.[0]}(?:[^A-Z0-9_]|$)`, 'i').test(text)
                ) {
                    matches.push(xy2expr(Number(columnIndex), rowIndex));
                }
            });
        });
        return matches;
    }

    insertImage(base64: string, extension: SheetImage['extension'], background = false): this {
        const imageId = this.datas.flatMap(data => data.images ?? [])
            .reduce((highest, item) => Math.max(highest, Number(item.imageId) || 0), 0) + 1;
        if (background) {
            this.data.changeData(() => {
                this.data.backgroundImage = { imageId, extension, base64 };
            });
        } else {
            const { ri, ci } = this.getSelection();
            this.data.changeData(() => {
                this.data.images = this.data.images ?? [];
                this.data.images.push({
                    id: `image-${imageId}`,
                    imageId,
                    extension,
                    base64,
                    anchor: { row: ri, col: ci, width: 320, height: 200, editAs: 'oneCell' },
                });
            });
        }
        this.sheet.resetData(this.data);
        this.sheet.trigger('change');
        return this;
    }

    arrangeSelectedImage(direction: 'forward' | 'backward'): this {
        const images = this.data.images ?? [];
        if (images.length < 2) return this;
        const selected = (this.sheet as any).sheetImages?.selectedIndex;
        const index = selected >= 0 ? selected : images.length - 1;
        const target = direction === 'forward'
            ? Math.min(images.length - 1, index + 1)
            : Math.max(0, index - 1);
        if (index === target) return this;
        this.data.changeData(() => {
            const [image] = images.splice(index, 1);
            images.splice(target, 0, image);
        });
        this.sheet.resetData(this.data);
        return this;
    }

    sortSelection(order: 'asc' | 'desc'): this {
        const { ri, ci } = this.getSelection();
        if (!this.data.autoFilter.active()) {
            const content = this.data.contentRange();
            const hasValue = (row: number, col: number) => {
                const cell = this.data.getCell(row, col);
                return cell != null && `${cell.text ?? ''}`.trim() !== '';
            };
            let sri = ri;
            let eri = ri;
            while (sri > 0 && hasValue(sri - 1, ci)) sri -= 1;
            while (eri < content.eri && hasValue(eri + 1, ci)) eri += 1;
            if (sri === eri) return this;

            const columnHasValue = (col: number) => {
                for (let row = sri; row <= eri; row += 1) {
                    if (hasValue(row, col)) return true;
                }
                return false;
            };
            let sci = ci;
            let eci = ci;
            while (sci > 0 && columnHasValue(sci - 1)) sci -= 1;
            while (eci < content.eci && columnHasValue(eci + 1)) eci += 1;
            this.data.changeData(() => {
                this.data.autoFilter.ref = new CellRange(sri, sci, eri, eci).toString();
            });
        }
        const filterRange = this.data.autoFilter.range();
        if (ci < filterRange.sci || ci > filterRange.eci) return this;
        const items = this.data.autoFilter.items(
            ci,
            (ri: number, columnIndex: number) => this.data.rows.getCell(ri, columnIndex),
        );
        this.data.setAutoFilter(ci, order, 'in', Object.keys(items));
        this.reRender();
        return this;
    }

    autoFitColumns(sheetIndex = this.getActiveSheetIndex()): this {
        const data = this.datas[sheetIndex];
        if (!data || typeof document === 'undefined') return this;

        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        if (!context) return this;

        const defaultStyle = data.defaultStyle();
        const rowLen = data.rows?.len ?? 0;
        const colLen = data.cols?.len ?? 0;
        const minWidth = data.cols?.minWidth ?? 60;
        const maxWidth = 720;
        const textPadding = 22;

        data.changeData(() => {
            for (let ci = 0; ci < colLen; ci += 1) {
                let width = minWidth;

                for (let ri = 0; ri < rowLen; ri += 1) {
                    const cell = data.getCell(ri, ci);
                    if (!cell || !cell.text || cell.merge) continue;

                    const style = data.getCellStyleOrDefault(ri, ci) ?? defaultStyle;
                    const fontName = style.font?.name || defaultStyle.font.name || 'Arial';
                    const fontSize = getFontSizePxByPt(style.font?.size || defaultStyle.font.size || 11);
                    const fontWeight = style.font?.bold ? '700' : '400';
                    const fontStyle = style.font?.italic ? 'italic' : 'normal';
                    context.font = `${fontStyle} ${fontWeight} ${fontSize}px ${fontName}`;

                    const lines = String(cell.text).split(/\r\n|\r|\n/);
                    for (const line of lines) {
                        const measured = Math.ceil(context.measureText(line).width) + textPadding;
                        if (measured > width) {
                            width = measured;
                        }
                    }
                }

                data.cols.setWidth(ci, Math.max(minWidth, Math.min(maxWidth, width)));
            }
        });

        if (sheetIndex === this.getActiveSheetIndex()) {
            this.sheet.resetData(data);
        }

        return this;
    }

    onSheetChange(cb: (index: number) => void): this {
        this.sheetChangeListeners.push(cb);
        return this;
    }

    onOpenLink(cb: (payload: { link: string; tooltip?: string }) => void): this {
        this.sheet.on('open-link', cb);
        return this;
    }

    onProtectedCellDblClick(cb: () => void): this {
        this.sheet.on('protected-cell-dblclick', cb);
        return this;
    }

    onValidationError(cb: (message: string) => void): this {
        this.sheet.on('validation-error', cb);
        return this;
    }

    findFirst(text: string, options: FindOptions = {}): FindMatch | null {
        return findFirstMatch(this.getData(), text, options, this.getActiveSheetIndex());
    }

    findNext(text: string, from: FindMatch, options: FindOptions = {}, backward = false): FindMatch | null {
        return findNextMatch(this.getData(), text, from, options, this.getActiveSheetIndex(), backward);
    }

    findAll(text: string, options: FindOptions = {}): FindMatch[] {
        return findAllInSheets(this.getData(), text, options, this.getActiveSheetIndex());
    }

    gotoMatch(match: FindMatch): this {
        this.scrollToCell(match.ri, match.ci, match.sheetIndex);
        return this;
    }

    replaceAt(match: FindMatch, findText: string, replaceText: string, options: FindOptions = {}): boolean {
        const sheets = this.getData();
        const changed = replaceCellText(sheets, match, findText, replaceText, options);
        if (changed) {
            const data = this.datas[match.sheetIndex];
            if (data) {
                data.setData(sheets[match.sheetIndex]);
                if (match.sheetIndex === this.getActiveSheetIndex()) {
                    this.sheet.resetData(data);
                }
                this.sheet.trigger('change');
            }
        }
        return changed;
    }

    replaceAll(findText: string, replaceText: string, options: FindOptions = {}): number {
        const sheets = this.getData();
        const count = replaceAllInSheets(sheets, findText, replaceText, options, this.getActiveSheetIndex());
        if (count > 0) {
            for (let i = 0; i < this.datas.length; i += 1) {
                this.datas[i].setData(sheets[i]);
            }
            this.sheet.resetData(this.data);
            this.sheet.trigger('change');
        }
        return count;
    }

    followHyperlink(payload: { link: string }): this {
        const parsed = parseSpreadsheetLink(payload.link);
        if (parsed.type === 'external') {
            this.sheet.trigger('open-external-link', parsed.url);
            return this;
        }
        const sheetIndex = this.datas.findIndex((d) => d.name === parsed.sheetName);
        if (sheetIndex < 0) return this;
        const [ci, ri] = expr2xy(parsed.ref);
        this.scrollToCell(ri, ci, sheetIndex);
        return this;
    }

    on(eventName: string, func: (...args: any[]) => void): this {
        this.sheet.on(eventName, func);
        return this;
    }

    validate(): boolean {
        const { validations } = this.data;
        return validations.errors.size <= 0;
    }

    change(cb: (json: SpreadsheetData) => void): this {
        this.sheet.on('change', cb);
        return this;
    }

    static locale(lang: string, message: object): void {
        locale(lang, message);
    }
}

const spreadsheet = (el: string | HTMLElement, options: Options = {}): Spreadsheet => new Spreadsheet(el, options);

if (window) {
    (window as any).x_spreadsheet = spreadsheet;
    (window as any).x_spreadsheet.locale = (lang: string, message: object) => locale(lang, message);
}

export default Spreadsheet;
export { spreadsheet };
export type { FindMatch, FindOptions } from '../excel_find'; 
