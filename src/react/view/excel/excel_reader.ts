import type ExcelJS from '@cweijan/exceljs';
import type * as XLSX from 'xlsx';
import {
    normalizeA1Range,
    SIMPLE_A1_RANGE,
    type SheetTableData,
} from '../../../common/excelWorkbookObjects';
import { decodeCsvBuffer } from './csvEncoding';
import { DEFAULT_ROW_HEIGHT_PX, excelFreezeToExpr, excelRowHeightToPx, readAutofilterRef } from './excel_meta';
import {
    normalizeOoxmlRelationshipTargets,
    normalizeSpreadsheetMlElementPrefixes,
} from './ooxml_namespace';
import { readWorksheetSortStateXml } from './excel_sort_state';
import {
    readOoxmlChartInventory,
    type OoxmlChartInventoryResult,
} from './chart-ooxml-reader';
import {
    inspectOoxmlZipCentralDirectory,
    validateOoxmlZipInflationBounds,
} from './ooxml-zip-preflight';
import {
    createExcelColorResolver,
    excelJsCellToStyle,
    excelJsStyleToCellStyle,
    hexToArgb,
    type ExcelColorResolver,
    StyleRegistry,
} from './excel_styles';
import { mergeHyperlinkMaps, readCellHyperlink } from './excel_hyperlink';
import { readWorksheetBackgroundImage, readWorksheetImages } from './excel_images';
import { readWorksheetValidations } from './excel_validation';
import {
    isWorksheetProtected,
    readCellEditableFromExcel,
    readWorksheetProtection,
} from './excel_protection';
import type {
    CellData,
    SheetCommentData,
    SheetConditionalFormatting,
    SheetData,
} from './x-spreadsheet/index';

type RowMap = NonNullable<SheetData['rows']>;

type ExcelJsWorksheetWithMerges = ExcelJS.Worksheet & {
    _merges?: Record<string, string | { range?: string }>;
    model: ExcelJS.Worksheet['model'] & {
        mergeCells?: string[];
        merges?: string[];
    };
};

type ExcelJsWorksheetExtras = ExcelJS.Worksheet & {
    conditionalFormattings?: ExcelJS.ConditionalFormattingOptions[];
};

type ExcelJsTableModel = {
    name?: string;
    displayName?: string;
    ref?: string;
    tableRef?: string;
    autoFilterRef?: string;
    headerRow?: boolean;
    totalsRow?: boolean;
    style?: {
        theme?: string;
        showFirstColumn?: boolean;
        showLastColumn?: boolean;
        showRowStripes?: boolean;
        showColumnStripes?: boolean;
    };
};

type ExcelJsTableRuntime = Partial<ExcelJS.Table> & {
    model?: ExcelJsTableModel;
};

function cloneJson<T>(value: T): T {
    return JSON.parse(JSON.stringify(value)) as T;
}

function normalizeConditionalColors(value: unknown, resolveColor: ExcelColorResolver): unknown {
    if (Array.isArray(value)) {
        return value.map(item => normalizeConditionalColors(item, resolveColor));
    }
    if (!value || typeof value !== 'object') return value;
    const record = value as Record<string, unknown>;
    const isColor = ['argb', 'theme', 'indexed', 'tint', 'auto']
        .some(key => Object.prototype.hasOwnProperty.call(record, key));
    if (isColor) {
        const color = resolveColor(record as Partial<ExcelJS.Color>);
        const argb = hexToArgb(color);
        return argb ? { argb } : cloneJson(record);
    }
    return Object.fromEntries(
        Object.entries(record).map(([key, item]) => [
            key,
            normalizeConditionalColors(item, resolveColor),
        ]),
    );
}

function readConditionalFormattings(
    worksheet: ExcelJS.Worksheet,
    resolveColor: ExcelColorResolver,
): SheetConditionalFormatting[] {
    const source = (worksheet as ExcelJsWorksheetExtras).conditionalFormattings ?? [];
    return source
        .filter(item => typeof item?.ref === 'string' && Array.isArray(item.rules))
        .map(item => {
            const normalized = normalizeConditionalColors(item, resolveColor) as ExcelJS.ConditionalFormattingOptions;
            return {
                ref: normalized.ref,
                rules: normalized.rules.map(rule => {
                    const displayStyle = excelJsStyleToCellStyle(
                        rule.style as Partial<ExcelJS.Style> | undefined,
                        resolveColor,
                    );
                    return {
                        ...rule,
                        ...(displayStyle ? { displayStyle } : {}),
                    };
                }),
            } as SheetConditionalFormatting;
        });
}

function noteToComment(note: ExcelJS.Cell['note']): SheetCommentData | undefined {
    if (!note) return undefined;
    if (typeof note === 'string') return { text: note };
    if (Array.isArray(note)) {
        const text = note.map(part => part?.text ?? '').join('');
        return text ? { text } : undefined;
    }
    const noteValue = note as {
        texts?: { text?: string }[];
        author?: string;
    };
    const text = noteValue.texts?.map(part => part.text ?? '').join('') ?? '';
    return text ? { text, ...(noteValue.author ? { author: noteValue.author } : {}) } : undefined;
}

function formulaResult(cell: ExcelJS.Cell): CellData['formulaResult'] | undefined {
    const result = cell.result;
    if (result == null) return undefined;
    if (result instanceof Date) return result.toISOString();
    if (typeof result === 'string' || typeof result === 'number' || typeof result === 'boolean') {
        return result;
    }
    return undefined;
}

export interface ExcelData {
    sheets: SheetData[];
    maxCols: number;
    maxLength?: number;
    /** Detected column delimiter when loading CSV/TSV */
    csvDelimiter?: string;
}

const MIN_COL_WIDTH = 70;
const MAX_COL_WIDTH = 300;
const DEFAULT_COL_WIDTH = 100;
const CHAR_WIDTH = 8;
const MAX_ROWS_TO_CHECK = 10;

const clampColWidth = (width: number) => Math.min(Math.max(width, MIN_COL_WIDTH), MAX_COL_WIDTH);

const calculateColWidth = (rows: any[], colIndex: number): number => {
    let maxLength = 0;
    for (let i = 0; i < Math.min(rows.length, MAX_ROWS_TO_CHECK); i += 1) {
        const cell = rows[i][colIndex];
        if (cell) {
            const length = String(cell).length;
            if (length > maxLength) {
                maxLength = length;
            }
        }
    }
    return clampColWidth(maxLength * CHAR_WIDTH);
};

const excelColWidthToPx = (width?: number) => {
    if (width == null) return null;
    return Math.round(width * 7 + 5);
};

const normalizeMergeRange = (merge: unknown): string | null => {
    if (typeof merge === 'string') return merge;
    if (merge && typeof merge === 'object' && 'range' in merge) {
        const range = (merge as { range?: unknown }).range;
        return typeof range === 'string' ? range : null;
    }
    return null;
};

const readWorksheetMerges = (worksheet: ExcelJS.Worksheet): string[] => {
    const ws = worksheet as ExcelJsWorksheetWithMerges;
    const mergeCandidates = [
        ...(ws.model?.merges ?? []),
        ...(ws.model?.mergeCells ?? []),
        ...Object.values(ws._merges ?? {}),
    ];
    const merges = mergeCandidates
        .map(normalizeMergeRange)
        .filter((it): it is string => Boolean(it));
    return Array.from(new Set(merges));
};

const readWorksheetTables = (worksheet: ExcelJS.Worksheet): SheetTableData[] => {
    const rawTables = (worksheet as unknown as { getTables?: () => unknown[] }).getTables?.() ?? [];
    return rawTables.flatMap((raw, index) => {
        const candidate = (Array.isArray(raw) ? raw[0] : raw) as ExcelJsTableRuntime | undefined;
        if (!candidate) return [];
        const model = candidate.model ?? {};
        const rangeRef = normalizeA1Range(
            model.tableRef ?? model.autoFilterRef ?? model.ref ?? candidate.ref ?? '',
        );
        if (!SIMPLE_A1_RANGE.test(rangeRef)) return [];
        const fallbackName = `Table_${worksheet.id}_${index + 1}`;
        const name = String(model.name ?? candidate.name ?? fallbackName);
        const displayName = String(model.displayName ?? candidate.displayName ?? name);
        const style = model.style ?? candidate.style ?? {};
        return [{
            id: `table:${worksheet.id}:${rangeRef}`,
            name,
            displayName,
            rangeRef,
            headerRow: (model.headerRow ?? candidate.headerRow) !== false,
            totalsRow: (model.totalsRow ?? candidate.totalsRow) === true,
            style: {
                name: String(style.theme ?? 'TableStyleMedium2'),
                showFirstColumn: style.showFirstColumn === true,
                showLastColumn: style.showLastColumn === true,
                showRowStripes: style.showRowStripes === true,
                showColumnStripes: style.showColumnStripes === true,
            },
        } satisfies SheetTableData];
    });
};

const expandSizeForMerge = (merge: string, size: { maxRow: number; maxCols: number }) => {
    const endAddress = merge.split(':').at(-1)?.replace(/\$/g, '');
    const match = /^([A-Z]+)(\d+)$/i.exec(endAddress ?? '');
    if (!match) return;
    let col = 0;
    const letters = match[1].toUpperCase();
    for (let i = 0; i < letters.length; i += 1) {
        col = col * 26 + letters.charCodeAt(i) - 64;
    }
    size.maxRow = Math.max(size.maxRow, Number(match[2]));
    size.maxCols = Math.max(size.maxCols, col);
};

type SheetJsUtils = typeof import('xlsx')['utils'];

const readSheetJsMerges = (worksheet: XLSX.WorkSheet, utils: SheetJsUtils) => (worksheet['!merges'] ?? [])
    .map(merge => utils.encode_range(merge));

const expandSizeForSheetJsMerge = (merge: XLSX.Range, size: { maxRow: number; maxCols: number }) => {
    size.maxRow = Math.max(size.maxRow, merge.e.r + 1);
    size.maxCols = Math.max(size.maxCols, merge.e.c + 1);
};

const buildCsvCols = (rows: any[][], colCount: number) => {
    const cols: Record<number, { width: number }> = {};
    for (let i = 0; i < colCount; i += 1) {
        cols[i] = { width: calculateColWidth(rows, i) };
    }
    return cols;
};

const buildColsFromWorksheet = (worksheet: ExcelJS.Worksheet, colCount: number) => {
    const cols: Record<number, { width: number }> = {};
    for (let i = 1; i <= colCount; i += 1) {
        const width = excelColWidthToPx(worksheet.getColumn(i).width) ?? DEFAULT_COL_WIDTH;
        cols[i - 1] = { width: clampColWidth(width) };
    }
    return cols;
};

const formatCellText = (cell: ExcelJS.Cell) => {
    const raw = cell.value;
    if (raw && typeof raw === 'object' && 'hyperlink' in raw) {
        const hv = raw as ExcelJS.CellHyperlinkValue;
        return hv.text || hv.hyperlink || '';
    }
    if (cell.formula) return `=${cell.formula}`;
    const value = cell.value;
    if (value && typeof value === 'object' && 'formula' in value) {
        const formula = (value as { formula?: string }).formula;
        if (formula) return `=${formula}`;
    }
    if (cell.value == null) return '';
    if (cell.text) return cell.text;
    if (cell.value instanceof Date) {
        return cell.value.toISOString().slice(0, 10);
    }
    return String(cell.value);
};

const readFreezeFromWorksheet = (worksheet: ExcelJS.Worksheet): string | undefined => {
    const views = worksheet.views;
    if (!views?.length) return undefined;
    for (let i = 0; i < views.length; i += 1) {
        const view = views[i];
        if (view.state === 'frozen') {
            const xSplit = view.xSplit ?? 0;
            const ySplit = view.ySplit ?? 0;
            return excelFreezeToExpr(xSplit, ySplit);
        }
    }
    return undefined;
};

type ExcelJsSheetExtras = Pick<SheetData, 'freeze' | 'autofilter'>;

const readSheetExtras = (worksheet: ExcelJS.Worksheet): ExcelJsSheetExtras => {
    const extras: ExcelJsSheetExtras = {};
    const freeze = readFreezeFromWorksheet(worksheet);
    if (freeze) extras.freeze = freeze;
    const autofilter = readAutofilterRef(worksheet.autoFilter);
    if (autofilter) extras.autofilter = autofilter;
    return extras;
};

const applyRowHeight = (rows: RowMap, ri: number, excelRow: ExcelJS.Row) => {
    if (excelRow.height == null) return;
    const px = excelRowHeightToPx(excelRow.height);
    if (Math.abs(px - DEFAULT_ROW_HEIGHT_PX) < 1) return;
    const existing = rows[ri];
    if (existing && typeof existing === 'object' && 'cells' in existing) {
        existing.height = px;
    } else {
        rows[ri] = { cells: {}, height: px };
    }
};

const prepareExcelJsWorkbook = async (buffer: ArrayBuffer) => {
    const zipMetadata = inspectOoxmlZipCentralDirectory(buffer);
    await validateOoxmlZipInflationBounds(buffer, zipMetadata);
    const { default: JSZip } = await import('jszip');
    const zip = await JSZip.loadAsync(buffer);
    const sortStateXmlMap =
        new Map<number, ReturnType<typeof readWorksheetSortStateXml>>();
    let changed = false;
    const packageHasChartParts = Object.keys(zip.files)
        .some(name => /^xl\/charts\//i.test(name) && !zip.files[name].dir);

    for (const [name, entry] of Object.entries(zip.files)) {
        if (
            entry.dir
            || (!name.toLowerCase().endsWith('.xml')
                && !name.toLowerCase().endsWith('.rels'))
        ) {
            continue;
        }
        const metadata = zipMetadata.entries.get(name.toLowerCase());
        if (!metadata || metadata.directory) {
            throw new Error(`Partie OOXML absente du préflight ZIP : ${name}`);
        }
        const xml = await entry.async('string');
        const prefixed = normalizeSpreadsheetMlElementPrefixes(xml);
        const normalized = name.toLowerCase().endsWith('.rels')
            ? normalizeOoxmlRelationshipTargets(prefixed.xml, name)
            : { xml: prefixed.xml, changed: false };
        if (prefixed.changed || normalized.changed) {
            zip.file(name, normalized.xml);
            changed = true;
        }

        const worksheet = /^xl\/worksheets\/sheet(\d+)\.xml$/i.exec(name);
        if (worksheet) {
            sortStateXmlMap.set(
                Number(worksheet[1]) - 1,
                readWorksheetSortStateXml(normalized.xml),
            );
        }
    }

    let chartInventory: OoxmlChartInventoryResult = { sheets: [], warnings: [] };
    try {
        chartInventory = await readOoxmlChartInventory(zip);
    } catch {
        chartInventory = {
            sheets: [],
            warnings: ['L’inventaire OOXML des graphiques a échoué; leur conservation native reste obligatoire.'],
        };
    }

    const excelJsBuffer = changed
        ? await zip.generateAsync({
            type: 'arraybuffer',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 },
        })
        : buffer;
    return { excelJsBuffer, sortStateXmlMap, chartInventory, packageHasChartParts };
};

const convertExcelJsWorksheet = (
    worksheet: ExcelJS.Worksheet,
    workbook: ExcelJS.Workbook,
): Pick<
    SheetData,
    'rows' | 'cols' | 'styles' | 'merges' | 'freeze' | 'autofilter' | 'tables' | 'hyperlinks'
    | 'validations' | 'sheetProtection' | 'images' | 'backgroundImage' | 'comments'
    | 'conditionalFormattings' | 'pageSetup'
> => {
    const rows: RowMap = {};
    const styleRegistry = new StyleRegistry();
    const hyperlinkParts: Record<string, { link: string; tooltip?: string }>[] = [];
    const comments: Record<string, SheetCommentData> = {};
    const resolveColor = createExcelColorResolver(workbook);
    const sheetProtected = isWorksheetProtected(worksheet);
    const sheetProtection = readWorksheetProtection(worksheet);
    let maxCols = 0;
    let maxRow = 0;

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (!row) return;
        const ri = rowNumber - 1;
        applyRowHeight(rows, ri, row);
        if (row.height != null && ri + 1 > maxRow) maxRow = ri + 1;
        if (row.cellCount === 0) return;
        const cells: Record<number, CellData> = {};
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            if (cell.isMerged && cell.address !== cell.master.address) return;

            const ci = colNumber - 1;
            const text = formatCellText(cell);
            const cellStyle = excelJsCellToStyle(cell, resolveColor);
            const editable = readCellEditableFromExcel(cell, sheetProtected);
            const hl = readCellHyperlink(cell, ri, ci);
            const comment = noteToComment(cell.note);
            const hasHyperlink = Object.keys(hl).length > 0;
            if (comment) comments[cell.address] = comment;
            if (!text && !cellStyle && editable === undefined && !hasHyperlink && !comment) return;

            const styleIndex = styleRegistry.add(cellStyle);
            const cellData: CellData = { text };
            if (styleIndex != null) cellData.style = styleIndex;
            if (editable !== undefined) cellData.editable = editable;
            const cachedResult = formulaResult(cell);
            if (cachedResult !== undefined) cellData.formulaResult = cachedResult;
            cells[ci] = cellData;
            if (ci + 1 > maxCols) maxCols = ci + 1;
            if (ri + 1 > maxRow) maxRow = ri + 1;
            if (hasHyperlink) hyperlinkParts.push(hl);
        });
        if (Object.keys(cells).length > 0) {
            const existing = rows[ri];
            rows[ri] = existing && typeof existing === 'object' && 'height' in existing
                ? { cells, height: existing.height }
                : { cells };
        }
    });

    const merges = readWorksheetMerges(worksheet);
    const tables = readWorksheetTables(worksheet);
    const sheetSize = { maxRow, maxCols };
    merges.forEach(merge => expandSizeForMerge(merge, sheetSize));
    maxRow = sheetSize.maxRow;
    maxCols = sheetSize.maxCols;

    const colCount = Math.max(maxCols, worksheet.columnCount || 0);
    const cols = buildColsFromWorksheet(worksheet, colCount);
    const styles = styleRegistry.getStyles();
    const sheetExtras = readSheetExtras(worksheet);
    const hyperlinks = mergeHyperlinkMaps(...hyperlinkParts);
    const validations = readWorksheetValidations(worksheet);
    const images = readWorksheetImages(worksheet, workbook);
    const backgroundImage = readWorksheetBackgroundImage(worksheet, workbook);
    const conditionalFormattings = readConditionalFormattings(worksheet, resolveColor);
    const pageSetup = cloneJson(worksheet.pageSetup ?? {});

    return {
        rows: { len: maxRow, ...rows },
        cols: { len: colCount, ...cols },
        styles: styles.length > 0 ? styles : undefined,
        merges: merges.length > 0 ? merges : undefined,
        ...(tables.length ? { tables } : {}),
        ...(Object.keys(hyperlinks).length ? { hyperlinks } : {}),
        ...(validations.length ? { validations } : {}),
        ...(sheetProtection ? { sheetProtection } : {}),
        ...(images.length ? { images } : {}),
        ...(backgroundImage ? { backgroundImage } : {}),
        ...(Object.keys(comments).length ? { comments } : {}),
        ...(conditionalFormattings.length ? { conditionalFormattings } : {}),
        ...(Object.keys(pageSetup).length ? { pageSetup } : {}),
        ...sheetExtras,
    };
};

const convertExcelJsWorkbook = (
    workbook: ExcelJS.Workbook,
    sortStateXmlMap?: Map<number, ReturnType<typeof readWorksheetSortStateXml>>,
    chartInventory?: OoxmlChartInventoryResult,
    packageHasChartParts = false,
): ExcelData => {
    const sheets: SheetData[] = [];
    let maxLength = 0;
    let maxCols = 26;

    const chartsBySheetName = new Map(
        (chartInventory?.sheets ?? []).map(item => [item.sheetName.toLocaleLowerCase(), item] as const),
    );
    const mappedChartParts = (chartInventory?.sheets ?? []).some(item => item.hasChartParts);

    workbook.worksheets.forEach((worksheet, index) => {
        const converted = convertExcelJsWorksheet(worksheet, workbook);
        const chartState = chartsBySheetName.get(worksheet.name.toLocaleLowerCase());
        const fallbackUnmappedChartParts = packageHasChartParts && !mappedChartParts && index === 0;
        const xmlAutofilter = sortStateXmlMap?.get(index);
        if (xmlAutofilter?.sort && converted.autofilter?.ref) {
            converted.autofilter.sort = xmlAutofilter.sort;
        }
        const rowCount = converted.rows?.len ?? 0;
        if (maxLength < rowCount) maxLength = rowCount;

        const colLen = converted.cols?.len ?? 0;
        if (colLen > maxCols) maxCols = colLen;

        sheets.push({
            name: worksheet.name,
            rows: converted.rows,
            cols: converted.cols,
            ...(converted.styles ? { styles: converted.styles } : {}),
            ...(converted.merges ? { merges: converted.merges } : {}),
            ...(converted.freeze ? { freeze: converted.freeze } : {}),
            ...(converted.autofilter ? { autofilter: converted.autofilter } : {}),
            ...(converted.tables ? { tables: converted.tables } : {}),
            ...(chartState?.charts.length ? { charts: chartState.charts } : {}),
            ...((chartState?.hasChartParts || fallbackUnmappedChartParts)
                ? { hasNativeChartParts: true }
                : {}),
            ...((chartState?.unsupportedChartCount || fallbackUnmappedChartParts)
                ? { unsupportedNativeChartCount: chartState?.unsupportedChartCount || 1 }
                : {}),
            ...(converted.hyperlinks ? { hyperlinks: converted.hyperlinks } : {}),
            ...(converted.validations ? { validations: converted.validations } : {}),
            ...(converted.sheetProtection ? { sheetProtection: converted.sheetProtection } : {}),
            ...(converted.images ? { images: converted.images } : {}),
            ...(converted.backgroundImage ? { backgroundImage: converted.backgroundImage } : {}),
            ...(converted.comments ? { comments: converted.comments } : {}),
            ...(converted.conditionalFormattings
                ? { conditionalFormattings: converted.conditionalFormattings }
                : {}),
            ...(converted.pageSetup ? { pageSetup: converted.pageSetup } : {}),
        });
    });

    return { sheets, maxLength, maxCols };
};

const loadWithExcelJs = async (buffer: ArrayBuffer): Promise<ExcelData> => {
    const [prepared, { default: ExcelJSRuntime }] = await Promise.all([
        prepareExcelJsWorkbook(buffer),
        import('@cweijan/exceljs'),
    ]);
    const workbook = new ExcelJSRuntime.Workbook();
    await workbook.xlsx.load(prepared.excelJsBuffer);
    return convertExcelJsWorkbook(
        workbook,
        prepared.sortStateXmlMap,
        prepared.chartInventory,
        prepared.packageHasChartParts,
    );
};

const sheetJsColWidthToPx = (col?: XLSX.ColInfo) => {
    if (!col) return null;
    if (col.wpx != null) return col.wpx;
    if (col.wch != null) return col.wch * CHAR_WIDTH;
    if (col.width != null) return col.width * CHAR_WIDTH;
    return null;
};

const buildColsFromSheetJsWorksheet = (worksheet: XLSX.WorkSheet, colCount: number) => {
    const cols: Record<number, { width: number }> = {};
    const sheetCols = worksheet['!cols'];
    for (let i = 0; i < colCount; i += 1) {
        const width = sheetJsColWidthToPx(sheetCols?.[i]) ?? DEFAULT_COL_WIDTH;
        cols[i] = { width: clampColWidth(width) };
    }
    return cols;
};

const formatSheetJsCell = (cell: XLSX.CellObject) => {
    if (cell.w) return cell.w;
    if (cell.v == null) return '';
    if (cell.v instanceof Date) return cell.v.toISOString().slice(0, 10);
    return String(cell.v);
};

const convertSheetJsWorksheet = (
    worksheet: XLSX.WorkSheet,
    utils: SheetJsUtils,
): Pick<SheetData, 'rows' | 'cols' | 'merges'> => {
    const rows: RowMap = {};
    let maxCols = 0;
    let maxRow = 0;
    const ref = worksheet['!ref'];
    if (!ref) {
        return { rows: { len: 0 }, cols: { len: 0 } };
    }

    const range = utils.decode_range(ref);
    for (let ri = range.s.r; ri <= range.e.r; ri += 1) {
        const cells: Record<number, CellData> = {};
        let hasContent = false;
        for (let ci = range.s.c; ci <= range.e.c; ci += 1) {
            const addr = utils.encode_cell({ r: ri, c: ci });
            const cell = worksheet[addr];
            if (!cell) continue;
            const text = formatSheetJsCell(cell);
            if (!text) continue;
            cells[ci] = { text };
            hasContent = true;
            if (ci + 1 > maxCols) maxCols = ci + 1;
            if (ri + 1 > maxRow) maxRow = ri + 1;
        }
        if (hasContent) {
            rows[ri] = { cells };
        }
    }

    const sheetSize = { maxRow, maxCols };
    (worksheet['!merges'] ?? []).forEach(merge => expandSizeForSheetJsMerge(merge, sheetSize));
    maxRow = sheetSize.maxRow;
    maxCols = sheetSize.maxCols;

    const colCount = Math.max(maxCols, range.e.c - range.s.c + 1);
    const merges = readSheetJsMerges(worksheet, utils);
    return {
        rows: { len: maxRow, ...rows },
        cols: { len: colCount, ...buildColsFromSheetJsWorksheet(worksheet, colCount) },
        merges: merges.length > 0 ? merges : undefined,
    };
};

const convertSheetJsWorkbook = (workbook: XLSX.WorkBook, utils: SheetJsUtils): ExcelData => {
    const sheets: SheetData[] = [];
    let maxLength = 0;
    let maxCols = 26;

    for (const sheetName of workbook.SheetNames) {
        const converted = convertSheetJsWorksheet(workbook.Sheets[sheetName], utils);
        const rowCount = converted.rows?.len ?? 0;
        if (maxLength < rowCount) maxLength = rowCount;

        const colLen = converted.cols?.len ?? 0;
        if (colLen > maxCols) maxCols = colLen;

        sheets.push({
            name: sheetName,
            rows: converted.rows,
            cols: converted.cols,
            ...(converted.merges ? { merges: converted.merges } : {}),
        });
    }

    return { sheets, maxLength, maxCols };
};

const loadWithSheetJs = async (buffer: ArrayBuffer): Promise<ExcelData> => {
    const XLSXRuntime = await import('xlsx');
    const workbook = XLSXRuntime.read(buffer, { type: 'array', cellDates: true });
    return convertSheetJsWorkbook(workbook, XLSXRuntime.utils);
};

const loadCsv = async (buffer: ArrayBuffer): Promise<ExcelData> => {
    let maxCols = 26;
    const emptySheet = { maxCols, sheets: [{ name: 'Sheet1', rows: { len: 0 } }] };
    let csvStr = decodeCsvBuffer(buffer);
    if (!csvStr) return emptySheet;

    try {
        const leadingEmptyRows = csvStr.match(/^(?:\r\n|\n|\r)+/)?.[0].match(/\r\n|\n|\r/g)?.length ?? 0;
        const csvToParse = leadingEmptyRows > 0 ? csvStr.replace(/^(?:\r\n|\n|\r)+/, '') : csvStr;
        if (!csvToParse) {
            return {
                maxCols,
                maxLength: leadingEmptyRows,
                sheets: [{
                    name: 'Sheet1',
                    rows: { len: leadingEmptyRows },
                }],
            };
        }
        let parseInput = csvToParse;
        if (!parseInput.includes('\n')) parseInput += '\n';
        const { inferSchema, initParser } = await import('udsv');
        const schema = inferSchema(parseInput, { header: () => [] });
        const rows = initParser(schema).stringArrs(parseInput);
        const colCount = rows.reduce((max, row) => Math.max(max, row.length), 0);

        const processedRows: RowMap = {};
        for (let i = 0; i < leadingEmptyRows; i += 1) {
            processedRows[i] = { cells: {} };
        }
        for (let i = 0; i < rows.length; i += 1) {
            const row = rows[i];
            const cells: Record<number, CellData> = {};
            for (let j = 0; j < row.length; j += 1) {
                cells[j] = { text: row[j] == null ? '' : String(row[j]) };
                if (j + 1 > maxCols) maxCols = j + 1;
            }
            processedRows[i + leadingEmptyRows] = { cells };
        }
        const csvRows = [
            ...Array.from({ length: leadingEmptyRows }, () => [] as string[]),
            ...rows,
        ];

        return {
            maxCols,
            maxLength: csvRows.length,
            csvDelimiter: schema.col,
            sheets: [{
                name: 'Sheet1',
                rows: { len: csvRows.length, ...processedRows },
                cols: { len: colCount, ...buildCsvCols(csvRows, colCount) },
            }],
        };
    } catch (error) {
        console.error(error);
        return { maxCols, sheets: [{ name: 'Sheet1', rows: { len: 1, 0: { cells: { 0: { text: error.message } } } } }] };
    }
};

const isCsvExt = (ext: string) => /csv|tsv/.test(ext.toLowerCase());
const isOdsExt = (ext: string) => ext.toLowerCase().includes('ods');
const isXlsExt = (ext: string) => ext.toLowerCase().replace(/^\./, '') === 'xls';

export async function loadSheets(buffer: ArrayBuffer, ext: string): Promise<ExcelData> {
    if (isCsvExt(ext)) {
        return loadCsv(buffer);
    }
    if (isXlsExt(ext) || isOdsExt(ext)) {
        return loadWithSheetJs(buffer);
    }
    return loadWithExcelJs(buffer);
}

export async function readCSV(buffer: ArrayBuffer): Promise<ExcelData> {
    return loadCsv(buffer);
}

export async function readExcel(buffer: ArrayBuffer): Promise<ExcelData> {
    return loadWithExcelJs(buffer);
}
