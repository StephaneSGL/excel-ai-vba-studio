import ExcelJS from '@cweijan/exceljs';
import JSZip from 'jszip';
import {
    buildExcelTableStyleCatalog,
    excelTableNameComparisonKey,
    isValidExcelTableName,
	minimumExcelTableRangeRows,
    normalizeA1Range,
    normalizeExcelTableName,
    SIMPLE_A1_RANGE,
    type SheetTableData,
} from '../../../common/excelWorkbookObjects';
import { handler } from "../../util/vscode";
import * as XLSX from 'xlsx';
import type Spreadsheet from './x-spreadsheet/index';
import type { CellData, RowData, SheetData } from './x-spreadsheet/index';
import { CsvEncoding, encodeCsvText } from './csvEncoding';
import { DEFAULT_ROW_HEIGHT_PX, freezeExprToExcelView, pxToExcelRowHeight } from './excel_meta';
import { patchWorksheetSortStateXml } from './excel_sort_state';
import { applySpreadsheetStyle } from './excel_styles';
import { hyperlinkKey, writeCellHyperlink, type SpreadsheetHyperlink } from './excel_hyperlink';
import { writeWorksheetValidations } from './excel_validation';
import { writeWorksheetProtection } from './excel_protection';
import { writeWorksheetImages } from './excel_images';

const DEFAULT_COL_WIDTH = 100;
const EXCEL_TABLE_STYLES = new Set(buildExcelTableStyleCatalog());

interface ParsedTableRange {
    ref: string;
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
}

export { buildFormattingSnapshot, hasFormattingChanged } from './excel_meta';

export interface ExportOptions {
    /** 通过另存为对话框保存，而非覆盖当前文件 */
    saveAs?: boolean;
    /** saveAs 时指定目标格式 */
    saveAsExt?: string;
}

function isRowData(row: RowData | number | undefined): row is RowData {
    return row != null && typeof row === 'object';
}

function excelColumnNumber(letters: string): number {
    let result = 0;
    for (const letter of letters) {
        result = result * 26 + letter.charCodeAt(0) - 64;
    }
    return result;
}

function parseTableRange(rangeRef: string): ParsedTableRange {
    const ref = normalizeA1Range(rangeRef);
    if (!SIMPLE_A1_RANGE.test(ref)) {
        throw new Error(`Invalid Excel table range: ${rangeRef}`);
    }
    const [start, rawEnd = start] = ref.split(':');
    const parseAddress = (address: string) => {
        const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(address);
        if (!match) throw new Error(`Invalid Excel table address: ${address}`);
        return { row: Number(match[2]), col: excelColumnNumber(match[1]) };
    };
    const first = parseAddress(start);
    const last = parseAddress(rawEnd);
    if (
        first.row > last.row || first.col > last.col
        || last.row > 1_048_576 || last.col > 16_384
    ) {
        throw new Error(`Excel table range is outside worksheet limits: ${rangeRef}`);
    }
    return {
        ref,
        startRow: first.row,
        startCol: first.col,
        endRow: last.row,
        endCol: last.col,
    };
}

function rangesOverlap(left: ParsedTableRange, right: ParsedTableRange): boolean {
    return left.startRow <= right.endRow
        && left.startCol <= right.endCol
        && right.startRow <= left.endRow
        && right.startCol <= left.endCol;
}

function assertTableName(value: string, label: string): string {
    const name = normalizeExcelTableName(value);
    if (!isValidExcelTableName(name)) {
        throw new Error(`Invalid Excel table ${label}: ${value}`);
    }
    return name;
}

function validateWorkbookTables(sheets: SheetData[]): void {
    const workbookNames = new Set<string>();
    sheets.forEach((sheet, sheetIndex) => {
        const parsedTables: Array<{ table: SheetTableData; range: ParsedTableRange }> = [];
        (sheet.tables ?? []).forEach(table => {
            const range = parseTableRange(table.rangeRef);
            const rowCount = range.endRow - range.startRow + 1;
            const minimumRows = minimumExcelTableRangeRows(table.totalsRow);
            if (rowCount < minimumRows) {
                throw new Error(`Excel table ${table.name} does not contain a data row.`);
            }
            const names = [
                assertTableName(table.name, 'name'),
                assertTableName(table.displayName || table.name, 'display name'),
            ];
            for (const name of new Set(names.map(excelTableNameComparisonKey))) {
                if (workbookNames.has(name)) {
                    throw new Error(`Duplicate Excel table name: ${table.name}`);
                }
                workbookNames.add(name);
            }
            if (!EXCEL_TABLE_STYLES.has(table.style.name)) {
                throw new Error(`Unsupported Excel table style: ${table.style.name}`);
            }
            for (const existing of parsedTables) {
                if (rangesOverlap(range, existing.range)) {
                    throw new Error(
                        `Excel tables ${existing.table.name} and ${table.name} overlap on `
                        + `${sheet.name || `Sheet${sheetIndex + 1}`}.`
                    );
                }
            }
            parsedTables.push({ table, range });
        });
    });
}

function uniqueColumnName(value: string, fallback: string, usedNames: Set<string>): string {
    const base = value.trim() || fallback;
    let candidate = base;
    let suffix = 2;
    while (usedNames.has(candidate.toLocaleLowerCase())) {
        candidate = `${base}_${suffix}`;
        suffix += 1;
    }
    usedNames.add(candidate.toLocaleLowerCase());
    return candidate;
}

function cellTextForTableColumn(cell: ExcelJS.Cell): string {
    if (cell.text) return cell.text;
    if (cell.value == null) return '';
    if (cell.value instanceof Date) return cell.value.toISOString();
    if (typeof cell.value === 'object' && 'formula' in cell.value) {
        return String((cell.value as ExcelJS.CellFormulaValue).result ?? '');
    }
    return String(cell.value);
}

function writeWorksheetTables(worksheet: ExcelJS.Worksheet, tables: SheetTableData[] | undefined): void {
    for (const table of tables ?? []) {
        const name = assertTableName(table.name, 'name');
        const displayName = assertTableName(table.displayName || table.name, 'display name');
        const range = parseTableRange(table.rangeRef);
        const columnCount = range.endCol - range.startCol + 1;
        const headerOffset = table.headerRow ? 1 : 0;
        const totalsOffset = table.totalsRow ? 1 : 0;
        const dataStartRow = range.startRow + headerOffset;
        const dataEndRow = range.endRow - totalsOffset;
        const usedColumnNames = new Set<string>();
        const columns: ExcelJS.TableColumnProperties[] = [];

        for (let offset = 0; offset < columnCount; offset += 1) {
            const col = range.startCol + offset;
            const sourceName = table.headerRow
                ? cellTextForTableColumn(worksheet.getCell(range.startRow, col))
                : '';
            const column: ExcelJS.TableColumnProperties = {
                name: uniqueColumnName(sourceName, `Column${offset + 1}`, usedColumnNames),
                filterButton: table.headerRow,
            };
            if (table.totalsRow) {
                const totalsCell = worksheet.getCell(range.endRow, col);
                if (offset === 0) {
                    column.totalsRowLabel = cellTextForTableColumn(totalsCell) || 'Total';
                } else if (
                    totalsCell.value
                    && typeof totalsCell.value === 'object'
                    && 'formula' in totalsCell.value
                ) {
                    const formulaValue = totalsCell.value as ExcelJS.CellFormulaValue;
                    column.totalsRowFunction = 'custom';
                    column.totalsRowFormula = formulaValue.formula;
                } else {
                    column.totalsRowFunction = 'none';
                }
            }
            columns.push(column);
        }

        const rows: ExcelJS.CellValue[][] = [];
        for (let row = dataStartRow; row <= dataEndRow; row += 1) {
            rows.push(Array.from({ length: columnCount }, (_, offset) => (
                worksheet.getCell(row, range.startCol + offset).value
            )));
        }

        worksheet.addTable({
            name,
            displayName,
            ref: worksheet.getCell(range.startRow, range.startCol).address,
            headerRow: table.headerRow,
            totalsRow: table.totalsRow,
            style: {
                theme: table.style.name as ExcelJS.TableStyleProperties['theme'],
                showFirstColumn: table.style.showFirstColumn,
                showLastColumn: table.style.showLastColumn,
                showRowStripes: table.style.showRowStripes,
                showColumnStripes: table.style.showColumnStripes,
            },
            columns,
            rows,
        });
    }
}

function getColWidth(cols: SheetData['cols'], ci: number) {
    const col = cols?.[ci];
    if (col && typeof col === 'object' && col.width != null) return col.width;
    return DEFAULT_COL_WIDTH;
}

function pxToExcelColWidth(px: number) {
    return Math.max((px - 5) / 7, 0);
}

function setCellValue(
    cell: ExcelJS.Cell,
    text: string,
    cachedResult?: CellData['formulaResult'],
) {
    if (!text) {
        cell.value = null;
        return;
    }
    if (text.startsWith('=')) {
        cell.value = {
            formula: text.slice(1),
            ...(cachedResult !== undefined ? { result: cachedResult } : {}),
        };
        return;
    }
    const num = Number(text);
    if (text.trim() !== '' && !Number.isNaN(num) && String(num) === text.trim()) {
        cell.value = num;
        return;
    }
    cell.value = text;
}

function applySheetMeta(worksheet: ExcelJS.Worksheet, sheetData: SheetData) {
    if (sheetData.pageSetup) {
        Object.assign(worksheet.pageSetup, sheetData.pageSetup);
    }
    if (sheetData.freeze) {
        const frozen = freezeExprToExcelView(sheetData.freeze);
        if (frozen) {
            worksheet.views = [{
                state: 'frozen',
                xSplit: frozen.xSplit,
                ySplit: frozen.ySplit,
                topLeftCell: sheetData.freeze,
            }];
        }
    }

    const autofilter = sheetData.autofilter;
    if (autofilter?.ref) {
        worksheet.autoFilter = autofilter.ref;
    }
}

function writeConditionalFormattings(
    worksheet: ExcelJS.Worksheet,
    conditionalFormattings: SheetData['conditionalFormattings'],
) {
    conditionalFormattings?.forEach(item => {
        const rules = item.rules.map(rule => {
            const { displayStyle: _displayStyle, ...excelRule } = rule;
            return excelRule;
        });
        worksheet.addConditionalFormatting({
            ref: item.ref,
            rules,
        } as ExcelJS.ConditionalFormattingOptions);
    });
}

function writeComments(worksheet: ExcelJS.Worksheet, comments: SheetData['comments']) {
    Object.entries(comments ?? {}).forEach(([address, comment]) => {
        if (!comment?.text) return;
        worksheet.getCell(address).note = comment.text;
    });
}

function writeRowHeights(worksheet: ExcelJS.Worksheet, rows: SheetData['rows']) {
    if (!rows?.len) return;
    for (let ri = 0; ri < rows.len; ri += 1) {
        const row = rows[ri];
        if (!isRowData(row) || row.height == null) continue;
        if (Math.abs(row.height - DEFAULT_ROW_HEIGHT_PX) < 1) continue;
        worksheet.getRow(ri + 1).height = pxToExcelRowHeight(row.height);
    }
}

async function writeSheetToExcelJs(worksheet: ExcelJS.Worksheet, workbook: ExcelJS.Workbook, sheetData: SheetData) {
    const { rows, cols, styles = [], merges = [], hyperlinks = {}, validations } = sheetData;

    if (cols?.len) {
        for (let ci = 0; ci < cols.len; ci += 1) {
            worksheet.getColumn(ci + 1).width = pxToExcelColWidth(getColWidth(cols, ci));
        }
    }

    for (let i = 0; i < merges.length; i += 1) {
        worksheet.mergeCells(merges[i]);
    }

    applySheetMeta(worksheet, sheetData);
    writeRowHeights(worksheet, rows);
    writeWorksheetValidations(worksheet, validations);
    writeConditionalFormattings(worksheet, sheetData.conditionalFormattings);
    writeComments(worksheet, sheetData.comments);

    if (rows) {
        const rowLen = rows.len ?? 0;
        for (let ri = 0; ri < rowLen; ri += 1) {
            const row = rows[ri];
            if (!isRowData(row) || !row.cells) continue;
            for (const ciKey of Object.keys(row.cells)) {
                const ci = Number(ciKey);
                if (Number.isNaN(ci)) continue;
                const cellData = row.cells[ci];
                const excelCell = worksheet.getCell(ri + 1, ci + 1);
                const hl = hyperlinks[hyperlinkKey(ri, ci)] as SpreadsheetHyperlink | undefined;
                if (hl?.link) {
                    writeCellHyperlink(excelCell, cellData.text ?? '', hl);
                } else {
                    setCellValue(excelCell, cellData.text ?? '', cellData.formulaResult);
                }
                if (cellData.style != null && styles[cellData.style]) {
                    applySpreadsheetStyle(excelCell, styles[cellData.style]);
                }
            }
        }
    }

    writeWorksheetTables(worksheet, sheetData.tables);
    writeWorksheetImages(worksheet, workbook, sheetData.images, sheetData.backgroundImage);
    await writeWorksheetProtection(worksheet, sheetData);
}

async function emitSave(buffer: Uint8Array, options?: ExportOptions) {
    const content = [...buffer];
    if (options?.saveAs) {
        handler.emit('saveAs', { content, ext: options.saveAsExt ?? 'xlsx' });
        return;
    }
    handler.emit('save', content);
}

async function patchWorkbookSortStates(buffer: Uint8Array, sheets: SheetData[]) {
    const zip = await JSZip.loadAsync(buffer);
    for (let i = 0; i < sheets.length; i += 1) {
        const file = zip.file(`xl/worksheets/sheet${i + 1}.xml`);
        if (!file) continue;
        const xml = await file.async('string');
        zip.file(`xl/worksheets/sheet${i + 1}.xml`, patchWorksheetSortStateXml(xml, sheets[i].autofilter));
    }
    return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
}

export async function buildExcelWorkbookBuffer(sheets: SheetData[]): Promise<Uint8Array> {
    validateWorkbookTables(sheets);
    const workbook = new ExcelJS.Workbook();
    for (let i = 0; i < sheets.length; i += 1) {
        const sheetData = sheets[i];
        const worksheet = workbook.addWorksheet(sheetData.name || `Sheet${i + 1}`);
        await writeSheetToExcelJs(worksheet, workbook, sheetData);
    }
    const buffer = new Uint8Array(await workbook.xlsx.writeBuffer());
    return patchWorkbookSortStates(buffer, sheets);
}

async function exportWithExcelJs(sheets: SheetData[], options?: ExportOptions) {
    const buffer = await buildExcelWorkbookBuffer(sheets);
    await emitSave(buffer, options);
}

function applyColWidths(ws: XLSX.WorkSheet, xws: SheetData) {
    const cols = xws.cols;
    if (!cols?.len) return;
    const colWidths = [];
    for (let ci = 0; ci < cols.len; ci += 1) {
        colWidths.push({ wpx: getColWidth(cols, ci) });
    }
    ws['!cols'] = colWidths;
}

function dataToSheetJs(xws: SheetData) {
    const aoa: string[][] = [];
    const rowobj = xws.rows;
    if (!rowobj?.len) {
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        applyColWidths(ws, xws);
        return ws;
    }
    for (let ri = 0; ri < rowobj.len; ri += 1) {
        const row = rowobj[ri];
        if (!isRowData(row)) continue;
        aoa[ri] = [];
        for (const ciKey of Object.keys(row.cells ?? {})) {
            const ci = Number(ciKey);
            if (Number.isNaN(ci)) continue;
            aoa[ri][ci] = row.cells[ci].text;
        }
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    applyColWidths(ws, xws);
    return ws;
}

function exportWithSheetJs(sheets: SheetData[], bookType: XLSX.BookType) {
    const workbook = XLSX.utils.book_new();
    for (let i = 0; i < sheets.length; i += 1) {
        const sheetData = sheets[i];
        XLSX.utils.book_append_sheet(workbook, dataToSheetJs(sheetData), sheetData.name || `Sheet${i + 1}`);
    }
    const buffer = XLSX.write(workbook, { bookType, type: 'array' });
    handler.emit('save', [...new Uint8Array(buffer)]);
}

export async function exportSaveAs(
    spreadSheet: Spreadsheet,
    targetExt: string,
    csvEncoding: CsvEncoding = 'utf8',
    csvDelimiter: string = ',',
) {
    const ext = targetExt.replace('.', '').toLowerCase();
    const sheets = spreadSheet.getData();
    if (ext === 'xlsx' || ext === 'xlsm') {
        await exportWithExcelJs(sheets, { saveAs: true, saveAsExt: ext });
        return;
    }
    if (ext === 'xls' || ext === 'ods') {
        const wb = XLSX.utils.book_new();
        sheets.forEach((s, i) => {
            const ws = dataToSheetJs(s);
            XLSX.utils.book_append_sheet(wb, ws, s.name || `Sheet${i + 1}`);
        });
        const buf = XLSX.write(wb, { bookType: ext as XLSX.BookType, type: 'array' });
        handler.emit('saveAs', { content: [...new Uint8Array(buf)], ext });
        return;
    }
    if (ext === 'csv' || ext === 'tsv') {
        const fs = ext === 'tsv' ? '\t' : csvDelimiter;
        const csvContent = XLSX.utils.sheet_to_csv(dataToSheetJs(sheets[0]), { FS: fs });
        const bytes = encodeCsvText(csvContent, csvEncoding);
        handler.emit('saveAs', { content: [...bytes], ext });
        return;
    }
    throw new Error(`Unsupported spreadsheet save format: ${ext || '(none)'}`);
}

export async function export_xlsx(
    spreadSheet: Spreadsheet,
    extName: string,
    csvEncoding: CsvEncoding = 'utf8',
    options?: ExportOptions,
    csvDelimiter: string = ',',
) {
    const ext = extName.replace('.', '').toLowerCase();
    const sheets = spreadSheet.getData();

    if (ext === 'xlsx' || ext === 'xlsm' || options?.saveAs) {
        await exportWithExcelJs(sheets, options?.saveAs ? { saveAs: true } : undefined);
        return;
    }
    if (ext === 'xls' || ext === 'ods') {
        exportWithSheetJs(sheets, ext);
        return;
    }
    if (ext === 'csv' || ext === 'tsv') {
        const fs = ext === 'tsv' ? '\t' : csvDelimiter;
        const csvContent = XLSX.utils.sheet_to_csv(dataToSheetJs(sheets[0]), { FS: fs });
        const bytes = encodeCsvText(csvContent, csvEncoding);
        handler.emit('save', [...bytes]);
    }
}
