import type {
    NativeExcelCellEdit,
    NativeExcelCellValue,
    NativeExcelStylePatch,
} from '@/common/nativeExcelEdits';
import type {
    CellData,
    CellStyle,
    RowData,
    SheetData,
} from './x-spreadsheet/index';

type CellPosition = { row: number; column: number };

export interface NativeExcelEditPlan {
    operations: NativeExcelCellEdit[];
    unsupportedChanges: string[];
}

function isRowData(value: RowData | number | undefined): value is RowData {
    return Boolean(value && typeof value === 'object' && 'cells' in value);
}

function cellPositions(sheet: SheetData | undefined): Map<string, CellPosition> {
    const result = new Map<string, CellPosition>();
    if (!sheet?.rows) {
        return result;
    }
    for (const [rowKey, rowValue] of Object.entries(sheet.rows)) {
        if (!isRowData(rowValue)) {
            continue;
        }
        const row = Number(rowKey);
        for (const columnKey of Object.keys(rowValue.cells || {})) {
            const column = Number(columnKey);
            if (Number.isInteger(row) && Number.isInteger(column)) {
                result.set(`${row}:${column}`, { row, column });
            }
        }
    }
    return result;
}

function getCell(
    sheet: SheetData | undefined,
    row: number,
    column: number
): CellData | undefined {
    const rowData = sheet?.rows?.[row];
    return isRowData(rowData) ? rowData.cells?.[column] : undefined;
}

function getStyle(
    sheet: SheetData | undefined,
    cell: CellData | undefined
): CellStyle | undefined {
    return cell?.style === undefined ? undefined : sheet?.styles?.[cell.style];
}

function sameValue(left: unknown, right: unknown): boolean {
    return JSON.stringify(left) === JSON.stringify(right);
}

function sheetFeatureSnapshot(sheet: SheetData): Record<string, unknown> {
    const {
        name: _name,
        styles: _styles,
        rows: _rows,
        cols: _cols,
        ...features
    } = sheet;
    return features;
}

function columnDimensionSnapshot(sheet: SheetData): Record<string, unknown> {
    const { len: _len, ...dimensions } = sheet.cols ?? {};
    return dimensions;
}

function rowIndexes(sheet: SheetData | undefined): Set<number> {
    const indexes = new Set<number>();
    for (const [key, value] of Object.entries(sheet?.rows ?? {})) {
        if (key !== 'len' && Number.isInteger(Number(key)) && isRowData(value)) {
            indexes.add(Number(key));
        }
    }
    return indexes;
}

function toNativeValue(text: string | undefined): NativeExcelCellValue {
    if (!text) {
        return { kind: 'blank' };
    }
    if (text.startsWith('=')) {
        return { kind: 'formula', value: text };
    }
    const trimmed = text.trim();
    const numericLiteral =
        /^[+-]?(?:(?:0|[1-9]\d*)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;
    if (numericLiteral.test(trimmed)) {
        const number = Number(trimmed);
        if (!Number.isNaN(number) && Number.isFinite(number)) {
            return { kind: 'number', value: number };
        }
    }
    return { kind: 'text', value: text };
}

function buildStylePatch(
    before: CellStyle | undefined,
    after: CellStyle | undefined
): NativeExcelStylePatch | undefined {
    const patch: NativeExcelStylePatch = {};
    const scalarKeys = [
        'align',
        'valign',
        'bgcolor',
        'color',
        'format',
        'textwrap',
        'underline',
        'strike',
    ] as const;

    for (const key of scalarKeys) {
        if (!sameValue(before?.[key], after?.[key])) {
            (patch as Record<string, unknown>)[key] = after?.[key] ?? null;
        }
    }

    const fontPatch: NonNullable<NativeExcelStylePatch['font']> = {};
    for (const key of ['name', 'size', 'bold', 'italic'] as const) {
        if (!sameValue(before?.font?.[key], after?.font?.[key])) {
            (fontPatch as Record<string, unknown>)[key] =
                after?.font?.[key] ?? null;
        }
    }
    if (Object.keys(fontPatch).length > 0) {
        patch.font = fontPatch;
    }

    const borderPatch: NonNullable<NativeExcelStylePatch['border']> = {};
    for (const key of ['top', 'right', 'bottom', 'left'] as const) {
        if (!sameValue(before?.border?.[key], after?.border?.[key])) {
            borderPatch[key] = after?.border?.[key] ?? null;
        }
    }
    if (Object.keys(borderPatch).length > 0) {
        patch.border = borderPatch;
    }

    return Object.keys(patch).length > 0 ? patch : undefined;
}

export function buildNativeExcelEditPlan(
    beforeSheets: SheetData[],
    afterSheets: SheetData[]
): NativeExcelEditPlan {
    const unsupportedChanges = new Set<string>();
    const beforeNames = beforeSheets.map(sheet => sheet.name ?? '');
    const afterNames = afterSheets.map(sheet => sheet.name ?? '');
    if (!sameValue(beforeNames, afterNames)) {
        unsupportedChanges.add('worksheets');
    }

    const beforeByName = new Map(
        beforeSheets
            .filter((sheet): sheet is SheetData & { name: string } =>
                Boolean(sheet.name)
            )
            .map(sheet => [sheet.name, sheet] as const)
    );
    const operations: NativeExcelCellEdit[] = [];

    for (const afterSheet of afterSheets) {
        if (!afterSheet.name) {
            unsupportedChanges.add('worksheets');
            continue;
        }
        const beforeSheet = beforeByName.get(afterSheet.name);
        if (!beforeSheet) {
            continue;
        }

        if (
            !sameValue(
                sheetFeatureSnapshot(beforeSheet),
                sheetFeatureSnapshot(afterSheet)
            )
        ) {
            unsupportedChanges.add(`${afterSheet.name}:worksheet-features`);
        }
        if (
            !sameValue(
                columnDimensionSnapshot(beforeSheet),
                columnDimensionSnapshot(afterSheet)
            )
        ) {
            unsupportedChanges.add(`${afterSheet.name}:column-dimensions`);
        }

        const rows = rowIndexes(beforeSheet);
        for (const row of rowIndexes(afterSheet)) {
            rows.add(row);
        }
        for (const row of rows) {
            const beforeRow = beforeSheet.rows?.[row];
            const afterRow = afterSheet.rows?.[row];
            const beforeHeight = isRowData(beforeRow)
                ? beforeRow.height
                : undefined;
            const afterHeight = isRowData(afterRow)
                ? afterRow.height
                : undefined;
            if (!sameValue(beforeHeight, afterHeight)) {
                unsupportedChanges.add(`${afterSheet.name}:row-dimensions`);
            }
        }

        const positions = cellPositions(beforeSheet);
        for (const [key, position] of cellPositions(afterSheet)) {
            positions.set(key, position);
        }

        for (const { row, column } of positions.values()) {
            const beforeCell = getCell(beforeSheet, row, column);
            const afterCell = getCell(afterSheet, row, column);
            if (
                !sameValue(beforeCell?.merge, afterCell?.merge) ||
                !sameValue(beforeCell?.editable, afterCell?.editable)
            ) {
                unsupportedChanges.add(`${afterSheet.name}:cell-structure`);
            }
            const operation: NativeExcelCellEdit = {
                sheetName: afterSheet.name,
                row: row + 1,
                column: column + 1,
            };

            if ((beforeCell?.text ?? '') !== (afterCell?.text ?? '')) {
                operation.value = toNativeValue(afterCell?.text);
            }
            const style = buildStylePatch(
                getStyle(beforeSheet, beforeCell),
                getStyle(afterSheet, afterCell)
            );
            if (style) {
                operation.style = style;
            }
            if (operation.value || operation.style) {
                operations.push(operation);
            }
        }
    }
    return {
        operations,
        unsupportedChanges: [...unsupportedChanges].sort(),
    };
}

export function buildNativeCellEditOperations(
    beforeSheets: SheetData[],
    afterSheets: SheetData[]
): NativeExcelCellEdit[] {
    return buildNativeExcelEditPlan(beforeSheets, afterSheets).operations;
}
