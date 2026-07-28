import type {
    NativeExcelCellEdit,
    NativeExcelCellValue,
    NativeExcelConditionalFormattingRule,
    NativeExcelEditOperation,
    NativeExcelStylePatch,
} from '@/common/nativeExcelEdits';
import type {
    CellData,
    CellStyle,
    RowData,
    SheetConditionalFormatting,
    SheetConditionalFormattingRule,
    SheetData,
} from './x-spreadsheet/index';

type CellPosition = { row: number; column: number };

const MAX_CONDITIONAL_FORMATTING_ADDS_PER_SHEET = 64;

export interface NativeExcelEditPlan {
    operations: NativeExcelEditOperation[];
    unsupportedChanges: string[];
}

interface NativeEditSheetAdapter {
    loadData(sheets: SheetData[]): unknown;
    getData(): SheetData[];
}

function cloneSheets(sheets: SheetData[]): SheetData[] {
    return JSON.parse(JSON.stringify(sheets)) as SheetData[];
}

export function initializeNativeEditSheets(
    spreadsheet: NativeEditSheetAdapter,
    sourceSheets: SheetData[],
    restoredSheets: SheetData[] | null = null
): SheetData[] {
    spreadsheet.loadData(sourceSheets);
    const baseline = cloneSheets(spreadsheet.getData());
    if (restoredSheets) {
        spreadsheet.loadData(restoredSheets);
    }
    return baseline;
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
        conditionalFormattings: _conditionalFormattings,
        ...features
    } = sheet;
    return features;
}

function isRecord(value: unknown): value is Record<string, unknown> {
    return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function hasOnlyKeys(
    value: Record<string, unknown>,
    allowedKeys: readonly string[]
): boolean {
    const allowed = new Set(allowedKeys);
    return Object.keys(value).every(key => allowed.has(key));
}

function columnIndexes(sheet: SheetData | undefined): Set<number> {
    const indexes = new Set<number>();
    for (const key of Object.keys(sheet?.cols ?? {})) {
        const index = Number(key);
        if (key !== 'len' && Number.isInteger(index) && index >= 0) {
            indexes.add(index);
        }
    }
    return indexes;
}

function getColumnWidth(
    sheet: SheetData | undefined,
    column: number
): number | undefined {
    const dimension = sheet?.cols?.[column];
    return (
        dimension &&
        typeof dimension === 'object' &&
        hasOnlyKeys(dimension as unknown as Record<string, unknown>, ['width'])
    )
        ? dimension.width
        : undefined;
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

const SIMPLE_A1_RANGE =
    /^[A-Z]{1,3}[1-9][0-9]{0,6}(?::[A-Z]{1,3}[1-9][0-9]{0,6})?$/;

function exactArgb(value: unknown, expected: string): boolean {
    return (
        isRecord(value) &&
        hasOnlyKeys(value, ['argb']) &&
        value.argb === expected
    );
}

function hasGeneratedHighlightStyle(
    rule: Record<string, unknown>
): boolean {
    if (!isRecord(rule.style) || !hasOnlyKeys(rule.style, ['fill', 'font'])) {
        return false;
    }
    const fill = rule.style.fill;
    const font = rule.style.font;
    if (
        !isRecord(fill) ||
        !hasOnlyKeys(fill, ['type', 'pattern', 'fgColor']) ||
        fill.type !== 'pattern' ||
        fill.pattern !== 'solid' ||
        !exactArgb(fill.fgColor, 'FFFFC7CE') ||
        !isRecord(font) ||
        !hasOnlyKeys(font, ['color']) ||
        !exactArgb(font.color, 'FF9C0006')
    ) {
        return false;
    }
    const displayStyle = rule.displayStyle;
    return (
        isRecord(displayStyle) &&
        hasOnlyKeys(displayStyle, ['bgcolor', 'color', 'font']) &&
        displayStyle.bgcolor === '#ffc7ce' &&
        displayStyle.color === '#9c0006' &&
        isRecord(displayStyle.font) &&
        hasOnlyKeys(displayStyle.font, ['bold']) &&
        displayStyle.font.bold === true
    );
}

function normalizeConditionalRule(
    rule: SheetConditionalFormattingRule
): NativeExcelConditionalFormattingRule | undefined {
    const value = rule as Record<string, unknown>;
    if (value.type === 'cellIs') {
        if (
            !hasOnlyKeys(value, [
                'type',
                'operator',
                'formulae',
                'style',
                'displayStyle',
                'priority',
            ]) ||
            !['greaterThan', 'lessThan', 'equal'].includes(
                String(value.operator)
            ) ||
            !Array.isArray(value.formulae) ||
            value.formulae.length !== 1 ||
            !hasGeneratedHighlightStyle(value)
        ) {
            return undefined;
        }
        const operand = value.formulae[0];
        if (
            !(
                (typeof operand === 'number' && Number.isFinite(operand)) ||
                (typeof operand === 'string' &&
                    operand.length > 0 &&
                    operand.length <= 255 &&
                    !operand.includes('\0') &&
                    !operand.startsWith('='))
            )
        ) {
            return undefined;
        }
        return {
            type: 'cellIs',
            operator: value.operator as 'greaterThan' | 'lessThan' | 'equal',
            operand,
            fillColor: '#ffc7ce',
            fontColor: '#9c0006',
            bold: true,
        };
    }

    if (value.type === 'containsText') {
        if (
            !hasOnlyKeys(value, [
                'type',
                'operator',
                'text',
                'formulae',
                'style',
                'displayStyle',
                'priority',
            ]) ||
            value.operator !== 'containsText' ||
            typeof value.text !== 'string' ||
            value.text.length === 0 ||
            value.text.length > 255 ||
            value.text.includes('\0') ||
            !Array.isArray(value.formulae) ||
            value.formulae.length !== 1 ||
            value.formulae[0] !== value.text ||
            !hasGeneratedHighlightStyle(value)
        ) {
            return undefined;
        }
        return {
            type: 'containsText',
            text: value.text,
            fillColor: '#ffc7ce',
            fontColor: '#9c0006',
            bold: true,
        };
    }

    if (value.type === 'colorScale') {
        if (
            !hasOnlyKeys(value, ['type', 'cfvo', 'color', 'priority']) ||
            !sameValue(value.cfvo, [
                { type: 'min' },
                { type: 'percentile', value: 50 },
                { type: 'max' },
            ]) ||
            !sameValue(value.color, [
                { argb: 'FFF8696B' },
                { argb: 'FFFFEB84' },
                { argb: 'FF63BE7B' },
            ])
        ) {
            return undefined;
        }
        return {
            type: 'colorScale',
            colors: ['#f8696b', '#ffeb84', '#63be7b'],
        };
    }

    if (value.type === 'dataBar') {
        if (
            !hasOnlyKeys(value, ['type', 'cfvo', 'color', 'priority']) ||
            !sameValue(value.cfvo, [{ type: 'min' }, { type: 'max' }]) ||
            !exactArgb(value.color, 'FF5B9BD5')
        ) {
            return undefined;
        }
        return { type: 'dataBar', color: '#5b9bd5' };
    }

    if (value.type === 'iconSet') {
        if (
            !hasOnlyKeys(value, [
                'type',
                'iconSet',
                'cfvo',
                'priority',
            ]) ||
            value.iconSet !== '3TrafficLights1' ||
            !sameValue(value.cfvo, [
                { type: 'min' },
                { type: 'percent', value: 33 },
                { type: 'percent', value: 67 },
            ])
        ) {
            return undefined;
        }
        return {
            type: 'iconSet',
            iconSet: '3TrafficLights1',
            thresholds: [33, 67],
        };
    }
    return undefined;
}

function buildConditionalFormattingOperations(
    sheetName: string,
    beforeDefinitions: SheetConditionalFormatting[] | undefined,
    afterDefinitions: SheetConditionalFormatting[] | undefined
): NativeExcelEditOperation[] | undefined {
    const before = beforeDefinitions ?? [];
    const after = afterDefinitions ?? [];
    if (sameValue(before, after)) {
        return [];
    }
    if (before.length > 0 && after.length === 0) {
        return [{ kind: 'clearConditionalFormatting', sheetName }];
    }
    if (
        after.length <= before.length ||
        after.length - before.length >
            MAX_CONDITIONAL_FORMATTING_ADDS_PER_SHEET ||
        !sameValue(before, after.slice(0, before.length))
    ) {
        return undefined;
    }

    const operations: NativeExcelEditOperation[] = [];
    for (const definition of after.slice(before.length)) {
        const rangeRef = definition.ref?.toUpperCase();
        if (
            !rangeRef ||
            !SIMPLE_A1_RANGE.test(rangeRef) ||
            definition.rules?.length !== 1
        ) {
            return undefined;
        }
        const rule = normalizeConditionalRule(definition.rules[0]);
        if (!rule) {
            return undefined;
        }
        operations.push({
            kind: 'addConditionalFormatting',
            sheetName,
            rangeRef,
            rule,
        });
    }
    return operations;
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
    const operations: NativeExcelEditOperation[] = [];

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
        const columns = columnIndexes(beforeSheet);
        for (const column of columnIndexes(afterSheet)) {
            columns.add(column);
        }
        for (const column of columns) {
            const beforeDimension = beforeSheet.cols?.[column];
            const afterDimension = afterSheet.cols?.[column];
            if (sameValue(beforeDimension, afterDimension)) {
                continue;
            }
            const widthPx = getColumnWidth(afterSheet, column);
            if (
                typeof widthPx === 'number' &&
                Number.isFinite(widthPx) &&
                widthPx > 5 &&
                widthPx <= 1_790
            ) {
                operations.push({
                    kind: 'columnWidth',
                    sheetName: afterSheet.name,
                    column: column + 1,
                    widthPx,
                });
            } else {
                unsupportedChanges.add(
                    `${afterSheet.name}:column-dimensions`
                );
            }
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
                if (
                    typeof afterHeight === 'number' &&
                    Number.isFinite(afterHeight) &&
                    afterHeight > 0 &&
                    afterHeight <= 546
                ) {
                    operations.push({
                        kind: 'rowHeight',
                        sheetName: afterSheet.name,
                        row: row + 1,
                        heightPx: afterHeight,
                    });
                } else {
                    unsupportedChanges.add(
                        `${afterSheet.name}:row-dimensions`
                    );
                }
            }
        }

        const conditionalOperations = buildConditionalFormattingOperations(
            afterSheet.name,
            beforeSheet.conditionalFormattings,
            afterSheet.conditionalFormattings
        );
        if (conditionalOperations) {
            operations.push(...conditionalOperations);
        } else {
            unsupportedChanges.add(
                `${afterSheet.name}:conditional-formatting`
            );
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
    return buildNativeExcelEditPlan(beforeSheets, afterSheets)
        .operations
        .filter((operation): operation is NativeExcelCellEdit =>
            operation.kind === undefined || operation.kind === 'cell'
        );
}
