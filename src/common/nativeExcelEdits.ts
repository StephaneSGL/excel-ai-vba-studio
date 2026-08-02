import type {
    SheetChartData,
    SheetTableData,
} from './excelWorkbookObjects';

export interface NativeExcelCellValue {
    kind: 'blank' | 'formula' | 'number' | 'text';
    value?: string | number;
}

export interface NativeExcelStylePatch {
    align?: 'left' | 'center' | 'right' | null;
    valign?: 'top' | 'middle' | 'bottom' | null;
    bgcolor?: string | null;
    color?: string | null;
    format?: string | null;
    textwrap?: boolean | null;
    underline?: boolean | null;
    strike?: boolean | null;
    font?: {
        name?: string | null;
        size?: number | null;
        bold?: boolean | null;
        italic?: boolean | null;
    };
    border?: {
        top?: string[] | null;
        right?: string[] | null;
        bottom?: string[] | null;
        left?: string[] | null;
    };
}

export interface NativeExcelCellEdit {
    kind?: 'cell';
    sheetName: string;
    row: number;
    column: number;
    value?: NativeExcelCellValue;
    style?: NativeExcelStylePatch;
}

export interface NativeExcelColumnWidthEdit {
    kind: 'columnWidth';
    sheetName: string;
    column: number;
    widthPx: number;
}

export interface NativeExcelRowHeightEdit {
    kind: 'rowHeight';
    sheetName: string;
    row: number;
    heightPx: number;
}

export type NativeExcelConditionalFormattingRule =
    | {
        type: 'cellIs';
        operator: 'greaterThan' | 'lessThan' | 'equal';
        operand: string | number;
        fillColor: string;
        fontColor: string;
        bold: boolean;
    }
    | {
        type: 'containsText';
        text: string;
        fillColor: string;
        fontColor: string;
        bold: boolean;
    }
    | {
        type: 'colorScale';
        colors: [string, string, string];
    }
    | {
        type: 'dataBar';
        color: string;
    }
    | {
        type: 'iconSet';
        iconSet: '3TrafficLights1';
        thresholds: [33, 67];
    };

export interface NativeExcelAddConditionalFormattingEdit {
    kind: 'addConditionalFormatting';
    sheetName: string;
    rangeRef: string;
    rule: NativeExcelConditionalFormattingRule;
}

export interface NativeExcelClearConditionalFormattingEdit {
    kind: 'clearConditionalFormatting';
    sheetName: string;
}

export interface NativeExcelCreateTableEdit {
    kind: 'createTable';
    sheetName: string;
    table: SheetTableData;
}

export interface NativeExcelUpdateTableEdit {
    kind: 'updateTable';
    sheetName: string;
    /** Current ListObject name. The complete target state may rename it. */
    name: string;
    table: SheetTableData;
}

export interface NativeExcelDeleteTableEdit {
    kind: 'deleteTable';
    sheetName: string;
    name: string;
}

export interface NativeExcelCreateChartEdit {
    kind: 'createChart';
    sheetName: string;
    chart: SheetChartData;
}

export interface NativeExcelUpdateChartEdit {
    kind: 'updateChart';
    sheetName: string;
    /** Current ChartObject name. The complete target state may rename it. */
    name: string;
    chart: SheetChartData;
    /** Keep the native ChartObject geometry when the UI model did not change it. */
    preserveAnchor?: boolean;
    /** Keep native series, including unsupported trendlines/error bars/point formats. */
    preserveSeries?: boolean;
	/** Apply an explicit ChartStyle without rebuilding series; style-controlled formatting may change. */
	allowSeriesFormattingChange?: boolean;
}

export interface NativeExcelDeleteChartEdit {
    kind: 'deleteChart';
    sheetName: string;
    name: string;
}

export type NativeExcelEditOperation =
    | NativeExcelCellEdit
    | NativeExcelColumnWidthEdit
    | NativeExcelRowHeightEdit
    | NativeExcelAddConditionalFormattingEdit
    | NativeExcelClearConditionalFormattingEdit
    | NativeExcelCreateTableEdit
    | NativeExcelUpdateTableEdit
    | NativeExcelDeleteTableEdit
    | NativeExcelCreateChartEdit
    | NativeExcelUpdateChartEdit
    | NativeExcelDeleteChartEdit;

export interface NativeExcelEditPayload {
    version: 2;
    transactionId: string;
    expectedWorkbookSha256: string;
    operations: NativeExcelEditOperation[];
}

export interface NativeExcelEditResult {
    backupPath: string;
    workbookSha256: string;
}
