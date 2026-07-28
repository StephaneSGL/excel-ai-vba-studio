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

export type NativeExcelEditOperation =
    | NativeExcelCellEdit
    | NativeExcelColumnWidthEdit
    | NativeExcelRowHeightEdit
    | NativeExcelAddConditionalFormattingEdit
    | NativeExcelClearConditionalFormattingEdit;

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
