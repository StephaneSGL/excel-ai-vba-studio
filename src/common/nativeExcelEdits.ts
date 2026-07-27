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
    sheetName: string;
    row: number;
    column: number;
    value?: NativeExcelCellValue;
    style?: NativeExcelStylePatch;
}

export interface NativeExcelEditPayload {
    version: 2;
    transactionId: string;
    expectedWorkbookSha256: string;
    operations: NativeExcelCellEdit[];
}

export interface NativeExcelEditResult {
    backupPath: string;
    workbookSha256: string;
}
