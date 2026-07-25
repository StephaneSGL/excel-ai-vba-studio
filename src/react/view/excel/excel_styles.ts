import type * as ExcelJS from '@cweijan/exceljs';
import type { CellStyle } from './x-spreadsheet/index';

type ExcelColorValue = Partial<ExcelJS.Color> & {
    indexed?: number;
    tint?: number;
    auto?: boolean;
};

export type ExcelColorResolver = (color?: ExcelColorValue) => string | undefined;

// ECMA-376 indexed colour palette used by legacy and compatibility styles.
// Indexes 64 and 65 are automatic foreground/background colours.
const INDEXED_COLORS = [
    '000000', 'ffffff', 'ff0000', '00ff00', '0000ff', 'ffff00', 'ff00ff', '00ffff',
    '000000', 'ffffff', 'ff0000', '00ff00', '0000ff', 'ffff00', 'ff00ff', '00ffff',
    '800000', '008000', '000080', '808000', '800080', '008080', 'c0c0c0', '808080',
    '9999ff', '993366', 'ffffcc', 'ccffff', '660066', 'ff8080', '0066cc', 'ccccff',
    '000080', 'ff00ff', 'ffff00', '00ffff', '800080', '800000', '008080', '0000ff',
    '00ccff', 'ccffff', 'ccffcc', 'ffff99', '99ccff', 'ff99cc', 'cc99ff', 'ffcc99',
    '3366ff', '33cccc', '99cc00', 'ffcc00', 'ff9900', 'ff6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
] as const;

const THEME_COLOR_KEYS = [
    'dk1', 'lt1', 'dk2', 'lt2',
    'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
    'hlink', 'folHlink',
] as const;

function normalizeRgb(value?: string): string | undefined {
    const normalized = value?.replace(/^#/, '').trim();
    if (!normalized) return undefined;
    if (/^[0-9a-f]{8}$/i.test(normalized)) return normalized.slice(2).toLowerCase();
    if (/^[0-9a-f]{6}$/i.test(normalized)) return normalized.toLowerCase();
    return undefined;
}

function rgbToHsl(rgb: string): [number, number, number] {
    const values = [0, 2, 4].map(index => parseInt(rgb.slice(index, index + 2), 16) / 255);
    const max = Math.max(...values);
    const min = Math.min(...values);
    const lightness = (max + min) / 2;
    if (max === min) return [0, 0, lightness];
    const delta = max - min;
    const saturation = lightness > 0.5
        ? delta / (2 - max - min)
        : delta / (max + min);
    let hue = 0;
    if (max === values[0]) hue = (values[1] - values[2]) / delta + (values[1] < values[2] ? 6 : 0);
    else if (max === values[1]) hue = (values[2] - values[0]) / delta + 2;
    else hue = (values[0] - values[1]) / delta + 4;
    return [hue / 6, saturation, lightness];
}

function hslToRgb(hue: number, saturation: number, lightness: number): string {
    const hueToRgb = (p: number, q: number, value: number) => {
        let t = value;
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1 / 6) return p + (q - p) * 6 * t;
        if (t < 1 / 2) return q;
        if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
        return p;
    };
    const values = saturation === 0
        ? [lightness, lightness, lightness]
        : (() => {
            const q = lightness < 0.5
                ? lightness * (1 + saturation)
                : lightness + saturation - lightness * saturation;
            const p = 2 * lightness - q;
            return [
                hueToRgb(p, q, hue + 1 / 3),
                hueToRgb(p, q, hue),
                hueToRgb(p, q, hue - 1 / 3),
            ];
        })();
    return values
        .map(value => Math.round(Math.max(0, Math.min(1, value)) * 255).toString(16).padStart(2, '0'))
        .join('');
}

function applyTint(rgb: string, tint?: number): string {
    if (tint == null || tint === 0) return rgb;
    const [hue, saturation, lightness] = rgbToHsl(rgb);
    const nextLightness = tint < 0
        ? lightness * (1 + tint)
        : lightness * (1 - tint) + tint;
    return hslToRgb(hue, saturation, Math.max(0, Math.min(1, nextLightness)));
}

function workbookThemeColors(workbook: ExcelJS.Workbook): string[] {
    const themes = (workbook as unknown as {
        themes?: Record<string, string> | string[];
    }).themes;
    const themeXml = Array.isArray(themes)
        ? themes.find(value => typeof value === 'string')
        : themes?.theme1 ?? Object.values(themes ?? {}).find(value => typeof value === 'string');
    if (!themeXml) return [];
    return THEME_COLOR_KEYS.map(key => {
        const block = themeXml.match(new RegExp(
            `<(?:[\\w-]+:)?${key}\\b[^>]*>([\\s\\S]*?)<\\/(?:[\\w-]+:)?${key}>`,
            'i',
        ))?.[1];
        if (!block) return '';
        const srgb = block.match(/<(?:[\w-]+:)?srgbClr\b[^>]*\bval="([^"]+)"/i)?.[1];
        const system = block.match(/<(?:[\w-]+:)?sysClr\b[^>]*\blastClr="([^"]+)"/i)?.[1]
            ?? block.match(/<(?:[\w-]+:)?sysClr\b[^>]*\bval="([^"]+)"/i)?.[1];
        return normalizeRgb(srgb ?? system) ?? '';
    });
}

export function createExcelColorResolver(workbook: ExcelJS.Workbook): ExcelColorResolver {
    const themeColors = workbookThemeColors(workbook);
    return color => colorToHex(color, themeColors);
}

const BORDER_FROM_EXCEL: Record<string, string> = {
    thin: 'thin',
    hair: 'dotted',
    dotted: 'dotted',
    medium: 'medium',
    mediumDashed: 'dashed',
    mediumDashDot: 'dashed',
    mediumDashDotDot: 'dashed',
    dashed: 'dashed',
    dashDot: 'dashed',
    dashDotDot: 'dashed',
    slantDashDot: 'dashed',
    thick: 'thick',
    double: 'double',
};

const BORDER_TO_EXCEL: Record<string, ExcelJS.BorderStyle> = {
    thin: 'thin',
    medium: 'medium',
    thick: 'thick',
    dashed: 'dashed',
    dotted: 'dotted',
    double: 'double',
};

const VALIGN_FROM_EXCEL: Record<string, CellStyle['valign']> = {
    top: 'top',
    middle: 'middle',
    bottom: 'bottom',
};

const VALIGN_TO_EXCEL: Record<string, ExcelJS.Alignment['vertical']> = {
    top: 'top',
    middle: 'middle',
    bottom: 'bottom',
};

const FORMAT_TO_NUMFMT: Record<string, string> = {
    normal: 'General',
    text: '@',
    number: '#,##0.00',
    number_plain: '0.00',
    percent: '0.00%',
    rmb: '¥#,##0.00',
    usd: '$#,##0.00',
    eur: '€#,##0.00',
    date: 'yyyy/m/d',
    time: 'h:mm:ss',
    datetime: 'yyyy/m/d h:mm',
    duration: '[h]:mm:ss',
};

const NUMFMT_PATTERNS: { pattern: RegExp; format: string }[] = [
    { pattern: /^general$/i, format: 'normal' },
    { pattern: /^@$/, format: 'text' },
    { pattern: /%/, format: 'percent' },
    { pattern: /[¥￥]/, format: 'rmb' },
    { pattern: /\$/, format: 'usd' },
    { pattern: /€/, format: 'eur' },
    { pattern: /\[h\]:mm/i, format: 'duration' },
    { pattern: /yyyy.*h:mm|m\/d.*h:mm/i, format: 'datetime' },
    { pattern: /h:mm|hh:mm/i, format: 'time' },
    { pattern: /yyyy|m\/d|d\/m|dd\/mm/i, format: 'date' },
    { pattern: /,/, format: 'number' },
    { pattern: /#|0\.0/, format: 'number_plain' },
];

function numFmtToSpreadsheetFormat(numFmt?: string): string | undefined {
    if (!numFmt || numFmt === 'General') return undefined;
    for (let i = 0; i < NUMFMT_PATTERNS.length; i += 1) {
        if (NUMFMT_PATTERNS[i].pattern.test(numFmt)) {
            return NUMFMT_PATTERNS[i].format;
        }
    }
    return undefined;
}

function spreadsheetFormatToNumFmt(format?: string): string | undefined {
    if (!format || format === 'normal') return undefined;
    return FORMAT_TO_NUMFMT[format];
}

export function colorToHex(
    color?: ExcelColorValue,
    themeColors: string[] = [],
): string | undefined {
    if (!color) return undefined;
    const direct = normalizeRgb(color.argb);
    if (direct) return `#${applyTint(direct, color.tint)}`;
    if (color.theme != null) {
        const themed = themeColors[color.theme];
        if (themed) return `#${applyTint(themed, color.tint)}`;
    }
    if (color.indexed != null && color.indexed >= 0 && color.indexed < INDEXED_COLORS.length) {
        return `#${applyTint(INDEXED_COLORS[color.indexed], color.tint)}`;
    }
    return undefined;
}

export function hexToArgb(hex?: string): string | undefined {
    if (!hex) return undefined;
    const normalized = hex.replace(/^#/, '');
    if (normalized.length === 6) return `FF${normalized.toUpperCase()}`;
    if (normalized.length === 8) return normalized.toUpperCase();
    return undefined;
}

function borderSideToSpreadsheet(
    side?: Partial<ExcelJS.Border>,
    resolveColor: ExcelColorResolver = colorToHex,
): string[] | undefined {
    if (!side?.style) return undefined;
    const style = BORDER_FROM_EXCEL[side.style] ?? 'thin';
    const color = resolveColor(side.color) ?? '#000000';
    return [style, color];
}

function bordersToSpreadsheet(
    borders?: Partial<ExcelJS.Borders>,
    resolveColor: ExcelColorResolver = colorToHex,
): CellStyle['border'] | undefined {
    if (!borders) return undefined;
    const border: CellStyle['border'] = {};
    const top = borderSideToSpreadsheet(borders.top, resolveColor);
    const right = borderSideToSpreadsheet(borders.right, resolveColor);
    const bottom = borderSideToSpreadsheet(borders.bottom, resolveColor);
    const left = borderSideToSpreadsheet(borders.left, resolveColor);
    if (top) border.top = top;
    if (right) border.right = right;
    if (bottom) border.bottom = bottom;
    if (left) border.left = left;
    return Object.keys(border).length > 0 ? border : undefined;
}

function borderSideToExcelJs(side?: string[]): Partial<ExcelJS.Border> | undefined {
    if (!side?.[0]) return undefined;
    const style = BORDER_TO_EXCEL[side[0]] ?? 'thin';
    const argb = hexToArgb(side[1]);
    return {
        style,
        color: argb ? { argb } : undefined,
    };
}

function bordersToExcelJs(border?: CellStyle['border']): Partial<ExcelJS.Borders> | undefined {
    if (!border) return undefined;
    const borders: Partial<ExcelJS.Borders> = {};
    const top = borderSideToExcelJs(border.top);
    const right = borderSideToExcelJs(border.right);
    const bottom = borderSideToExcelJs(border.bottom);
    const left = borderSideToExcelJs(border.left);
    if (top) borders.top = top;
    if (right) borders.right = right;
    if (bottom) borders.bottom = bottom;
    if (left) borders.left = left;
    return Object.keys(borders).length > 0 ? borders : undefined;
}

export function excelJsCellToStyle(
    cell: ExcelJS.Cell,
    resolveColor: ExcelColorResolver = colorToHex,
): CellStyle | null {
    const style: CellStyle = {};
    let hasStyle = false;

    const font = cell.font;
    if (font) {
        const fontStyle: NonNullable<CellStyle['font']> = {};
        let hasFont = false;
        if (font.name) {
            fontStyle.name = font.name;
            hasFont = true;
        }
        if (font.size) {
            fontStyle.size = font.size;
            hasFont = true;
        }
        if (font.bold) {
            fontStyle.bold = true;
            hasFont = true;
        }
        if (font.italic) {
            fontStyle.italic = true;
            hasFont = true;
        }
        if (hasFont) {
            style.font = fontStyle;
            hasStyle = true;
        }
        const color = resolveColor(font.color);
        if (color) {
            style.color = color;
            hasStyle = true;
        }
        if (font.strike) {
            style.strike = true;
            hasStyle = true;
        }
        if (font.underline && font.underline !== 'none' && font.underline !== false) {
            style.underline = true;
            hasStyle = true;
        }
    }

    const alignment = cell.alignment;
    if (alignment) {
        if (alignment.horizontal === 'left' || alignment.horizontal === 'center' || alignment.horizontal === 'right') {
            style.align = alignment.horizontal;
            hasStyle = true;
        }
        const valign = alignment.vertical ? VALIGN_FROM_EXCEL[alignment.vertical] : undefined;
        if (valign) {
            style.valign = valign;
            hasStyle = true;
        }
        if (alignment.wrapText) {
            style.textwrap = true;
            hasStyle = true;
        }
    }

    const fill = cell.fill;
    if (fill && fill.type === 'pattern') {
        const bgcolor = resolveColor(fill.fgColor) ?? resolveColor(fill.bgColor);
        if (bgcolor && bgcolor !== '#ffffff') {
            style.bgcolor = bgcolor;
            hasStyle = true;
        }
    }

    if (cell.border) {
        const border = bordersToSpreadsheet(cell.border, resolveColor);
        if (border) {
            style.border = border;
            hasStyle = true;
        }
    }

    const numFmt = cell.numFmt;
    const format = numFmtToSpreadsheetFormat(numFmt);
    if (format) {
        style.format = format;
        hasStyle = true;
    }

    return hasStyle ? style : null;
}

export function excelJsStyleToCellStyle(
    style?: Partial<ExcelJS.Style>,
    resolveColor: ExcelColorResolver = colorToHex,
): CellStyle | null {
    if (!style) return null;
    return excelJsCellToStyle(style as unknown as ExcelJS.Cell, resolveColor);
}

export function applySpreadsheetStyle(cell: ExcelJS.Cell, style: CellStyle) {
    const font: Partial<ExcelJS.Font> = { ...(cell.font ?? {}) };

    if (style.font) {
        if (style.font.name) font.name = style.font.name;
        if (style.font.size) font.size = style.font.size;
        if (style.font.bold != null) font.bold = style.font.bold;
        if (style.font.italic != null) font.italic = style.font.italic;
    }
    if (style.color) {
        const argb = hexToArgb(style.color);
        if (argb) font.color = { argb };
    }
    if (style.strike != null) font.strike = style.strike;
    if (style.underline != null) font.underline = style.underline ? 'single' : false;
    if (Object.keys(font).length > 0) cell.font = font;

    if (style.align || style.valign || style.textwrap != null) {
        const alignment: Partial<ExcelJS.Alignment> = { ...(cell.alignment ?? {}) };
        if (style.align) alignment.horizontal = style.align;
        if (style.valign) alignment.vertical = VALIGN_TO_EXCEL[style.valign];
        if (style.textwrap != null) alignment.wrapText = style.textwrap;
        cell.alignment = alignment;
    }

    if (style.bgcolor) {
        const argb = hexToArgb(style.bgcolor);
        if (argb) {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb },
            };
        }
    }

    const border = bordersToExcelJs(style.border);
    if (border) cell.border = border;

    const numFmt = spreadsheetFormatToNumFmt(style.format);
    if (numFmt) cell.numFmt = numFmt;
}

export class StyleRegistry {
    private styles: CellStyle[] = [];

    add(style: CellStyle | null): number | undefined {
        if (!style) return undefined;
        for (let i = 0; i < this.styles.length; i += 1) {
            if (JSON.stringify(this.styles[i]) === JSON.stringify(style)) {
                return i;
            }
        }
        this.styles.push(style);
        return this.styles.length - 1;
    }

    getStyles(): CellStyle[] {
        return this.styles;
    }
}
