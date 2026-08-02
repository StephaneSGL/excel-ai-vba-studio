import {
    chartDataLabelsHaveEnabledShowOption,
    chartDataLabelsHaveExplicitShowOption,
    chartSeriesSupportsBubbleSizes,
    chartSeriesSupportsDataLabelPosition,
    chartSeriesSupportsPercentageDataLabels,
    chartSeriesSupportsSmooth,
    normalizeA1Range,
    SIMPLE_A1_RANGE,
    type ChartLegendPosition,
    type ChartMarkerStyle,
    type SheetChartAxisData,
    type SheetChartData,
    type SheetChartSeriesData,
} from '../../../common/excelWorkbookObjects';

export interface OoxmlZipEntryLike {
    async(type: 'string'): Promise<string>;
    dir?: boolean;
    _data?: { uncompressedSize?: number };
}

export interface OoxmlZipLike {
    file(path: string): OoxmlZipEntryLike | null;
}

export interface OoxmlWorksheetPart {
    sheetName: string;
    worksheetPart: string;
}

export interface OoxmlSheetChartInventory extends OoxmlWorksheetPart {
    charts: SheetChartData[];
    /** True when the worksheet drawing references any classic chart or chartEx part. */
    hasChartParts: boolean;
    /** Referenced chart parts that were deliberately not hydrated. */
    unsupportedChartCount: number;
}

export interface OoxmlChartInventoryResult {
    sheets: OoxmlSheetChartInventory[];
    warnings: string[];
}

export interface OoxmlChartReaderLimits {
    maxWorksheets: number;
    maxDrawingsPerWorksheet: number;
    maxAnchorsPerDrawing: number;
    maxCharts: number;
    maxSeriesPerChart: number;
    maxXmlCharacters: number;
    maxChartXmlCharacters: number;
    maxTotalXmlCharacters: number;
    maxXmlNodes: number;
    maxXmlDepth: number;
    maxFormulaCharacters: number;
    maxWarnings: number;
}

export const DEFAULT_OOXML_CHART_READER_LIMITS: Readonly<OoxmlChartReaderLimits> = Object.freeze({
    maxWorksheets: 256,
    maxDrawingsPerWorksheet: 16,
    maxAnchorsPerDrawing: 128,
    maxCharts: 512,
    maxSeriesPerChart: 255,
    maxXmlCharacters: 2 * 1024 * 1024,
    maxChartXmlCharacters: 4 * 1024 * 1024,
    maxTotalXmlCharacters: 64 * 1024 * 1024,
    maxXmlNodes: 50_000,
    maxXmlDepth: 128,
    maxFormulaCharacters: 1024,
    maxWarnings: 128,
});

interface XmlNode {
    name: string;
    localName: string;
    attributes: Record<string, string>;
    children: XmlNode[];
    text: string;
}

interface Relationship {
    id: string;
    type: string;
    target: string;
}

interface ReaderContext {
    zip: OoxmlZipLike;
    limits: OoxmlChartReaderLimits;
    warnings: string[];
    totalXmlCharacters: number;
    reservedXmlCharacters: number;
    chartCount: number;
    chartAttempts: number;
}

interface ParsedSeries extends SheetChartSeriesData {
    _valueRange?: ParsedRange;
    _categoryRange?: ParsedRange;
    _nameRange?: ParsedRange;
}

interface ParsedRange {
    ref: string;
    sri: number;
    sci: number;
    eri: number;
    eci: number;
}

interface ChartGroup {
    node: XmlNode;
    chartType: number;
    axisIds: string[];
    secondary: boolean;
}

const EMUS_PER_POINT = 12_700;
const DEFAULT_COLUMN_POINTS = 48;
const DEFAULT_ROW_POINTS = 15;
const MAX_COORDINATE_POINTS = 1_000_000;
const MAX_TEXT_CHARACTERS = 8_192;
const MAX_TRAVERSAL_NODES = 50_000;
const XML_NAME = /^[A-Za-z_][A-Za-z0-9_.:-]*$/;

function warn(context: ReaderContext, message: string): void {
    if (context.warnings.length < context.limits.maxWarnings) context.warnings.push(message);
}

function localName(name: string): string {
    const separator = name.lastIndexOf(':');
    return separator >= 0 ? name.slice(separator + 1) : name;
}

function decodeXml(value: string): string | null {
    let result = '';
    let cursor = 0;
    const expression = /&(#x[0-9a-f]+|#[0-9]+|amp|lt|gt|quot|apos);/gi;
    for (let match = expression.exec(value); match; match = expression.exec(value)) {
        const prefix = value.slice(cursor, match.index);
        if (prefix.includes('&')) return null;
        result += prefix;
        const entity = match[1].toLowerCase();
        if (entity === 'amp') result += '&';
        else if (entity === 'lt') result += '<';
        else if (entity === 'gt') result += '>';
        else if (entity === 'quot') result += '"';
        else if (entity === 'apos') result += "'";
        else {
            const codePoint = entity.startsWith('#x')
                ? Number.parseInt(entity.slice(2), 16)
                : Number.parseInt(entity.slice(1), 10);
            if (!Number.isInteger(codePoint)
                || codePoint < 0
                || codePoint > 0x10ffff
                || (codePoint >= 0xd800 && codePoint <= 0xdfff)) return null;
            result += String.fromCodePoint(codePoint);
        }
        cursor = match.index + match[0].length;
    }
    const suffix = value.slice(cursor);
    if (suffix.includes('&')) return null;
    return result + suffix;
}

function findTagEnd(xml: string, start: number): number {
    let quote = '';
    for (let index = start; index < xml.length; index += 1) {
        const character = xml[index];
        if (quote) {
            if (character === quote) quote = '';
        } else if (character === '"' || character === "'") {
            quote = character;
        } else if (character === '>') {
            return index;
        }
    }
    return -1;
}

function parseStartTag(source: string): { name: string; attributes: Record<string, string>; selfClosing: boolean } | null {
    let content = source.trim();
    const selfClosing = content.endsWith('/');
    if (selfClosing) content = content.slice(0, -1).trimEnd();
    const nameMatch = /^([^\s/>]+)/.exec(content);
    if (!nameMatch || !XML_NAME.test(nameMatch[1])) return null;
    const name = nameMatch[1];
    const attributes: Record<string, string> = {};
    let cursor = nameMatch[0].length;
    let attributeCount = 0;
    while (cursor < content.length) {
        while (/\s/.test(content[cursor] ?? '')) cursor += 1;
        if (cursor >= content.length) break;
        const attributeMatch = /^[^\s=/>]+/.exec(content.slice(cursor));
        if (!attributeMatch || !XML_NAME.test(attributeMatch[0])) return null;
        const attributeName = attributeMatch[0];
        cursor += attributeName.length;
        while (/\s/.test(content[cursor] ?? '')) cursor += 1;
        if (content[cursor] !== '=') return null;
        cursor += 1;
        while (/\s/.test(content[cursor] ?? '')) cursor += 1;
        const quote = content[cursor];
        if (quote !== '"' && quote !== "'") return null;
        cursor += 1;
        const end = content.indexOf(quote, cursor);
        if (end < 0 || end - cursor > MAX_TEXT_CHARACTERS) return null;
        const decoded = decodeXml(content.slice(cursor, end));
        if (decoded == null || Object.prototype.hasOwnProperty.call(attributes, attributeName)) return null;
        attributes[attributeName] = decoded;
        cursor = end + 1;
        attributeCount += 1;
        if (attributeCount > 64) return null;
    }
    return { name, attributes, selfClosing };
}

function parseXml(xml: string, limits: OoxmlChartReaderLimits): XmlNode | null {
    if (/<!DOCTYPE|<!ENTITY/i.test(xml)) return null;
    let cursor = 0;
    let root: XmlNode | null = null;
    let nodeCount = 0;
    let textCharacters = 0;
    const stack: XmlNode[] = [];
    const appendText = (raw: string, alreadyDecoded = false): boolean => {
        if (!raw) return true;
        const decoded = alreadyDecoded ? raw : decodeXml(raw);
        if (decoded == null) return false;
        textCharacters += decoded.length;
        if (textCharacters > limits.maxXmlCharacters) return false;
        if (stack.length) stack[stack.length - 1].text += decoded;
        else if (decoded.trim()) return false;
        return true;
    };

    while (cursor < xml.length) {
        const opening = xml.indexOf('<', cursor);
        if (opening < 0) return appendText(xml.slice(cursor)) && stack.length === 0 ? root : null;
        if (!appendText(xml.slice(cursor, opening))) return null;
        if (xml.startsWith('<!--', opening)) {
            const end = xml.indexOf('-->', opening + 4);
            if (end < 0) return null;
            cursor = end + 3;
            continue;
        }
        if (xml.startsWith('<![CDATA[', opening)) {
            const end = xml.indexOf(']]>', opening + 9);
            if (end < 0 || !appendText(xml.slice(opening + 9, end), true)) return null;
            cursor = end + 3;
            continue;
        }
        if (xml.startsWith('<?', opening)) {
            const end = xml.indexOf('?>', opening + 2);
            if (end < 0) return null;
            cursor = end + 2;
            continue;
        }
        if (xml.startsWith('<!', opening)) return null;
        const end = findTagEnd(xml, opening + 1);
        if (end < 0) return null;
        const content = xml.slice(opening + 1, end);
        if (content.startsWith('/')) {
            const closingName = content.slice(1).trim();
            if (!XML_NAME.test(closingName) || !stack.length || stack[stack.length - 1].name !== closingName) return null;
            stack.pop();
        } else {
            const parsed = parseStartTag(content);
            if (!parsed) return null;
            const node: XmlNode = {
                name: parsed.name,
                localName: localName(parsed.name),
                attributes: parsed.attributes,
                children: [],
                text: '',
            };
            nodeCount += 1;
            if (nodeCount > limits.maxXmlNodes) return null;
            if (stack.length) stack[stack.length - 1].children.push(node);
            else if (root) return null;
            else root = node;
            if (!parsed.selfClosing) {
                stack.push(node);
                if (stack.length > limits.maxXmlDepth) return null;
            }
        }
        cursor = end + 1;
    }
    return stack.length === 0 ? root : null;
}

function attribute(node: XmlNode | undefined, wanted: string): string | undefined {
    if (!node) return undefined;
    const direct = node.attributes[wanted];
    if (direct != null) return direct;
    const entries = Object.entries(node.attributes).filter(([name]) => localName(name).toLowerCase() === wanted.toLowerCase());
    return entries.length === 1 ? entries[0][1] : undefined;
}

function directChild(node: XmlNode | undefined, wanted: string): XmlNode | undefined {
    return node?.children.find(child => child.localName === wanted);
}

function directChildren(node: XmlNode | undefined, wanted: string): XmlNode[] {
    return node?.children.filter(child => child.localName === wanted) ?? [];
}

function descendants(node: XmlNode | undefined, wanted: string, maximum = 4096): XmlNode[] {
    if (!node) return [];
    const matches: XmlNode[] = [];
    const pending = [...node.children];
    for (let cursor = 0; cursor < pending.length && cursor < MAX_TRAVERSAL_NODES && matches.length < maximum; cursor += 1) {
        const current = pending[cursor];
        if (current.localName === wanted) matches.push(current);
        if (pending.length + current.children.length <= MAX_TRAVERSAL_NODES) pending.push(...current.children);
    }
    return matches;
}

function firstDescendant(node: XmlNode | undefined, wanted: string): XmlNode | undefined {
    return descendants(node, wanted, 1)[0];
}

function textContent(node: XmlNode | undefined, maximum = MAX_TEXT_CHARACTERS): string {
    if (!node) return '';
    let result = node.text;
    const pending = [...node.children];
    for (let cursor = 0; cursor < pending.length && cursor < MAX_TRAVERSAL_NODES && result.length <= maximum; cursor += 1) {
        const current = pending[cursor];
        result += current.text;
        if (pending.length + current.children.length <= MAX_TRAVERSAL_NODES) pending.push(...current.children);
    }
    return result.slice(0, maximum).trim();
}

function readValue(node: XmlNode | undefined, childName: string): string | undefined {
    const child = directChild(node, childName);
    return attribute(child, 'val') ?? (child ? textContent(child) : undefined);
}

function booleanValue(value: string | undefined, fallback = false): boolean {
    if (value == null) return fallback;
    return value === '1' || value.toLowerCase() === 'true';
}

function finiteNumber(value: string | undefined): number | undefined {
    if (value == null || value.trim() === '') return undefined;
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : undefined;
}

function boundedInteger(value: string | undefined, minimum: number, maximum: number): number | undefined {
    const parsed = finiteNumber(value);
    return parsed != null && Number.isInteger(parsed) && parsed >= minimum && parsed <= maximum ? parsed : undefined;
}

function stableHash(value: string): string {
    let hash = 0xcbf29ce484222325n;
    for (let index = 0; index < value.length; index += 1) {
        hash ^= BigInt(value.charCodeAt(index));
        hash = BigInt.asUintN(64, hash * 0x100000001b3n);
    }
    return hash.toString(16).padStart(16, '0');
}

function normalizePartName(value: string): string | null {
    if (!value || value.length > 1024 || value.includes('\\') || value.includes('\0')) return null;
    if (value.includes('?') || value.includes('#') || /^[A-Za-z][A-Za-z0-9+.-]*:/.test(value)) return null;
    let decoded: string;
    try {
        decoded = decodeURIComponent(value);
    } catch {
        return null;
    }
    if (decoded.includes('\\') || decoded.includes('\0') || decoded.includes('?') || decoded.includes('#')) return null;
    const segments: string[] = [];
    for (const segment of decoded.replace(/^\/+/, '').split('/')) {
        if (!segment || segment === '.') continue;
        if (segment === '..') {
            if (!segments.length) return null;
            segments.pop();
        } else {
            segments.push(segment);
        }
    }
    const normalized = segments.join('/');
    return normalized && normalized.startsWith('xl/') ? normalized : null;
}

function resolveRelationshipTarget(sourcePart: string, target: string): string | null {
    if (!target || target.startsWith('//')) return null;
    const base = target.startsWith('/')
        ? target
        : `${sourcePart.slice(0, Math.max(0, sourcePart.lastIndexOf('/') + 1))}${target}`;
    return normalizePartName(base);
}

function relationshipPart(sourcePart: string): string {
    const separator = sourcePart.lastIndexOf('/');
    const directory = separator >= 0 ? sourcePart.slice(0, separator + 1) : '';
    const filename = separator >= 0 ? sourcePart.slice(separator + 1) : sourcePart;
    return `${directory}_rels/${filename}.rels`;
}

async function readXmlPart(
    context: ReaderContext,
    partName: string,
    maximumCharacters: number,
): Promise<XmlNode | null> {
    const normalizedPart = normalizePartName(partName);
    if (!normalizedPart) {
        warn(context, `Partie OOXML rejetée : ${partName.slice(0, 160)}`);
        return null;
    }
    const entry = context.zip.file(normalizedPart);
    if (!entry || entry.dir) return null;
    const declaredSize = entry._data?.uncompressedSize;
    if (declaredSize == null || !Number.isFinite(declaredSize) || declaredSize < 0) {
        warn(context, `Taille OOXML inconnue, partie ignorée : ${normalizedPart}`);
        return null;
    }
    if (
        declaredSize > maximumCharacters
        || context.totalXmlCharacters + context.reservedXmlCharacters + declaredSize
            > context.limits.maxTotalXmlCharacters
    ) {
        warn(context, `Partie OOXML trop volumineuse ignorée : ${normalizedPart}`);
        return null;
    }
    context.reservedXmlCharacters += declaredSize;
    let xml: string;
    try {
        xml = await entry.async('string');
    } catch {
        warn(context, `Partie OOXML illisible ignorée : ${normalizedPart}`);
        return null;
    } finally {
        context.reservedXmlCharacters -= declaredSize;
    }
    if (xml.length > declaredSize
        || xml.length > maximumCharacters
        || context.totalXmlCharacters + xml.length > context.limits.maxTotalXmlCharacters) {
        warn(context, `Budget XML dépassé, partie ignorée : ${normalizedPart}`);
        return null;
    }
    context.totalXmlCharacters += xml.length;
    const parsed = parseXml(xml, context.limits);
    if (!parsed) warn(context, `XML mal formé ou hors limites ignoré : ${normalizedPart}`);
    return parsed;
}

function parseRelationships(
    root: XmlNode | null,
    sourcePart: string,
    context?: ReaderContext,
): Map<string, Relationship> {
    const relationships = new Map<string, Relationship>();
    const ambiguous = new Set<string>();
    for (const node of descendants(root ?? undefined, 'Relationship', 4096)) {
        const id = attribute(node, 'Id') ?? attribute(node, 'id');
        const type = attribute(node, 'Type') ?? attribute(node, 'type');
        const target = attribute(node, 'Target') ?? attribute(node, 'target');
        const external = (attribute(node, 'TargetMode') ?? '').toLowerCase() === 'external';
        if (!id || id.length > 255 || !type || type.length > 1024 || !target || target.length > 1024) continue;
        if (external) {
            if (context && ['/worksheet', '/drawing', '/chart', '/chartex'].some(suffix => type.toLowerCase().endsWith(suffix))) {
                warn(context, `Relation graphique ou feuille externe ignorée dans ${sourcePart}.`);
            }
            continue;
        }
        const resolved = resolveRelationshipTarget(sourcePart, target);
        if (!resolved) {
            if (context) warn(context, `Cible de relation invalide ignorée dans ${sourcePart}.`);
            continue;
        }
        if (relationships.has(id)) {
            ambiguous.add(id);
            relationships.delete(id);
            if (context) warn(context, `Relation dupliquée ${id.slice(0, 80)} ignorée dans ${sourcePart}.`);
        } else if (!ambiguous.has(id)) {
            relationships.set(id, { id, type, target: resolved });
        }
    }
    return relationships;
}

async function readRelationships(context: ReaderContext, sourcePart: string): Promise<Map<string, Relationship>> {
    const part = relationshipPart(sourcePart);
    const root = await readXmlPart(context, part, context.limits.maxXmlCharacters);
    return parseRelationships(root, sourcePart, context);
}

function relationshipKind(type: string): 'drawing' | 'chart' | 'chartEx' | 'other' {
    const normalized = type.toLowerCase();
    if (normalized.endsWith('/drawing')) return 'drawing';
    if (normalized.endsWith('/chartex')) return 'chartEx';
    if (normalized.endsWith('/chart')) return 'chart';
    return 'other';
}

function parseColumnName(value: string): number | null {
    let result = 0;
    for (const character of value.toUpperCase()) {
        const code = character.charCodeAt(0) - 64;
        if (code < 1 || code > 26) return null;
        result = result * 26 + code;
    }
    return result > 0 && result <= 16_384 ? result - 1 : null;
}

function parseA1Range(value: string): ParsedRange | null {
    const normalized = normalizeA1Range(value);
    if (!SIMPLE_A1_RANGE.test(normalized)) return null;
    const [start, end = start] = normalized.split(':');
    const parseCell = (cell: string): { ri: number; ci: number } | null => {
        const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(cell);
        if (!match) return null;
        const ci = parseColumnName(match[1]);
        const ri = Number(match[2]) - 1;
        return ci != null && ri >= 0 && ri < 1_048_576 ? { ri, ci } : null;
    };
    const first = parseCell(start);
    const last = parseCell(end);
    if (!first || !last || last.ri < first.ri || last.ci < first.ci) return null;
    return { ref: normalized, sri: first.ri, sci: first.ci, eri: last.ri, eci: last.ci };
}

function normalizeSheetName(value: string): string | null {
    const trimmed = value.trim();
    if (!trimmed) return null;
    if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
        return trimmed.slice(1, -1).replace(/''/g, "'");
    }
    if (trimmed.startsWith("'") || trimmed.endsWith("'")) return null;
    return trimmed;
}

function normalizeLocalFormula(
    value: string | undefined,
    sheetName: string,
    limits: OoxmlChartReaderLimits,
): ParsedRange | null {
    if (!value) return null;
    let formula = value.trim();
    if (formula.startsWith('=')) formula = formula.slice(1).trim();
    if (!formula || formula.length > limits.maxFormulaCharacters || /[\[\],();{}]/.test(formula)) return null;
    const separator = formula.lastIndexOf('!');
    if (separator >= 0) {
        if (formula.indexOf('!') !== separator) return null;
        const referencedSheet = normalizeSheetName(formula.slice(0, separator));
        if (!referencedSheet || referencedSheet.toLocaleLowerCase() !== sheetName.toLocaleLowerCase()) return null;
        formula = formula.slice(separator + 1);
    }
    return parseA1Range(formula);
}

function rangeFormula(
    container: XmlNode | undefined,
    sheetName: string,
    limits: OoxmlChartReaderLimits,
): ParsedRange | null {
    const formula = firstDescendant(container, 'f');
    return normalizeLocalFormula(formula ? textContent(formula, limits.maxFormulaCharacters + 1) : undefined, sheetName, limits);
}

function chartTypeForGroup(node: XmlNode): number | null {
    const grouping = (readValue(node, 'grouping') ?? 'standard').toLowerCase();
    const stackedOffset = grouping === 'stacked' ? 1 : grouping === 'percentstacked' ? 2 : 0;
    const exploded = directChildren(node, 'ser').some(series => (finiteNumber(readValue(series, 'explosion')) ?? 0) > 0);
    switch (node.localName) {
        case 'barChart': {
            const direction = (readValue(node, 'barDir') ?? 'col').toLowerCase();
            return direction === 'bar' ? 57 + stackedOffset : 51 + stackedOffset;
        }
        case 'bar3DChart': {
            const direction = (readValue(node, 'barDir') ?? 'col').toLowerCase();
            const shape = (readValue(node, 'shape') ?? 'box').toLowerCase();
            const shapedTypes: Record<string, { bar: [number, number, number]; column: [number, number, number, number] }> = {
                cone: { bar: [102, 103, 104], column: [99, 100, 101, 105] },
                conetomax: { bar: [102, 103, 104], column: [99, 100, 101, 105] },
                cylinder: { bar: [95, 96, 97], column: [92, 93, 94, 98] },
                pyramid: { bar: [109, 110, 111], column: [106, 107, 108, 112] },
                pyramidtomax: { bar: [109, 110, 111], column: [106, 107, 108, 112] },
            };
            const shaped = shapedTypes[shape];
            if (shaped) {
                if (direction === 'bar') return shaped.bar[stackedOffset];
                if (grouping === 'standard') return shaped.column[3];
                return shaped.column[stackedOffset];
            }
            if (!['box', ''].includes(shape)) return null;
            if (direction === 'bar') return 60 + stackedOffset;
            if (grouping === 'standard') return -4100;
            return 54 + stackedOffset;
        }
        case 'lineChart': {
            const markers = directChildren(node, 'ser').some(series => {
                const marker = directChild(series, 'marker');
                const symbol = readValue(marker, 'symbol');
                return marker != null && symbol !== 'none';
            });
            if (grouping === 'stacked') return markers ? 66 : 63;
            if (grouping === 'percentstacked') return markers ? 67 : 64;
            return markers ? 65 : 4;
        }
        case 'line3DChart': return -4101;
        case 'areaChart': return stackedOffset === 1 ? 76 : stackedOffset === 2 ? 77 : 1;
        case 'area3DChart': return stackedOffset === 1 ? 78 : stackedOffset === 2 ? 79 : -4098;
        case 'pieChart': return exploded ? 69 : 5;
        case 'pie3DChart': return exploded ? 70 : -4102;
        case 'doughnutChart': return exploded ? 80 : -4120;
        case 'ofPieChart': return (readValue(node, 'ofPieType') ?? '').toLowerCase() === 'bar' ? 71 : 68;
        case 'scatterChart': {
            const style = (readValue(node, 'scatterStyle') ?? 'marker').toLowerCase();
            if (style === 'smooth') return 73;
            if (style === 'smoothmarker') return 72;
            if (style === 'line') return 75;
            if (style === 'linemarker') return 74;
            return -4169;
        }
        case 'bubbleChart': return booleanValue(readValue(node, 'bubble3D')) ? 87 : 15;
        case 'radarChart': {
            const style = (readValue(node, 'radarStyle') ?? 'standard').toLowerCase();
            return style === 'filled' ? 82 : style === 'marker' ? 81 : -4151;
        }
        case 'stockChart': {
            const count = directChildren(node, 'ser').length;
            return count >= 5 ? 91 : count === 4 ? 89 : 88;
        }
        case 'surfaceChart': return booleanValue(readValue(node, 'wireframe')) ? 86 : 85;
        case 'surface3DChart': return booleanValue(readValue(node, 'wireframe')) ? 84 : 83;
        default: return null;
    }
}

function parseMarkerStyle(value: string | undefined): ChartMarkerStyle | undefined {
    const normalized = value === 'auto' ? 'automatic' : value;
    return [
        'automatic', 'circle', 'dash', 'diamond', 'dot', 'none',
        'picture', 'plus', 'square', 'star', 'triangle', 'x',
    ].includes(normalized ?? '') ? normalized as ChartMarkerStyle : undefined;
}

function parseColor(node: XmlNode | undefined): string | undefined {
    const solidFill = firstDescendant(node, 'solidFill');
    const rgb = attribute(firstDescendant(solidFill, 'srgbClr'), 'val');
    return rgb && /^[0-9a-f]{6}$/i.test(rgb) ? `#${rgb.toUpperCase()}` : undefined;
}

function parseDashStyle(value: string | undefined): SheetChartSeriesData['dashStyle'] | undefined {
    switch ((value ?? '').toLowerCase()) {
        case 'dash':
        case 'lgdash': return 'dash';
        case 'dot':
        case 'sysdot': return 'dot';
        case 'dashdot':
        case 'lgdashdot':
        case 'sysdashdot': return 'dashDot';
        case 'solid': return 'solid';
        default: return undefined;
    }
}

function dataLabelsForGroup(group: ChartGroup): SheetChartSeriesData['dataLabels'] | undefined {
    const labels = directChild(group.node, 'dLbls');
    if (!labels) return undefined;
    const positionMap: Record<string, NonNullable<SheetChartSeriesData['dataLabels']>['position']> = {
        t: 'above', b: 'below', bestFit: 'bestFit', ctr: 'center', inBase: 'insideBase',
        inEnd: 'insideEnd', l: 'left', outEnd: 'outsideEnd', r: 'right',
    };
    const positionValue = readValue(labels, 'dLblPos');
	const position = positionValue ? positionMap[positionValue] : undefined;
    const result: NonNullable<SheetChartSeriesData['dataLabels']> = {
		...(position && chartSeriesSupportsDataLabelPosition(group.chartType, position)
			? { position }
			: {}),
    };
    for (const [property, element] of [
        ['showValue', 'showVal'],
        ['showCategoryName', 'showCatName'],
        ['showSeriesName', 'showSerName'],
    ] as const) {
        const rawValue = readValue(labels, element);
        if (rawValue !== undefined) result[property] = booleanValue(rawValue);
    }
    const percentageValue = readValue(labels, 'showPercent');
    if (
        percentageValue !== undefined
        && chartSeriesSupportsPercentageDataLabels(group.chartType)
    ) result.showPercentage = booleanValue(percentageValue);
    const bubbleSizeValue = readValue(labels, 'showBubbleSize');
    if (
        bubbleSizeValue !== undefined
        && chartSeriesSupportsBubbleSizes(group.chartType)
    ) result.showBubbleSize = booleanValue(bubbleSizeValue);
	if (result.position !== undefined && !chartDataLabelsHaveEnabledShowOption(result)) {
		delete result.position;
	}
    // A position without a show flag is not a complete native-write request.
    // Excel's ApplyDataLabels() would otherwise turn ShowValue on implicitly.
    return chartDataLabelsHaveExplicitShowOption(result) ? result : undefined;
}

function cachedSeriesName(transaction: XmlNode | undefined): string | undefined {
    if (!transaction) return undefined;
    const direct = directChild(transaction, 'v');
    if (direct) return textContent(direct, 255) || undefined;
    const cache = firstDescendant(transaction, 'strCache');
    const point = directChildren(cache, 'pt')
        .sort((left, right) => (finiteNumber(attribute(left, 'idx')) ?? 0) - (finiteNumber(attribute(right, 'idx')) ?? 0))[0];
    const cached = directChild(point, 'v');
    return cached ? textContent(cached, 255) || undefined : undefined;
}

function parseSeries(
    node: XmlNode,
    group: ChartGroup,
    index: number,
    chartId: string,
    sheetName: string,
    limits: OoxmlChartReaderLimits,
): ParsedSeries | null {
    const transaction = directChild(node, 'tx');
    const nameRange = rangeFormula(transaction, sheetName, limits);
    const categoryContainer = directChild(node, 'cat');
    const xContainer = directChild(node, 'xVal');
    const valueContainer = directChild(node, 'val') ?? directChild(node, 'yVal');
    const bubbleContainer = directChild(node, 'bubbleSize');
    const categoryRange = rangeFormula(categoryContainer, sheetName, limits);
    const xValuesRange = rangeFormula(xContainer, sheetName, limits);
    const valuesRange = rangeFormula(valueContainer, sheetName, limits);
    const bubbleSizesRange = rangeFormula(bubbleContainer, sheetName, limits);
    if (!valuesRange) return null;

    const shape = directChild(node, 'spPr');
    const line = directChild(shape, 'ln');
    const marker = directChild(node, 'marker');
    const markerSize = boundedInteger(readValue(marker, 'size'), 2, 72);
    const lineWidthEmus = finiteNumber(attribute(line, 'w'));
    // A strRef cache is only a display fallback for the referenced name cell.
    // Emitting both fields would violate the native writer's name/nameRange
    // exclusivity and turn a harmless chart update into an invalid operation.
    const name = nameRange ? undefined : cachedSeriesName(transaction);
    const dataLabels = dataLabelsForGroup(group);
    const smoothValue = readValue(node, 'smooth');
    return {
        id: `series:ooxml:${stableHash(`${chartId}|${index}|${valuesRange.ref}`)}`,
        ...(name ? { name } : {}),
        ...(nameRange ? { nameRange: nameRange.ref, _nameRange: nameRange } : {}),
        ...(categoryRange ? { categoryRange: categoryRange.ref, _categoryRange: categoryRange } : {}),
        valuesRange: valuesRange.ref,
        _valueRange: valuesRange,
        ...(xValuesRange ? { xValuesRange: xValuesRange.ref } : {}),
        ...(bubbleSizesRange ? { bubbleSizesRange: bubbleSizesRange.ref } : {}),
        chartType: group.chartType,
        axisGroup: group.secondary ? 'secondary' : 'primary',
        ...(parseColor(shape) ? { color: parseColor(shape) } : {}),
        ...(parseColor(line) ? { lineColor: parseColor(line) } : {}),
        ...(lineWidthEmus != null && lineWidthEmus >= 0 && lineWidthEmus <= 20 * EMUS_PER_POINT
            ? { lineWidth: lineWidthEmus / EMUS_PER_POINT }
            : {}),
        ...(parseDashStyle(readValue(line, 'prstDash')) ? { dashStyle: parseDashStyle(readValue(line, 'prstDash')) } : {}),
        ...(parseMarkerStyle(readValue(marker, 'symbol')) ? { markerStyle: parseMarkerStyle(readValue(marker, 'symbol')) } : {}),
        ...(markerSize != null ? { markerSize } : {}),
        ...(smoothValue !== undefined && chartSeriesSupportsSmooth(group.chartType)
            ? { smooth: booleanValue(smoothValue) }
            : {}),
        visible: !booleanValue(readValue(node, 'delete')),
        ...(dataLabels ? { dataLabels } : {}),
    };
}

function inferPlotBy(chart: XmlNode, series: readonly ParsedSeries[]): 'columns' | 'rows' {
    const explicit = attribute(firstDescendant(chart, 'plotBy'), 'val')?.toLowerCase();
    if (explicit === 'row' || explicit === 'rows') return 'rows';
    if (explicit === 'col' || explicit === 'columns') return 'columns';
    const valueRanges = series.map(item => item._valueRange).filter((range): range is ParsedRange => range != null);
    return valueRanges.length && valueRanges.every(range => range.sri === range.eri && range.sci !== range.eci)
        ? 'rows'
        : 'columns';
}

function chartTitle(node: XmlNode | undefined): string | undefined {
    if (!node) return undefined;
    const textRuns = descendants(node, 't', 128).map(text => textContent(text, 1024)).filter(Boolean);
    if (textRuns.length) return textRuns.join('').slice(0, 255);
    const cachedValues = descendants(firstDescendant(node, 'strCache'), 'v', 16)
        .map(value => textContent(value, 1024)).filter(Boolean);
    return cachedValues.join('').slice(0, 255) || undefined;
}

function parseAxis(node: XmlNode): SheetChartAxisData {
    const scaling = directChild(node, 'scaling');
    const minimum = finiteNumber(readValue(scaling, 'min'));
    const maximum = finiteNumber(readValue(scaling, 'max'));
    const majorUnit = finiteNumber(readValue(node, 'majorUnit'));
    const minorUnit = finiteNumber(readValue(node, 'minorUnit'));
    const logarithmic = finiteNumber(readValue(scaling, 'logBase')) != null;
    const orientation = (readValue(scaling, 'orientation') ?? '').toLowerCase();
    const numberFormatNode = directChild(node, 'numFmt');
    const numberFormat = booleanValue(attribute(numberFormatNode, 'sourceLinked'))
        ? ''
        : attribute(numberFormatNode, 'formatCode');
    const title = chartTitle(directChild(node, 'title'));
    return {
        visible: !booleanValue(readValue(node, 'delete')),
        ...(title ? { title } : {}),
        minimumScale: minimum ?? null,
        maximumScale: maximum ?? null,
        majorUnit: majorUnit != null && majorUnit > 0 ? majorUnit : null,
        minorUnit: minorUnit != null && minorUnit > 0 ? minorUnit : null,
        logarithmic,
        reverseOrder: orientation === 'maxmin',
        ...(numberFormat !== undefined && numberFormat.length <= 255 ? { numberFormat } : {}),
        majorGridlines: directChild(node, 'majorGridlines') != null,
        minorGridlines: directChild(node, 'minorGridlines') != null,
    };
}

function parseAxes(
    plotArea: XmlNode,
    primaryAxisIds: ReadonlySet<string>,
    secondaryAxisIds: ReadonlySet<string>,
    hasScatterAxes: boolean,
): Pick<SheetChartData, 'categoryAxis' | 'valueAxis' | 'secondaryCategoryAxis' | 'secondaryValueAxis'> {
    const result: Pick<SheetChartData, 'categoryAxis' | 'valueAxis' | 'secondaryCategoryAxis' | 'secondaryValueAxis'> = {};
    for (const node of plotArea.children.filter(child => ['catAx', 'dateAx', 'valAx', 'serAx'].includes(child.localName))) {
        const id = readValue(node, 'axId');
        if (!id || (!primaryAxisIds.has(id) && !secondaryAxisIds.has(id))) continue;
        const secondary = secondaryAxisIds.has(id) && !primaryAxisIds.has(id);
        const position = (readValue(node, 'axPos') ?? '').toLowerCase();
        const category = node.localName !== 'valAx' || (hasScatterAxes && (position === 'b' || position === 't'));
        const key = secondary
            ? category ? 'secondaryCategoryAxis' : 'secondaryValueAxis'
            : category ? 'categoryAxis' : 'valueAxis';
        if (!result[key]) result[key] = parseAxis(node);
    }
    return result;
}

function parseLegend(chart: XmlNode): SheetChartData['legend'] | undefined {
    const legend = directChild(chart, 'legend');
    if (!legend) return { visible: false, position: 'right' };
    const positions: Record<string, ChartLegendPosition> = {
        b: 'bottom', l: 'left', r: 'right', t: 'top', tr: 'corner',
    };
    const position = descendants(legend, 'manualLayout', 64).length > 0
        ? 'custom'
        : positions[readValue(legend, 'legendPos') ?? 'r'] ?? 'right';
    return { visible: !booleanValue(readValue(legend, 'delete')), position };
}

function parseClassicChart(
    root: XmlNode,
    metadata: {
        chartId: string;
        name: string;
        alternativeText?: string;
        anchor: SheetChartData['anchor'];
        sheetName: string;
    },
    context: ReaderContext,
): SheetChartData | null {
    if (root.localName !== 'chartSpace') return null;
    if (descendants(root, 'pivotSource', 1).length > 0) {
        warn(context, `PivotChart non éditable conservé nativement : ${metadata.name}.`);
        return null;
    }
    const chart = directChild(root, 'chart') ?? firstDescendant(root, 'chart');
    const plotArea = directChild(chart, 'plotArea');
    if (!chart || !plotArea) return null;
    const candidateGroups = plotArea.children.filter(node => (
        node.localName.endsWith('Chart') && directChildren(node, 'ser').length > 0
    ));
    if (!candidateGroups.length || candidateGroups.some(node => chartTypeForGroup(node) == null)) return null;
    const groups: ChartGroup[] = candidateGroups
        .map(node => ({ node, chartType: chartTypeForGroup(node) }))
        .filter((item): item is { node: XmlNode; chartType: number } => item.chartType != null)
        .map(item => ({
            ...item,
            axisIds: directChildren(item.node, 'axId')
                .map(axis => attribute(axis, 'val') ?? textContent(axis, 64))
                .filter((id): id is string => Boolean(id && id.length <= 64)),
            secondary: false,
        }));
    if (!groups.length) return null;
    const firstAxisGroup = groups.find(group => group.axisIds.length);
    const primaryAxisIds = new Set(firstAxisGroup?.axisIds ?? []);
    const secondaryAxisIds = new Set<string>();
    groups.forEach(group => {
        group.secondary = primaryAxisIds.size > 0 && group.axisIds.some(id => !primaryAxisIds.has(id));
        if (group.secondary) group.axisIds.forEach(id => secondaryAxisIds.add(id));
    });

    const parsedSeries: ParsedSeries[] = [];
    for (const group of groups) {
        for (const seriesNode of directChildren(group.node, 'ser')) {
            if (parsedSeries.length >= context.limits.maxSeriesPerChart) {
                warn(context, `Graphique non éditable : trop de séries dans ${metadata.name}.`);
                return null;
            }
            const parsed = parseSeries(
                seriesNode,
                group,
                parsedSeries.length,
                metadata.chartId,
                metadata.sheetName,
                context.limits,
            );
            if (!parsed) {
                warn(context, `Graphique non éditable : série OOXML illisible dans ${metadata.name}.`);
                return null;
            }
            parsedSeries.push(parsed);
        }
    }
    if (!parsedSeries.length) return null;
    const distinctTypes = new Set(groups.map(group => group.chartType));
    const stockGroup = groups.find(group => group.node.localName === 'stockChart');
    const volumeGroup = groups.find(group => group.node.localName === 'barChart');
    const stockSeriesCount = stockGroup ? directChildren(stockGroup.node, 'ser').length : 0;
    const chartType = stockGroup && volumeGroup && groups.length === 2
        ? stockSeriesCount >= 4 ? 91 : 90
        : distinctTypes.size > 1 ? -4152 : groups[0].chartType;
    const titleText = chartTitle(directChild(chart, 'title'));
    const style = boundedInteger(attribute(directChild(root, 'style'), 'val'), 1, 48);
    const gapWidth = boundedInteger(readValue(groups[0].node, 'gapWidth'), 0, 500);
    const overlap = boundedInteger(readValue(groups[0].node, 'overlap'), -100, 100);
    const axes = parseAxes(
        plotArea,
        primaryAxisIds,
        secondaryAxisIds,
        groups.some(group => ['scatterChart', 'bubbleChart'].includes(group.node.localName)),
    );
    const cleanSeries = parsedSeries.map(({ _valueRange, _categoryRange, _nameRange, ...series }) => series);
    return {
        id: metadata.chartId,
        name: metadata.name,
        chartType,
        plotBy: inferPlotBy(chart, parsedSeries),
        anchor: metadata.anchor,
        title: { visible: Boolean(titleText), text: titleText ?? '' },
        legend: parseLegend(chart),
        ...axes,
        series: cleanSeries,
        ...(style != null ? { style } : {}),
        ...(gapWidth != null ? { gapWidth } : {}),
        ...(overlap != null ? { overlap } : {}),
        ...(metadata.alternativeText ? { alternativeText: metadata.alternativeText } : {}),
    };
}

function anchorCoordinate(node: XmlNode | undefined): { x: number; y: number } | null {
    if (!node) return null;
    const column = boundedInteger(readValue(node, 'col'), 0, 16_383);
    const row = boundedInteger(readValue(node, 'row'), 0, 1_048_575);
    const columnOffset = boundedInteger(readValue(node, 'colOff'), 0, 2_000_000_000) ?? 0;
    const rowOffset = boundedInteger(readValue(node, 'rowOff'), 0, 2_000_000_000) ?? 0;
    if (column == null || row == null) return null;
    const x = column * DEFAULT_COLUMN_POINTS + columnOffset / EMUS_PER_POINT;
    const y = row * DEFAULT_ROW_POINTS + rowOffset / EMUS_PER_POINT;
    return Number.isFinite(x) && Number.isFinite(y) && x <= MAX_COORDINATE_POINTS && y <= MAX_COORDINATE_POINTS
        ? { x, y }
        : null;
}

function extent(node: XmlNode | undefined): { width: number; height: number } | null {
    if (!node) return null;
    const cx = finiteNumber(attribute(node, 'cx'));
    const cy = finiteNumber(attribute(node, 'cy'));
    if (cx == null || cy == null || cx <= 0 || cy <= 0) return null;
    const width = cx / EMUS_PER_POINT;
    const height = cy / EMUS_PER_POINT;
    return width > 0 && height > 0 && width <= 100_000 && height <= 100_000 ? { width, height } : null;
}

function parseAnchor(node: XmlNode): SheetChartData['anchor'] | null {
    if (node.localName === 'twoCellAnchor') {
        const from = anchorCoordinate(directChild(node, 'from'));
        const to = anchorCoordinate(directChild(node, 'to'));
        if (!from || !to || to.x <= from.x || to.y <= from.y) return null;
        return { left: from.x, top: from.y, width: to.x - from.x, height: to.y - from.y };
    }
    if (node.localName === 'oneCellAnchor') {
        const from = anchorCoordinate(directChild(node, 'from'));
        const size = extent(directChild(node, 'ext'));
        return from && size ? { left: from.x, top: from.y, ...size } : null;
    }
    if (node.localName === 'absoluteAnchor') {
        const position = directChild(node, 'pos');
        const x = finiteNumber(attribute(position, 'x'));
        const y = finiteNumber(attribute(position, 'y'));
        const size = extent(directChild(node, 'ext'));
        if (x == null || y == null || x < 0 || y < 0 || !size) return null;
        const left = x / EMUS_PER_POINT;
        const top = y / EMUS_PER_POINT;
        return left <= MAX_COORDINATE_POINTS && top <= MAX_COORDINATE_POINTS
            ? { left, top, ...size }
            : null;
    }
    return null;
}

interface DrawingInventory {
    charts: SheetChartData[];
    hasChartParts: boolean;
    unsupportedChartCount: number;
}

async function readDrawingCharts(
    context: ReaderContext,
    sheetName: string,
    worksheetPart: string,
    drawingPart: string,
): Promise<DrawingInventory> {
    const result: DrawingInventory = { charts: [], hasChartParts: false, unsupportedChartCount: 0 };
    const [drawing, relationships] = await Promise.all([
        readXmlPart(context, drawingPart, context.limits.maxXmlCharacters),
        readRelationships(context, drawingPart),
    ]);
    const chartRelationships = [...relationships.values()].filter(relation => {
        const kind = relationshipKind(relation.type);
        return kind === 'chart' || kind === 'chartEx';
    });
    if (chartRelationships.length) result.hasChartParts = true;
    if (!drawing) {
        result.unsupportedChartCount = chartRelationships.length;
        return result;
    }
    const anchors = drawing.children.filter(child => (
        child.localName === 'twoCellAnchor'
        || child.localName === 'oneCellAnchor'
        || child.localName === 'absoluteAnchor'
    ));
    if (anchors.length > context.limits.maxAnchorsPerDrawing) {
        warn(context, `Ancres supplémentaires ignorées dans ${drawingPart}.`);
    }
    const handledRelationships = new Set<string>();
    for (const [anchorIndex, anchorNode] of anchors.slice(0, context.limits.maxAnchorsPerDrawing).entries()) {
        const chartNodes = descendants(anchorNode, 'chart', 2);
        if (chartNodes.length !== 1) continue;
        const relationshipId = attribute(chartNodes[0], 'id');
        const relationship = relationshipId ? relationships.get(relationshipId) : undefined;
        if (!relationship) continue;
        const kind = relationshipKind(relationship.type);
        if (kind !== 'chart' && kind !== 'chartEx') continue;
        handledRelationships.add(relationship.id);
        result.hasChartParts = true;
        if (context.chartAttempts >= context.limits.maxCharts) {
            result.unsupportedChartCount += 1;
            warn(context, `Limite globale de tentatives de graphiques atteinte dans ${drawingPart}.`);
            continue;
        }
        context.chartAttempts += 1;
        if (kind === 'chartEx') {
            result.unsupportedChartCount += 1;
            warn(context, `Graphique chartEx détecté mais non hydraté : ${relationship.target}.`);
            continue;
        }
        const anchor = parseAnchor(anchorNode);
        if (!anchor) {
            result.unsupportedChartCount += 1;
            warn(context, `Ancre de graphique invalide ignorée dans ${drawingPart}.`);
            continue;
        }
        const nonVisual = firstDescendant(anchorNode, 'cNvPr');
        const objectId = attribute(nonVisual, 'id') ?? String(anchorIndex + 1);
        const rawName = attribute(nonVisual, 'name')?.trim();
        const name = rawName && rawName.length <= 255 ? rawName : `Chart ${anchorIndex + 1}`;
        const rawAlternativeText = attribute(nonVisual, 'descr') ?? attribute(nonVisual, 'title');
        const alternativeText = rawAlternativeText?.trim().slice(0, 1000) || undefined;
        const chartId = `chart:ooxml:${stableHash(`${worksheetPart}|${drawingPart}|${relationship.target}|${objectId}`)}`;
        const chartRoot = await readXmlPart(context, relationship.target, context.limits.maxChartXmlCharacters);
        const chart = chartRoot ? parseClassicChart(chartRoot, {
            chartId,
            name,
            alternativeText,
            anchor,
            sheetName,
        }, context) : null;
        if (!chart) {
            result.unsupportedChartCount += 1;
            warn(context, `Graphique classique non reconnu ou ambigu ignoré : ${relationship.target}.`);
            continue;
        }
        context.chartCount += 1;
        result.charts.push(chart);
    }
    for (const relationship of chartRelationships) {
        if (handledRelationships.has(relationship.id)) continue;
        result.unsupportedChartCount += 1;
        const kind = relationshipKind(relationship.type);
        warn(context, `${kind === 'chartEx' ? 'Graphique chartEx' : 'Partie graphique'} sans ancre exploitable : ${relationship.target}.`);
    }
    return result;
}

async function discoverWorksheetParts(context: ReaderContext): Promise<OoxmlWorksheetPart[]> {
    const workbookPart = 'xl/workbook.xml';
    const [workbook, relationships] = await Promise.all([
        readXmlPart(context, workbookPart, context.limits.maxXmlCharacters),
        readRelationships(context, workbookPart),
    ]);
    const sheetsContainer = firstDescendant(workbook ?? undefined, 'sheets');
    if (!sheetsContainer) return [];
    const result: OoxmlWorksheetPart[] = [];
    const usedNames = new Set<string>();
    const usedRelationshipIds = new Set<string>();
    for (const sheet of directChildren(sheetsContainer, 'sheet')) {
        if (result.length >= context.limits.maxWorksheets) {
            warn(context, 'Feuilles supplémentaires ignorées : limite atteinte.');
            break;
        }
        const sheetName = attribute(sheet, 'name')?.trim();
        const relationshipId = attribute(sheet, 'id');
        if (!sheetName || sheetName.length > 31 || !relationshipId) continue;
        const normalizedName = sheetName.toLocaleLowerCase();
        if (usedNames.has(normalizedName) || usedRelationshipIds.has(relationshipId)) {
            warn(context, `Feuille ou relation dupliquée ignorée : ${sheetName.slice(0, 80)}.`);
            continue;
        }
        const relationship = relationships.get(relationshipId);
        if (!relationship || !relationship.type.toLowerCase().endsWith('/worksheet')) {
            warn(context, `Relation de feuille absente ou ambiguë : ${sheetName.slice(0, 80)}.`);
            continue;
        }
        usedNames.add(normalizedName);
        usedRelationshipIds.add(relationshipId);
        result.push({ sheetName, worksheetPart: relationship.target });
    }
    return result;
}

function normalizeWorksheetMapping(
    worksheets: readonly OoxmlWorksheetPart[],
    context: ReaderContext,
): OoxmlWorksheetPart[] {
    const result: OoxmlWorksheetPart[] = [];
    const usedNames = new Set<string>();
    const usedParts = new Set<string>();
    for (const worksheet of worksheets.slice(0, context.limits.maxWorksheets)) {
        const sheetName = worksheet.sheetName?.trim();
        const worksheetPart = normalizePartName(worksheet.worksheetPart);
        if (!sheetName || sheetName.length > 31 || !worksheetPart || !worksheetPart.startsWith('xl/worksheets/')) {
            warn(context, 'Mapping feuille/partie invalide ignoré.');
            continue;
        }
        const normalizedName = sheetName.toLocaleLowerCase();
        if (usedNames.has(normalizedName) || usedParts.has(worksheetPart)) {
            warn(context, `Mapping feuille/partie dupliqué ignoré : ${sheetName.slice(0, 80)}.`);
            continue;
        }
        usedNames.add(normalizedName);
        usedParts.add(worksheetPart);
        result.push({ sheetName, worksheetPart });
    }
    return result;
}

export async function readOoxmlChartsForWorksheet(
    zip: OoxmlZipLike,
    worksheet: OoxmlWorksheetPart,
    limitOverrides: Partial<OoxmlChartReaderLimits> = {},
): Promise<OoxmlSheetChartInventory> {
    const result = await readOoxmlChartInventory(zip, [worksheet], limitOverrides);
    return result.sheets[0] ?? {
        sheetName: worksheet.sheetName,
        worksheetPart: worksheet.worksheetPart,
        charts: [],
        hasChartParts: false,
        unsupportedChartCount: 0,
    };
}

export async function readOoxmlChartInventory(
    zip: OoxmlZipLike,
    worksheets?: readonly OoxmlWorksheetPart[],
    limitOverrides: Partial<OoxmlChartReaderLimits> = {},
): Promise<OoxmlChartInventoryResult> {
    const limits: OoxmlChartReaderLimits = {
        ...DEFAULT_OOXML_CHART_READER_LIMITS,
        ...Object.fromEntries(Object.entries(limitOverrides).filter(([, value]) => (
            typeof value === 'number' && Number.isInteger(value) && value > 0
        ))),
    };
    const context: ReaderContext = {
        zip,
        limits,
        warnings: [],
        totalXmlCharacters: 0,
        reservedXmlCharacters: 0,
        chartCount: 0,
        chartAttempts: 0,
    };
    const worksheetParts = worksheets
        ? normalizeWorksheetMapping(worksheets, context)
        : await discoverWorksheetParts(context);
    const sheets: OoxmlSheetChartInventory[] = [];
    for (const worksheet of worksheetParts) {
        const worksheetRoot = await readXmlPart(context, worksheet.worksheetPart, limits.maxXmlCharacters);
        const relationships = await readRelationships(context, worksheet.worksheetPart);
        const drawingIds = descendants(worksheetRoot ?? undefined, 'drawing', limits.maxDrawingsPerWorksheet + 1)
            .map(node => attribute(node, 'id'))
            .filter((id): id is string => Boolean(id));
        const orderedDrawingRelationships: Relationship[] = [];
        const queuedRelationshipIds = new Set<string>();
        drawingIds.forEach(id => {
            const relationship = relationships.get(id);
            if (relationship && relationshipKind(relationship.type) === 'drawing' && !queuedRelationshipIds.has(relationship.id)) {
                orderedDrawingRelationships.push(relationship);
                queuedRelationshipIds.add(relationship.id);
            }
        });
        [...relationships.values()].forEach(relationship => {
            if (relationshipKind(relationship.type) === 'drawing' && !queuedRelationshipIds.has(relationship.id)) {
                orderedDrawingRelationships.push(relationship);
                queuedRelationshipIds.add(relationship.id);
                warn(context, `Drawing non référencé explicitement inspecté par sécurité : ${relationship.target}.`);
            }
        });
        if (orderedDrawingRelationships.length > limits.maxDrawingsPerWorksheet) {
            warn(context, `Drawings supplémentaires ignorés dans ${worksheet.worksheetPart}.`);
        }
        const sheetResult: OoxmlSheetChartInventory = {
            ...worksheet,
            charts: [],
            hasChartParts: false,
            unsupportedChartCount: 0,
        };
        const usedDrawings = new Set<string>();
        for (const relationship of orderedDrawingRelationships.slice(0, limits.maxDrawingsPerWorksheet)) {
            if (usedDrawings.has(relationship.target)) continue;
            usedDrawings.add(relationship.target);
            const drawing = await readDrawingCharts(
                context,
                worksheet.sheetName,
                worksheet.worksheetPart,
                relationship.target,
            );
            sheetResult.charts.push(...drawing.charts);
            sheetResult.hasChartParts ||= drawing.hasChartParts;
            sheetResult.unsupportedChartCount += drawing.unsupportedChartCount;
        }
        sheets.push(sheetResult);
    }
    return { sheets, warnings: context.warnings };
}
